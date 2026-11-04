import json
import os
import re
import sys
import time
import threading
import traceback

import vk_api
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from vk_api.bot_longpoll import VkBotEventType, VkBotLongPoll
from vk_api.upload import VkUpload
from vk_api.utils import get_random_id

from database import (
    STATUS_DELIVERED,
    add_to_cart, checkout_cart, clear_cart, get_all_products, get_cart, get_order,
    get_order_items, get_order_total, get_orders_for_export, get_product,
    get_product_by_vk_id, get_product_sizes, init_db, load_all_states, remove_cart_item, save_state,
    set_status, update_cart_quantity, update_cart_size, upsert_market_items,
)
from keyboards import (
    get_cart_edit_keyboard, get_cart_keyboard, get_item_input_keyboard,
    get_main_keyboard, get_market_confirm_keyboard,
    get_payment_confirmation_keyboard, get_payment_keyboard, get_size_keyboard,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ------------------------------------------------------------------ настройки

def load_env(path):
    """Читает файл .env (строки вида КЛЮЧ=значение) в переменные окружения."""
    if not os.path.exists(path):
        return
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            key, value = line.split('=', 1)
            os.environ.setdefault(key.strip(), value.strip().strip('"\''))


def require_env(name):
    value = os.environ.get(name)
    if not value:
        sys.exit(f'Не задана переменная {name}. Создайте файл .env рядом с main.py '
                 f'(пример — в .env.example).')
    return value


load_env(os.path.join(BASE_DIR, '.env'))

GROUP_TOKEN = require_env('GROUP_TOKEN')   # ключ сообщества (сообщения)
USER_TOKEN = require_env('USER_TOKEN')     # пользовательский токен (VK Market)
GROUP_ID = 242321155
ADMIN_ID = 588644580

MAX_QTY = 99          # максимум штук одной позиции за раз
MAX_SIZE_LEN = 20     # максимальная длина размера
CATALOG_SYNC_INTERVAL = 30 * 60   # как часто обновлять каталог из VK Market (секунды)

ORDERS_XLSX = os.path.join(BASE_DIR, 'orders.xlsx')
CATALOG_XLSX = os.path.join(BASE_DIR, 'catalog.xlsx')
QR_PATH = os.path.join(BASE_DIR, 'qr-code.png')

auth = vk_api.VkApi(token=GROUP_TOKEN)          # для сообщений
user_auth = vk_api.VkApi(token=USER_TOKEN)      # для Маркета
upload = VkUpload(auth)
longpoll = VkBotLongPoll(auth, group_id=GROUP_ID)

REQUISITES_TEXT = (
    "💳 Оплата по реквизитам:\n"
    "Банк: Т-Банк\n"
    "Номер карты: 2200 7011 2345 6789\n"
    "Получатель: Иван И.\n\n"
    "После оплаты нажмите кнопку «Я оплатил(а)» и пришлите чек."
)

SBP_TEXT = (
    "📱 Оплата через СБП:\n"
    "Отсканируйте QR-код в приложении вашего банка.\n\n"
    "После оплаты нажмите кнопку «Я оплатил(а)» и пришлите чек."
)

# Состояние диалогов (хранится в памяти, при перезапуске бота сбрасывается)
user_screen = {}      # user_id -> название экрана
temp_item = {}        # user_id -> данные товара, который сейчас вводится
pending_market = {}   # user_id -> товар из VK Market, ожидающий «В корзину»

# Клавиатура для ответа «не понял», чтобы пользователь видел кнопки своего экрана
SCREEN_KEYBOARDS = {
    'cart': get_cart_keyboard,
    'payment_choice': get_payment_keyboard,
    'payment_waiting_receipt': get_payment_confirmation_keyboard,
}

# Подсказки для сообщений, которые не подходят текущему шагу диалога.
INVALID_ACTION_MESSAGES = {
    'main': 'Не удалось распознать команду. Выберите «Корзина» или «Оплатить Заказ» либо нажмите «Начать».',
    'cart': 'Такого действия в корзине нет. Используйте кнопки «Добавить товар», «Изменить корзину», «Очистить корзину» или «Назад».',
    'cart_edit': 'Не удалось распознать действие. Выберите удаление товара, изменение количества или размера либо нажмите «Назад».',
    'cart_article': 'Сейчас нужно ввести артикул товара из каталога. Либо нажмите «Отмена».',
    'cart_size': f'Сейчас нужно ввести размер товара (не более {MAX_SIZE_LEN} символов). Либо нажмите «Отмена».',
    'cart_quantity': f'Сейчас нужно ввести целое количество от 1 до {MAX_QTY}. Либо нажмите «Отмена».',
    'cart_remove': 'Сейчас нужно указать номер товара из списка корзины. Введите номер или нажмите «Отмена».',
    'cart_change_num': 'Сейчас нужно указать номер товара, количество которого хотите изменить. Введите номер из списка или нажмите «Отмена».',
    'cart_change_qty': f'Сейчас нужно указать новое количество — целое число от 1 до {MAX_QTY}. Либо нажмите «Отмена».',
    'cart_resize_num': 'Сейчас нужно указать номер товара, у которого есть размер. Введите номер из списка или нажмите «Отмена».',
    'cart_resize_new': f'Сейчас нужно написать новый размер (не более {MAX_SIZE_LEN} символов). Либо нажмите «Отмена».',
    'payment_choice': 'Выберите один из предложенных способов оплаты: «Оплатить по реквизитам» или «Оплатить через СБП». Для выхода нажмите «Назад в меню».',
    'payment_waiting_receipt': 'Ожидаю подтверждение оплаты. Нажмите «Я оплатил(а)», если уже оплатили, или «Отмена», чтобы вернуться в меню.',
    'market_confirm': 'Чтобы продолжить, нажмите «В корзину» под карточкой товара или «Отмена».',
}


def cart_has_sized_items(user_id):
    """True, если в корзине есть хотя бы один товар с размером."""
    return any(size and size != '-' for _, size, _ in get_cart(user_id))


def cart_edit_keyboard(user_id):
    """Меню «Изменить корзину»; кнопка «Изменить размер» только если есть товары с размером."""
    return get_cart_edit_keyboard(show_change_size=cart_has_sized_items(user_id))


def keyboard_for_screen(user_id, screen):
    if screen == 'cart_edit':
        return cart_edit_keyboard(user_id)
    if screen in ('cart_size', 'cart_resize_new'):
        sizes = (temp_item.get(user_id) or {}).get('sizes')
        return get_size_keyboard(sizes) if sizes else get_item_input_keyboard()
    if screen in ('cart_article', 'cart_quantity', 'cart_remove',
                  'cart_change_num', 'cart_change_qty', 'cart_resize_num'):
        return get_item_input_keyboard()
    if screen == 'market_confirm':
        pending = pending_market.get(user_id)
        if pending:
            return get_market_confirm_keyboard(pending['article'])
    return SCREEN_KEYBOARDS.get(screen, get_main_keyboard)()


def send_invalid_action(sender, screen, attachments=None):
    # Объясняем ожидаемое действие и оставляем пользователя на текущем шаге.
    message = INVALID_ACTION_MESSAGES.get(
        screen,
        'Не удалось распознать действие. Нажмите «Начать», чтобы открыть главное меню.'
    )
    if attachments:
        message += '\n\nПохоже, вы отправили вложение. На этом шаге нужно выполнить действие, указанное выше.'
    write_massage(sender, message, keyboard_for_screen(sender, screen))


def go_main_menu(sender, text='Главное меню:'):
    """Сброс состояния и возврат в главное меню (корзина в базе остаётся)."""
    reset_state(sender)
    write_massage(sender, text, get_main_keyboard())


# ------------------------------------------------------------------ каталог

NO_SIZE_WORDS = {'-', '—', '–', 'нет', 'б/р', 'без размера', 'one size', 'onesize'}


def parse_sizes(description):
    """Размеры из описания товара: строка вида «Размер: 41, 43, 45» -> ['41', '43', '45']."""
    m = re.search(r'^[ \t]*Размеры?[ \t]*:[ \t]*(.+)$', description or '', re.IGNORECASE | re.MULTILINE)
    if not m:
        return []
    sizes = []
    for part in re.split(r'[,;]', m.group(1)):
        part = part.strip()
        if (part and len(part) <= MAX_SIZE_LEN and part not in sizes
                and part.casefold() not in NO_SIZE_WORDS):
            sizes.append(part)
    return sizes


def match_size(sizes, text):
    """Возвращает размер из списка, совпадающий с текстом (без учёта регистра), либо None."""
    for size in sizes:
        if size.casefold() == text.casefold():
            return size
    return None


def market_item_record(item):
    """Товар из VK Market (market.get / market.getById) -> запись для таблицы products.
    Артикул = только SKU из карточки VK (или «Артикул: …» в описании).
    Если артикула нет — article = None. id товара VK артикулом не считается."""
    sku = str(item.get('sku') or '').strip()
    if sku.lower() == 'none':
        sku = ''
    if not sku:
        # запасной вариант: «Артикул: 546» в тексте описания
        found = re.search(r'Артикул[ \t]*[:№-]?[ \t]*([\w./-]+)', str(item.get('description') or ''),
                          re.IGNORECASE)
        sku = found.group(1) if found else ''
    price_data = item.get('price') or {}
    try:
        price = int(price_data.get('amount', 0)) // 100   # VK отдаёт цену в копейках
    except (TypeError, ValueError, AttributeError):
        price = 0
    return {
        'vk_id': item['id'],
        'article': sku or None,
        'name': str(item.get('title') or 'Товар').strip(),
        'price': price,
        'sizes': parse_sizes(item.get('description')),
    }


def fetch_market_catalog():
    """Скачивает все товары сообщества из VK Market (постранично по 200)."""
    items, offset = [], 0
    while True:
        resp = user_auth.method('market.get', {
            'owner_id': -GROUP_ID,
            'count': 200,
            'offset': offset,
        })
        batch = resp.get('items', [])
        items.extend(batch)
        offset += len(batch)
        if not batch or offset >= resp.get('count', 0):
            break
        time.sleep(0.4)   # лимит VK для пользовательского токена — 3 запроса в секунду
    return items


def sync_catalog():
    """Подтягивает каталог из VK Market в таблицу products.
    Возвращает (добавлено, изменено, пропущено без артикула, всего в VK)
    или None, если VK ничего не вернул."""
    items = fetch_market_catalog()
    if not items:
        print('Каталог VK Market пуст или недоступен — таблица products не тронута.')
        return None
    records = [market_item_record(i) for i in items if i.get('id')]
    no_article = sum(1 for r in records if not r['article'])
    added, changed, skipped = upsert_market_items(records)
    print(f'Каталог синхронизирован: в VK {len(records)}, добавлено {added}, изменено {changed}, '
          f'пропущено (нет артикула) {skipped}; без артикула в VK: {no_article}.')
    return added, changed, skipped, len(records)


def catalog_sync_loop():
    """Фоновый поток: обновляет каталог при старте и затем раз в CATALOG_SYNC_INTERVAL секунд."""
    while True:
        try:
            sync_catalog()
        except Exception:
            traceback.print_exc()
        time.sleep(CATALOG_SYNC_INTERVAL)


# ------------------------------------------------------------------ отправка

def write_massage(sender, massage, keyboard=None, attachment=None):
    payload = {'user_id': sender, 'message': massage, 'random_id': get_random_id()}
    if keyboard is not None:
        payload['keyboard'] = keyboard.get_keyboard()
    if attachment:
        payload['attachment'] = attachment
    auth.method('messages.send', payload)


def send_file(sender, path, message=''):
    """Загружает файл как документ и отправляет его пользователю в личные сообщения."""
    result = upload.document_message(path, title=os.path.basename(path), peer_id=sender)
    doc = result.get('doc', result) if isinstance(result, dict) else result[0]
    attachment = f"doc{doc['owner_id']}_{doc['id']}"
    if doc.get('access_key'):
        attachment += f"_{doc['access_key']}"
    write_massage(sender, message, attachment=attachment)


_photo_cache = {}   # путь к файлу -> вложение VK (чтобы не грузить QR каждый раз)


def write_massage_with_photo(sender, massage, photo_path, keyboard=None):
    """Отправляет сообщение с фото (фото загружается в VK один раз)."""
    try:
        attachment = _photo_cache.get(photo_path)
        if attachment is None:
            photo = upload.photo_messages(photos=photo_path)[0]
            attachment = f"photo{photo['owner_id']}_{photo['id']}"
            _photo_cache[photo_path] = attachment
        write_massage(sender, massage, keyboard, attachment)
    except Exception as e:
        _photo_cache.pop(photo_path, None)
        print(f'Ошибка отправки фото: {e}')
        # Если фото не ушло, отправляем хотя бы текст
        write_massage(sender, massage, keyboard)


def get_vk_user_name(user_id):
    """Возвращает 'Имя Фамилия' пользователя ВК."""
    try:
        user = auth.method('users.get', {'user_ids': user_id})[0]
        return f"{user['first_name']} {user['last_name']}"
    except Exception as e:
        print(f'Не удалось получить имя пользователя {user_id}: {e}')
        return f'Пользователь {user_id}'


# ------------------------------------------------------------------ корзина

def build_cart_lines(cart):
    """Строки корзины и итоговая сумма."""
    lines, total = [], 0
    for i, (article, size, qty) in enumerate(cart, start=1):
        product = get_product(article)
        if product:
            name, price, _ = product
            price = price or 0
        else:
            name, price = f'Артикул {article}', 0

        item_sum = price * qty
        total += item_sum

        if size and size != '-':
            lines.append(f'{i}. {name} / {size} / {qty} шт. — {item_sum} ₽')
        else:
            lines.append(f'{i}. {name} / {qty} шт. — {item_sum} ₽')
    return lines, total


def format_cart_text(sender, title='Ваша корзина:'):
    """Возвращает (текст корзины, итого). Для пустой корзины — (None, 0)."""
    cart = get_cart(sender)
    if not cart:
        return None, 0
    lines, total = build_cart_lines(cart)
    return '\n'.join([title, *lines, '', f'Итого: {total} ₽']), total


def show_cart(sender):
    user_screen[sender] = 'cart'
    text, _ = format_cart_text(sender)
    if text is None:
        write_massage(sender, 'Корзина пуста. Нажмите «Добавить товар», чтобы начать.',
                      get_cart_keyboard())
        return
    write_massage(sender, text, get_cart_keyboard())


def show_cart_edit(sender):
    text, _ = format_cart_text(sender, title='Что изменить:')
    if text is None:
        show_cart(sender)
        return
    user_screen[sender] = 'cart_edit'
    write_massage(sender, text, cart_edit_keyboard(sender))


def reset_state(sender):
    """Сбрасывает диалог в главное меню (корзина в базе остаётся)."""
    user_screen[sender] = 'main'
    temp_item.pop(sender, None)
    pending_market.pop(sender, None)


def parse_quantity(text):
    """Целое число от 1 до MAX_QTY или None."""
    if not (text.isascii() and text.isdigit()):
        return None
    qty = int(text)
    return qty if 1 <= qty <= MAX_QTY else None


def parse_position(text):
    """Номер позиции в корзине или None, если это не число."""
    if text.isascii() and text.isdigit():
        return int(text)
    return None


def ask_item_details(sender, article, product):
    """Начинает ввод размера/количества для выбранного товара."""
    name, _, has_size = product
    temp_item[sender] = {'article': article}
    if has_size:
        sizes = get_product_sizes(article)
        user_screen[sender] = 'cart_size'
        if sizes:
            temp_item[sender]['sizes'] = sizes
            write_massage(sender, f'Товар: {name}. Выберите размер:', get_size_keyboard(sizes))
        else:
            write_massage(sender, f'Товар: {name}. Введите размер:', get_item_input_keyboard())
    else:
        temp_item[sender]['size'] = '-'
        user_screen[sender] = 'cart_quantity'
        write_massage(sender, f'Товар: {name}. Введите количество (от 1 до {MAX_QTY}):',
                      get_item_input_keyboard())


# ------------------------------------------------------------------ VK Market

def extract_market_item(attachments):
    """Ищет товар из VK Market во вложениях и возвращает его данные."""
    if not attachments or not isinstance(attachments, list):
        return None

    for attach in attachments:
        if attach.get('type') != 'market':
            continue
        market_data = attach.get('market', {})
        owner_id = market_data.get('owner_id')
        item_id = market_data.get('id')
        if not owner_id or not item_id:
            continue

        try:
            response = user_auth.method('market.getById', {
                'item_ids': f'{owner_id}_{item_id}',
                'extended': 1,
            })
            if response.get('items'):
                return response['items'][0]
        except Exception as e:
            print(f'Ошибка market.getById: {e}')
    return None


def handle_market_item(sender, market_item):
    """Пользователь прислал товар из Market: обновляем каталог по свежим данным ВК
    и показываем карточку с кнопками."""
    record = market_item_record(market_item)
    # Товар мог появиться в VK после последней синхронизации — добавляем/обновляем сразу.
    upsert_market_items([record])
    product = get_product_by_vk_id(record['vk_id'])
    if not product:
        write_massage(sender, 'У этого товара в каталоге не указан артикул, поэтому добавить его '
                              'в корзину нельзя. Напишите администратору.', get_main_keyboard())
        reset_state(sender)
        return
    article, title, price, _ = product
    sizes = get_product_sizes(article)

    pending_market[sender] = {'article': article, 'title': title, 'price': price}
    user_screen[sender] = 'market_confirm'

    write_massage(
        sender,
        f'Вы прислали товар:\n'
        f'📦 {title}\n'
        f'💰 Цена: {price} ₽\n'
        f'🆔 Артикул: {article}\n'
        + (f'📏 Размеры: {", ".join(sizes)}\n' if sizes else '')
        + '\nДобавить его в корзину?',
        get_market_confirm_keyboard(article)
    )


def handle_add_market_to_cart(sender, article):
    """Кнопка «В корзину» под карточкой товара из Market."""
    product = get_product(article)

    if not product:
        write_massage(sender, 'Товар не найден. Пришлите его ещё раз или введите артикул '
                              'через «Корзина» → «Добавить товар».', get_main_keyboard())
        reset_state(sender)
        return

    if not product[1]:
        write_massage(sender, 'У этого товара не указана цена, добавить его в корзину нельзя. '
                              'Напишите администратору.', get_main_keyboard())
        reset_state(sender)
        return

    pending_market.pop(sender, None)
    ask_item_details(sender, article, product)


# ------------------------------------------------------------------ оплата и заказ

def start_payment(sender):
    text, _ = format_cart_text(sender)
    if text is None:
        write_massage(sender, 'Корзина пуста — сначала добавьте товары.', get_main_keyboard())
        return
    user_screen[sender] = 'payment_choice'
    write_massage(sender, text)
    write_massage(sender, 'Выберите удобный способ оплаты:', get_payment_keyboard())


def extract_receipt_photo(attachments):
    """Возвращает (id фото 'owner_id', вложение для отправки) или (None, None), если фото нет."""
    if not isinstance(attachments, list):
        return None, None
    for att in attachments:
        if att.get('type') == 'photo':
            photo = att['photo']
            photo_id = f"{photo['owner_id']}_{photo['id']}"
            attachment = f'photo{photo_id}'
            if photo.get('access_key'):
                attachment += f"_{photo['access_key']}"
            return photo_id, attachment
    return None, None


def handle_receipt(sender, text, attachments):
    """Экран ожидания чека: принимаем ТОЛЬКО фото, всё остальное отклоняем."""
    if text == 'Начать':
        reset_state(sender)
        write_massage(sender, 'Возврат в главное меню. Корзина сохранена.', get_main_keyboard())
        return
    if text == 'Отмена':
        go_main_menu(sender, 'Оплата отменена, заказ не создан. Корзина сохранена.\nГлавное меню:')
        return

    photo_id, photo_attachment = extract_receipt_photo(attachments)
    if not photo_id:
        write_massage(
            sender,
            '⚠️ Сейчас я жду только фото чека (скриншот или фото оплаты).\n'
            'Сообщения и другие файлы принять не могу. '
            'Пришлите фото или нажмите «Отмена».',
            get_item_input_keyboard()
        )
        return

    order_id = checkout_cart(sender, get_vk_user_name(sender), photo_id)
    if order_id is None:
        reset_state(sender)
        write_massage(sender, 'Корзина уже пуста.', get_main_keyboard())
        return

    reset_state(sender)
    write_massage(
        sender,
        f'✅ Чек получен! Заказ №{order_id} на сумму {get_order_total(order_id)} ₽ '
        f'отправлен на проверку.\nСпасибо за заказ! Мы сообщим, когда он будет доставлен.',
        get_main_keyboard()
    )
    notify_admin_new_order(order_id, photo_attachment)


def order_item_text(item, with_price=False):
    """'Название / размер — N шт.' для сообщений и Excel."""
    title = item['name'] or f"Артикул {item['article']}"
    size_part = f" / {item['size']}" if item['size'] and item['size'] != '-' else ''
    text = f"{title}{size_part} — {item['quantity']} шт."
    if with_price:
        text += f" × {item['price']} ₽"
    return text


def notify_admin_new_order(order_id, photo_attachment=None):
    """Уведомление админу №1: заказ сформирован."""
    order = get_order(order_id)
    items = get_order_items(order_id)
    lines = '\n'.join('• ' + order_item_text(i, with_price=True) for i in items)
    text = (
        f'🆕 Заказ №{order_id} сформирован\n'
        f'Покупатель: {order["user_name"]}\n'
        f'Страница: https://vk.com/id{order["user_id"]}\n\n'
        f'Товары:\n{lines}\n\n'
        f'Сумма: {get_order_total(order_id)} ₽\n'
        f'Статус: {order["status"]}\n\n'
        f'Когда заказ доставят, отправьте: /delivered {order_id}'
    )
    try:
        write_massage(ADMIN_ID, text, attachment=photo_attachment)
    except Exception as e:
        print(f'Не удалось отправить админу заказ №{order_id} с чеком: {e}')
        try:
            write_massage(ADMIN_ID, text + '\n\n(чек не удалось прикрепить — он есть в /export)')
        except Exception as e2:
            print(f'Не удалось уведомить админа о заказе №{order_id}: {e2}')


def mark_delivered(order_id):
    """Уведомление пользователю №2: заказ доставлен. Возвращает текст-ответ для админа."""
    order = get_order(order_id)
    if not order:
        return f'Заказ №{order_id} не найден.'
    if order['status'] == STATUS_DELIVERED:
        return f'Заказ №{order_id} уже отмечен как доставленный.'

    set_status(order_id, STATUS_DELIVERED)
    try:
        write_massage(
            order['user_id'],
            f'📦 Ваш заказ №{order_id} доставлен!\n'
            f'Спасибо за покупку в Evromodnik, будем рады видеть вас снова.',
            get_main_keyboard()
        )
    except Exception as e:
        print(f'Не удалось уведомить пользователя {order["user_id"]}: {e}')
        return (f'Заказ №{order_id} отмечен как доставленный, но сообщение покупателю '
                f'не отправилось: {e}')
    return f'✅ Заказ №{order_id} отмечен как доставленный, покупатель ({order["user_name"]}) уведомлён.'


# ------------------------------------------------------------------ экспорт

def export_orders_to_excel(path=ORDERS_XLSX):
    """Экспортирует заказы в Excel. Возвращает число заказов или None, если файл не удалось сохранить."""
    orders = get_orders_for_export()

    wb = Workbook()
    ws = wb.active
    ws.title = 'Заказы'

    headers = ['Имя пользователя', 'Страница VK', 'Номер заказа', 'Товары',
               'Сумма заказа', 'Статус', 'Фото оплаты']
    widths = [28, 32, 14, 60, 16, 16, 45]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'

    for o in orders:
        vk_link = f"https://vk.com/id{o['user_id']}"
        photo = o['payment_photo']
        photo_link = f'https://vk.com/photo{photo}' if photo and '_' in str(photo) else None
        items_text = '\n'.join(order_item_text(i) for i in o['items']) or '— нет товаров —'

        ws.append([
            o['user_name'] or f"Пользователь {o['user_id']}",
            vk_link,
            f"№{o['id']}",
            items_text,
            o['total'],
            o['status'],
            photo_link or '— нет фото —',
        ])
        row = ws.max_row
        for col in range(1, len(headers) + 1):
            ws.cell(row, col).alignment = Alignment(vertical='top', wrap_text=(col == 4))
        ws.cell(row, 5).number_format = '#,##0 "₽"'
        ws.cell(row, 2).hyperlink = vk_link
        ws.cell(row, 2).font = Font(color='0563C1', underline='single')
        if photo_link:
            ws.cell(row, 7).hyperlink = photo_link
            ws.cell(row, 7).font = Font(color='0563C1', underline='single')
        ws.row_dimensions[row].height = max(15, 15 * len(o['items']))

    try:
        wb.save(path)
        return len(orders)
    except PermissionError:
        print(f'Не удалось сохранить {path} — файл открыт в Excel.')
        return None


def export_catalog_to_excel(path=CATALOG_XLSX):
    """Экспортирует каталог товаров в Excel. Возвращает число товаров или None, если файл не сохранился."""
    products = get_all_products()

    wb = Workbook()
    ws = wb.active
    ws.title = 'Каталог'

    headers = ['Артикул', 'Название', 'Цена', 'Нужен размер', 'Размеры', 'Ссылка на товар']
    widths = [16, 70, 12, 14, 28, 50]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}1'

    for article, name, price, has_size, sizes, vk_id in products:
        link = f'https://vk.com/market-{GROUP_ID}?w=product-{GROUP_ID}_{vk_id}' if vk_id else None
        ws.append([
            article,
            name,
            price,
            'да' if has_size else 'нет',
            ', '.join(sizes),
            link or '— нет в VK —',
        ])
        row = ws.max_row
        ws.cell(row, 1).number_format = '@'          # артикул — текстом (чтобы не терять нули в начале)
        ws.cell(row, 3).number_format = '#,##0 "₽"'
        if link:
            ws.cell(row, 6).hyperlink = link
            ws.cell(row, 6).font = Font(color='0563C1', underline='single')

    try:
        wb.save(path)
        return len(products)
    except PermissionError:
        print(f'Не удалось сохранить {path} — файл открыт в Excel.')
        return None


# ------------------------------------------------------------------ команды админа

def handle_admin_command(sender, text):
    if text.lower() == '/export':
        count = export_orders_to_excel()
        if count is None:
            write_massage(sender, 'Не удалось сохранить orders.xlsx — закройте файл в Excel и повторите.')
        else:
            caption = f'Экспортировано заказов: {count}'
            try:
                send_file(sender, ORDERS_XLSX, caption)
            except Exception as e:
                print(f'Не удалось отправить orders.xlsx: {e}')
                write_massage(sender, f'{caption}\nФайл сохранён на сервере (orders.xlsx), '
                                      f'но отправить его в чат не получилось: {e}')
        return

    if text.lower() == '/sync':
        write_massage(sender, 'Обновляю каталог из VK Market, это может занять до минуты...')
        try:
            result = sync_catalog()
        except Exception as e:
            print(f'Ошибка синхронизации каталога: {e}')
            write_massage(sender, f'Не удалось обновить каталог: {e}')
            return
        if result is None:
            write_massage(sender, 'VK Market не вернул товаров — таблица не изменена.')
        else:
            added, changed, skipped, total = result
            text = (f'✅ Каталог обновлён.\nТоваров в VK: {total}\n'
                    f'Добавлено новых: {added}\nИзменено: {changed}')
            if skipped:
                text += f'\nПропущено (в карточке нет артикула): {skipped}'
            write_massage(sender, text)
        return

    if text.lower() == '/catalog':
        count = export_catalog_to_excel()
        if count is None:
            write_massage(sender, 'Не удалось сохранить catalog.xlsx — закройте файл в Excel и повторите.')
        else:
            caption = f'Товаров в каталоге: {count}'
            try:
                send_file(sender, CATALOG_XLSX, caption)
            except Exception as e:
                print(f'Не удалось отправить catalog.xlsx: {e}')
                write_massage(sender, f'{caption}\nФайл сохранён на сервере (catalog.xlsx), '
                                      f'но отправить его в чат не получилось: {e}')
        return

    m = re.fullmatch(r'/delivered(?:\s+(.*))?', text, re.IGNORECASE | re.DOTALL)
    if m:
        rest = (m.group(1) or '').strip()
        if not rest:
            write_massage(sender, 'Укажите номер заказа или несколько номеров, например:\n'
                                  '/delivered 5\n/delivered 5 6 7\n/delivered 5, 6, 7')
        elif not re.fullmatch(r'[\d\s,;№#]+', rest):
            write_massage(sender, 'Не понял номера заказов. Укажите только числа через пробел '
                                  'или запятую, например: /delivered 5 6 7')
        else:
            order_ids = list(dict.fromkeys(int(n) for n in re.findall(r'\d+', rest)))
            if not order_ids or any(order_id < 1 for order_id in order_ids):
                write_massage(sender, 'Номер заказа должен быть положительным целым числом. Пример: /delivered 5')
            else:
                results = [mark_delivered(order_id) for order_id in order_ids]
                write_massage(sender, '\n'.join(results))
        return

    write_massage(sender, 'Команда не распознана или введена неправильно.\n\n'
                          'Доступные команды администратора:\n'
                          '/export — выгрузить заказы в orders.xlsx (без дополнительных параметров)\n'
                          '/sync — сразу обновить каталог из VK Market\n'
                          '/catalog — выгрузить каталог товаров в catalog.xlsx\n'
                          '/delivered N — отметить заказ №N доставленным и уведомить покупателя\n'
                          '/delivered N M K — отметить несколько заказов (номера через пробел или запятую).')


# ------------------------------------------------------------------ обработка сообщений

def parse_payload(payload):
    """payload кнопки приходит строкой JSON — превращаем в словарь."""
    if isinstance(payload, dict):
        return payload
    if isinstance(payload, str):
        try:
            data = json.loads(payload)
            return data if isinstance(data, dict) else {}
        except json.JSONDecodeError:
            pass
    return {}


def load_saved_states():
    """Восстанавливает состояние диалогов из базы после перезапуска бота."""
    for user_id, (screen, temp, pending) in load_all_states().items():
        user_screen[user_id] = screen
        if temp:
            temp_item[user_id] = temp
        if pending:
            pending_market[user_id] = pending


def persist_state(user_id):
    try:
        save_state(user_id, user_screen.get(user_id, 'main'),
                   temp_item.get(user_id), pending_market.get(user_id))
    except Exception:
        traceback.print_exc()


def handle_message(message):
    """Обрабатывает сообщение и после этого сохраняет состояние диалога в базу."""
    sender = message['from_id']
    try:
        process_message(message)
    finally:
        persist_state(sender)


def process_message(message):
    text = (message.get('text') or '').strip()
    sender = message['from_id']
    attachments = message.get('attachments') or []
    payload = message.get('payload')
    screen = user_screen.get(sender, 'main')

    # Проверяем целостность временных данных после перезапуска/сбоя состояния.
    if screen in ('cart_size', 'cart_quantity', 'cart_change_qty', 'cart_resize_new'):
        if not temp_item.get(sender):
            reset_state(sender)
            write_massage(sender,
                          'Не удалось восстановить шаг добавления или изменения товара. '
                          'Действие отменено, корзина не изменена. Начните заново.',
                          get_main_keyboard())
            return

    # --- Команды админа ---
    if sender == ADMIN_ID and text.startswith('/'):
        handle_admin_command(sender, text)
        return

    # --- Ожидание чека: принимаем только фото ---
    if screen == 'payment_waiting_screenshot':
        handle_receipt(sender, text, attachments)
        return

    # --- Товар из VK Market ---
    market_item = extract_market_item(attachments)
    if market_item:
        handle_market_item(sender, market_item)
        return

    if text == 'В корзину':
        article = parse_payload(payload).get('article')
        if article:
            handle_add_market_to_cart(sender, str(article))
        else:
            write_massage(sender, 'Не удалось определить товар по кнопке. Откройте карточку товара в VK Market и отправьте её боту ещё раз.',
                          get_main_keyboard())
        return

    # --- Команды меню (работают с любого экрана) ---
    if text == 'Начать':
        reset_state(sender)
        write_massage(sender, 'Вас приветствует сообщество Evromodnik! Выберите раздел.',
                      get_main_keyboard())
        return

    if text == 'Корзина':
        show_cart(sender)
        return

    if text == 'Оплатить Заказ':
        start_payment(sender)
        return

    # --- Меню корзины ---
    if screen == 'cart':
        if text == 'Добавить товар':
            user_screen[sender] = 'cart_article'
            write_massage(sender, 'Введите артикул товара:', get_item_input_keyboard())
            return
        if text == 'Изменить корзину':
            show_cart_edit(sender)
            return
        if text == 'Очистить корзину':
            clear_cart(sender)
            user_screen[sender] = 'main'
            write_massage(sender, 'Корзина очищена.', get_main_keyboard())
            return
        if text == 'Назад':
            user_screen[sender] = 'main'
            write_massage(sender, 'Главное меню:', get_main_keyboard())
            return

    # --- Выбор способа оплаты ---
    if screen == 'payment_choice':
        if text == 'Оплатить по реквизитам':
            user_screen[sender] = 'payment_waiting_receipt'
            write_massage(sender, REQUISITES_TEXT, get_payment_confirmation_keyboard())
            return
        if text == 'Оплатить через СБП':
            user_screen[sender] = 'payment_waiting_receipt'
            write_massage_with_photo(sender, SBP_TEXT, QR_PATH, get_payment_confirmation_keyboard())
            return
        if text == 'Назад в меню':
            go_main_menu(sender)
            return

    # --- Подтверждение оплаты ---
    if screen == 'payment_waiting_receipt':
        if text == 'Я оплатил(а)':
            if not get_cart(sender):
                reset_state(sender)
                write_massage(sender, 'Корзина уже пуста.', get_main_keyboard())
                return
            # Заказ создаётся только после получения чека
            user_screen[sender] = 'payment_waiting_screenshot'
            write_massage(sender,
                          'Пришлите фото чека. Заказ будет оформлен после получения чека.',
                          get_item_input_keyboard())
            return
        if text == 'Отмена':
            go_main_menu(sender, 'Оплата отменена. Корзина сохранена.\nГлавное меню:')
            return

    # --- Меню редактирования корзины ---
    if screen == 'cart_edit':
        if not get_cart(sender):
            user_screen[sender] = 'cart'
            write_massage(sender, 'Корзина уже пуста. Сначала добавьте товар.', get_cart_keyboard())
            return
        if text == 'Удалить товар':
            user_screen[sender] = 'cart_remove'
            write_massage(sender, 'Введите номер товара для удаления:', get_item_input_keyboard())
            return
        if text == 'Изменить количество':
            user_screen[sender] = 'cart_change_num'
            write_massage(sender, 'Введите номер товара:', get_item_input_keyboard())
            return
        if text == 'Изменить размер':
            if not cart_has_sized_items(sender):
                write_massage(sender, 'В корзине нет товаров с размером.',
                              cart_edit_keyboard(sender))
                return
            user_screen[sender] = 'cart_resize_num'
            write_massage(sender, 'Введите номер товара, у которого нужно изменить размер:',
                          get_item_input_keyboard())
            return
        if text == 'Назад':
            show_cart(sender)
            return

    # --- Удаление товара ---
    if screen == 'cart_remove':
        if text == 'Отмена':
            show_cart(sender)
            return
        pos = parse_position(text)
        if pos is None:
            write_massage(sender, 'Номер должен быть целым положительным числом. Например: 1. Попробуйте ещё раз или нажмите «Отмена».', get_item_input_keyboard())
            return
        if remove_cart_item(sender, pos):
            write_massage(sender, '✅ Товар удалён.')
            show_cart(sender)
        else:
            write_massage(sender, 'Такого номера нет в корзине. Проверьте список и введите существующий номер или нажмите «Отмена».',
                          get_item_input_keyboard())
        return

    # --- Выбор номера для изменения количества ---
    if screen == 'cart_change_num':
        if text == 'Отмена':
            show_cart(sender)
            return
        pos = parse_position(text)
        if pos is None:
            write_massage(sender, 'Номер должен быть целым положительным числом. Например: 1. Попробуйте ещё раз или нажмите «Отмена».', get_item_input_keyboard())
            return
        if not 1 <= pos <= len(get_cart(sender)):
            write_massage(sender, 'Такого номера нет. Введите номер из списка.',
                          get_item_input_keyboard())
            return
        temp_item[sender] = {'position': pos}
        user_screen[sender] = 'cart_change_qty'
        write_massage(sender, f'Введите новое количество (от 1 до {MAX_QTY}):',
                      get_item_input_keyboard())
        return

    # --- Ввод нового количества ---
    if screen == 'cart_change_qty':
        if text == 'Отмена':
            temp_item.pop(sender, None)
            show_cart(sender)
            return
        qty = parse_quantity(text)
        if qty is None:
            write_massage(sender, f'Введите целое число от 1 до {MAX_QTY}.',
                          get_item_input_keyboard())
            return
        if update_cart_quantity(sender, temp_item[sender]['position'], qty):
            write_massage(sender, 'Количество обновлено.')
        else:
            write_massage(sender, 'Не удалось обновить.')
        temp_item.pop(sender, None)
        show_cart(sender)
        return

    # --- Выбор номера для изменения размера ---
    if screen == 'cart_resize_num':
        if text == 'Отмена':
            show_cart(sender)
            return
        pos = parse_position(text)
        cart = get_cart(sender)
        if pos is None:
            write_massage(sender, 'Номер должен быть целым положительным числом. Например: 1. Попробуйте ещё раз или нажмите «Отмена».', get_item_input_keyboard())
            return
        if not 1 <= pos <= len(cart):
            write_massage(sender, 'Такого номера нет. Введите номер из списка.',
                          get_item_input_keyboard())
            return
        if cart[pos - 1][1] == '-':
            write_massage(sender, 'У этого товара нет размера. Введите номер другого товара.',
                          get_item_input_keyboard())
            return
        sizes = get_product_sizes(cart[pos - 1][0])
        temp_item[sender] = {'position': pos}
        user_screen[sender] = 'cart_resize_new'
        if sizes:
            temp_item[sender]['sizes'] = sizes
            write_massage(sender, f'Текущий размер: {cart[pos - 1][1]}. Выберите новый размер:',
                          get_size_keyboard(sizes))
        else:
            write_massage(sender, f'Текущий размер: {cart[pos - 1][1]}. Введите новый размер:',
                          get_item_input_keyboard())
        return

    # --- Ввод нового размера ---
    if screen == 'cart_resize_new':
        if text == 'Отмена':
            temp_item.pop(sender, None)
            show_cart(sender)
            return
        sizes = temp_item[sender].get('sizes')
        if sizes:
            chosen = match_size(sizes, text)
            if not chosen:
                write_massage(sender, 'Выберите размер кнопкой из списка или нажмите «Отмена».',
                              get_size_keyboard(sizes))
                return
            text = chosen
        if not text or len(text) > MAX_SIZE_LEN:
            write_massage(sender, f'Введите размер текстом (до {MAX_SIZE_LEN} символов).',
                          get_item_input_keyboard())
            return
        result = update_cart_size(sender, temp_item[sender]['position'], text)
        temp_item.pop(sender, None)
        if result == 'ok':
            write_massage(sender, 'Размер обновлён.')
        else:
            write_massage(sender, 'Не удалось изменить размер.')
        show_cart(sender)
        return

    # --- Отмена при вводе товара ---
    if screen in ('cart_article', 'cart_size', 'cart_quantity') and text == 'Отмена':
        temp_item.pop(sender, None)
        show_cart(sender)
        return

    # --- Шаги ввода товара ---
    if screen == 'cart_article':
        product = get_product(text)
        if not product:
            write_massage(sender, 'Такого артикула нет. Попробуйте другой.',
                          get_item_input_keyboard())
            return
        ask_item_details(sender, text, product)
        return

    if screen == 'cart_size':
        sizes = temp_item[sender].get('sizes')
        if sizes:
            chosen = match_size(sizes, text)
            if not chosen:
                write_massage(sender, 'Выберите размер кнопкой из списка или нажмите «Отмена».',
                              get_size_keyboard(sizes))
                return
            text = chosen
        if not text or len(text) > MAX_SIZE_LEN:
            write_massage(sender, f'Введите размер текстом (до {MAX_SIZE_LEN} символов).',
                          get_item_input_keyboard())
            return
        temp_item[sender]['size'] = text
        temp_item[sender].pop('sizes', None)
        user_screen[sender] = 'cart_quantity'
        write_massage(sender, f'Введите количество (от 1 до {MAX_QTY}):',
                      get_item_input_keyboard())
        return

    if screen == 'cart_quantity':
        qty = parse_quantity(text)
        if qty is None:
            write_massage(sender, f'Введите целое число от 1 до {MAX_QTY}.',
                          get_item_input_keyboard())
            return
        item = temp_item.pop(sender)
        add_to_cart(sender, item['article'], item['size'], qty)
        user_screen[sender] = 'main'
        write_massage(sender, '✅ Товар успешно добавлен в корзину!', get_main_keyboard())
        return

    # --- «Отмена» на любом другом экране (например, под карточкой товара из Market) ---
    if text == 'Отмена':
        reset_state(sender)
        write_massage(sender, 'Отменено. Главное меню:', get_main_keyboard())
        return

    # --- Любое другое действие, не предусмотренное текущим экраном ---
    send_invalid_action(sender, screen, attachments)


# ------------------------------------------------------------------ запуск

def main():
    init_db()
    load_saved_states()
    threading.Thread(target=catalog_sync_loop, daemon=True).start()
    print('Бот запущен.')

    while True:
        try:
            for event in longpoll.listen():
                if event.type != VkBotEventType.MESSAGE_NEW:
                    continue
                message = event.message
                try:
                    handle_message(message)
                except Exception:
                    # Ошибка в одном сообщении не должна ронять бота
                    traceback.print_exc()
                    try:
                        write_massage(message['from_id'],
                                      'Что-то пошло не так. Нажмите «Начать» и попробуйте ещё раз.',
                                      get_main_keyboard())
                    except Exception:
                        pass
        except Exception as e:
            print(f'⚠️ Ошибка подключения или сбой: {e}')
            print('Переподключение через 5 секунд...')
            time.sleep(5)


if __name__ == '__main__':
    main()