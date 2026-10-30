import json
import os
import re
import sys
import time
import traceback

import vk_api
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from vk_api.bot_longpoll import VkBotEventType, VkBotLongPoll
from vk_api.upload import VkUpload
from vk_api.utils import get_random_id

from database import (
    STATUS_DELIVERED,
    add_product, add_to_cart, checkout_cart, clear_cart, get_cart, get_order,
    get_order_items, get_order_total, get_orders_for_export, get_product,
    get_product_by_name, init_db, load_all_states, remove_cart_item, save_state,
    set_status, update_cart_quantity, update_cart_size,
)
from keyboards import (
    get_cart_edit_keyboard, get_cart_keyboard, get_item_input_keyboard,
    get_main_keyboard, get_market_confirm_keyboard,
    get_payment_confirmation_keyboard, get_payment_keyboard,
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

PRODUCTS_XLSX = os.path.join(BASE_DIR, 'products.xlsx')
ORDERS_XLSX = os.path.join(BASE_DIR, 'orders.xlsx')
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


def cart_has_sized_items(user_id):
    """True, если в корзине есть хотя бы один товар с размером."""
    return any(size and size != '-' for _, size, _ in get_cart(user_id))


def cart_edit_keyboard(user_id):
    """Меню «Изменить корзину»; кнопка «Изменить размер» только если есть товары с размером."""
    return get_cart_edit_keyboard(show_change_size=cart_has_sized_items(user_id))


def keyboard_for_screen(user_id, screen):
    if screen == 'cart_edit':
        return cart_edit_keyboard(user_id)
    return SCREEN_KEYBOARDS.get(screen, get_main_keyboard)()


def go_main_menu(sender, text='Главное меню:'):
    """Сброс состояния и возврат в главное меню (корзина в базе остаётся)."""
    reset_state(sender)
    write_massage(sender, text, get_main_keyboard())


# ------------------------------------------------------------------ каталог

def load_products_from_excel(path=PRODUCTS_XLSX):
    """Загружает каталог из Excel: article | name | price | has_size."""
    try:
        wb = load_workbook(path)
    except FileNotFoundError:
        print(f'Файл {path} не найден — пропускаю загрузку.')
        return

    def clean_article(value):
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        return str(value).strip() if value is not None else None

    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        article = clean_article(row[0])
        name = str(row[1]).strip() if row[1] else None
        if not article or not name:
            continue

        price = row[2]
        has_size_raw = row[3] if len(row) > 3 else None
        # Пустая ячейка = 1 (нужен размер). 0 остаётся 0.
        has_size = 1 if has_size_raw in (None, '') else int(has_size_raw)

        add_product(article, name, int(price or 0), has_size)


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
        user_screen[sender] = 'cart_size'
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
    """Пользователь прислал товар из Market: показываем карточку с кнопками."""
    item_id = market_item['id']
    title = str(market_item.get('title', 'Товар')).strip()

    price_data = market_item.get('price', {})
    price_amount = price_data.get('amount', 0) if isinstance(price_data, dict) else 0
    price = int(price_amount) // 100

    sku = market_item.get('sku')
    article = None

    # 1. Ищем по названию в каталоге (самый надёжный способ)
    found_by_name = get_product_by_name(title)
    if found_by_name:
        article = found_by_name[0]
    # 2. Если нет — проверяем SKU из ВК
    elif sku and str(sku).strip() and str(sku).lower() != 'none':
        if get_product(str(sku).strip()):
            article = str(sku).strip()

    if article:
        # Товар есть в каталоге — обновляем цену и название (вдруг поменяли в ВК)
        has_size = get_product(article)[2]
        add_product(article, title, price, has_size=has_size)
    else:
        # Товара нет в каталоге — используем ID из ВК, в базу попадёт только после «В корзину»
        article = str(item_id)

    pending_market[sender] = {'article': article, 'title': title, 'price': price}
    user_screen[sender] = 'market_confirm'

    write_massage(
        sender,
        f'Вы прислали товар:\n'
        f'📦 {title}\n'
        f'💰 Цена: {price} ₽\n'
        f'🆔 Артикул: {article}\n\n'
        f'Добавить его в корзину?',
        get_market_confirm_keyboard(article)
    )


def handle_add_market_to_cart(sender, article):
    """Кнопка «В корзину» под карточкой товара из Market."""
    product = get_product(article)

    if not product:
        # Товара нет в каталоге — добавляем его из данных, которые прислал ВК
        pending = pending_market.get(sender)
        if pending and pending['article'] == article and pending['price'] > 0:
            add_product(article, pending['title'], pending['price'], has_size=1)
            product = get_product(article)
        elif pending and pending['article'] == article:
            write_massage(sender, 'У этого товара не указана цена, добавить его в корзину нельзя. '
                                  'Напишите администратору.', get_main_keyboard())
            reset_state(sender)
            return

    if not product:
        write_massage(sender, 'Товар не найден. Пришлите его ещё раз или введите артикул '
                              'через «Корзина» → «Добавить товар».', get_main_keyboard())
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


# ------------------------------------------------------------------ команды админа

def handle_admin_command(sender, text):
    if text == '/export':
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
            results = [mark_delivered(order_id) for order_id in order_ids]
            write_massage(sender, '\n'.join(results))
        return

    write_massage(sender, 'Команды админа:\n'
                          '/export — выгрузить заказы в orders.xlsx\n'
                          '/delivered N — отметить заказ №N доставленным и уведомить покупателя\n'
                          '/delivered N M K — то же для нескольких заказов сразу (через пробел или запятую)')


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
            write_massage(sender, 'Введите число.', get_item_input_keyboard())
            return
        if remove_cart_item(sender, pos):
            write_massage(sender, 'Товар удалён.')
        else:
            write_massage(sender, 'Такого номера нет.')
        show_cart(sender)
        return

    # --- Выбор номера для изменения количества ---
    if screen == 'cart_change_num':
        if text == 'Отмена':
            show_cart(sender)
            return
        pos = parse_position(text)
        if pos is None:
            write_massage(sender, 'Введите число.', get_item_input_keyboard())
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
            write_massage(sender, 'Введите число.', get_item_input_keyboard())
            return
        if not 1 <= pos <= len(cart):
            write_massage(sender, 'Такого номера нет. Введите номер из списка.',
                          get_item_input_keyboard())
            return
        if cart[pos - 1][1] == '-':
            write_massage(sender, 'У этого товара нет размера. Введите номер другого товара.',
                          get_item_input_keyboard())
            return
        temp_item[sender] = {'position': pos}
        user_screen[sender] = 'cart_resize_new'
        write_massage(sender, f'Текущий размер: {cart[pos - 1][1]}. Введите новый размер:',
                      get_item_input_keyboard())
        return

    # --- Ввод нового размера ---
    if screen == 'cart_resize_new':
        if text == 'Отмена':
            temp_item.pop(sender, None)
            show_cart(sender)
            return
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
        if not text or len(text) > MAX_SIZE_LEN:
            write_massage(sender, f'Введите размер текстом (до {MAX_SIZE_LEN} символов).',
                          get_item_input_keyboard())
            return
        temp_item[sender]['size'] = text
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

    # --- Непонятное сообщение ---
    keyboard = keyboard_for_screen(sender, screen)
    write_massage(sender, 'Не понял вас. Воспользуйтесь кнопками ниже или нажмите «Начать».',
                  keyboard)


# ------------------------------------------------------------------ запуск

def main():
    init_db()
    load_saved_states()
    load_products_from_excel()
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