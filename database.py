import json
import os
import sqlite3
from contextlib import contextmanager

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'bot_db.sqlite')

STATUS_PROCESSING = 'В обработке'
STATUS_DELIVERED = 'Доставлен'


@contextmanager
def connect():
    """Соединение с базой: commit при успехе, rollback при ошибке, close всегда."""
    conn = sqlite3.connect(DB_PATH)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with connect() as conn:
        cursor = conn.cursor()

        # --- Каталог товаров ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                article TEXT UNIQUE NOT NULL,
                name TEXT NOT NULL,
                price INTEGER,
                has_size INTEGER DEFAULT 1
            )
        ''')

        # --- Заказы ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                status TEXT DEFAULT 'В обработке'
            )
        ''')

        # --- Строки заказов ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS order_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                article TEXT NOT NULL,
                size TEXT NOT NULL,
                quantity INTEGER NOT NULL
            )
        ''')

        # --- Корзины ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS carts (
                user_id INTEGER PRIMARY KEY,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # --- Товары в корзине ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cart_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                article TEXT NOT NULL,
                size TEXT NOT NULL,
                quantity INTEGER NOT NULL
            )
        ''')

        # --- Состояние диалогов (экран, вводимый товар) — переживает перезапуск бота ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_state (
                user_id INTEGER PRIMARY KEY,
                screen TEXT NOT NULL,
                temp_item TEXT,
                pending_market TEXT
            )
        ''')

        # --- Миграции: добавляем колонки, если их ещё нет ---
        for col_sql in [
            'ALTER TABLE orders ADD COLUMN user_name TEXT',
            'ALTER TABLE orders ADD COLUMN payment_photo TEXT',
            'ALTER TABLE order_items ADD COLUMN price INTEGER',  # цена на момент заказа
            'ALTER TABLE products ADD COLUMN vk_id INTEGER',     # id товара в VK Market
            'ALTER TABLE products ADD COLUMN sizes TEXT',        # доступные размеры (JSON-список)
        ]:
            try:
                cursor.execute(col_sql)
            except sqlite3.OperationalError:
                pass  # колонка уже существует

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_products_vk_id ON products(vk_id)')

        # Раньше при отсутствии артикула подставлялся id товара VK. Такие строки (если на них
        # нет ссылок из заказов и корзин) удаляем — правильный артикул подтянется при синхронизации.
        cursor.execute('''
            DELETE FROM products
            WHERE vk_id IS NOT NULL
              AND (article = CAST(vk_id AS TEXT) OR article = 'vk' || vk_id)
              AND article NOT IN (SELECT article FROM order_items)
              AND article NOT IN (SELECT article FROM cart_items)
        ''')

        # Старым строкам заказов проставляем цену из каталога (один раз)
        cursor.execute('''
            UPDATE order_items
            SET price = (SELECT p.price FROM products p WHERE p.article = order_items.article)
            WHERE price IS NULL
        ''')


# ---------------------------------------------------------------- состояние диалогов

def load_all_states():
    """{user_id: (screen, temp_item | None, pending_market | None)} — вызывается при старте бота."""
    with connect() as conn:
        rows = conn.execute(
            'SELECT user_id, screen, temp_item, pending_market FROM user_state'
        ).fetchall()
    states = {}
    for user_id, screen, temp, pending in rows:
        try:
            temp = json.loads(temp) if temp else None
            pending = json.loads(pending) if pending else None
        except json.JSONDecodeError:
            temp, pending = None, None
        states[user_id] = (screen, temp, pending)
    return states


def save_state(user_id, screen, temp_item=None, pending_market=None):
    """Сохраняет состояние диалога. Для «чистого» главного меню запись удаляется."""
    with connect() as conn:
        if screen == 'main' and not temp_item and not pending_market:
            conn.execute('DELETE FROM user_state WHERE user_id = ?', (user_id,))
            return
        conn.execute('''
            INSERT INTO user_state (user_id, screen, temp_item, pending_market)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                screen = excluded.screen,
                temp_item = excluded.temp_item,
                pending_market = excluded.pending_market
        ''', (
            user_id, screen,
            json.dumps(temp_item, ensure_ascii=False) if temp_item else None,
            json.dumps(pending_market, ensure_ascii=False) if pending_market else None,
        ))


# ---------------------------------------------------------------- каталог

def add_product(article, name, price=None, has_size=1):
    with connect() as conn:
        conn.execute('''
            INSERT INTO products (article, name, price, has_size)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(article) DO UPDATE SET
                name = excluded.name,
                price = excluded.price,
                has_size = excluded.has_size
        ''', (article, name, price, has_size))


def get_product(article):
    """Возвращает (name, price, has_size) по артикулу, либо None."""
    with connect() as conn:
        return conn.execute(
            'SELECT name, price, has_size FROM products WHERE article = ?', (article,)
        ).fetchone()


def get_product_by_vk_id(vk_id):
    """Возвращает (article, name, price, has_size) по id товара в VK Market, либо None."""
    with connect() as conn:
        return conn.execute(
            'SELECT article, name, price, has_size FROM products WHERE vk_id = ?', (vk_id,)
        ).fetchone()


def get_product_sizes(article):
    """Список доступных размеров товара ([] — если не заданы)."""
    with connect() as conn:
        row = conn.execute('SELECT sizes FROM products WHERE article = ?', (article,)).fetchone()
    if not row or not row[0]:
        return []
    try:
        sizes = json.loads(row[0])
    except json.JSONDecodeError:
        return []
    return [str(s) for s in sizes] if isinstance(sizes, list) else []


def _norm_name(name):
    # casefold() работает с кириллицей, в отличие от SQLite LOWER()
    return ' '.join((name or '').split()).casefold()


def upsert_market_items(records):
    """Синхронизирует товары из VK Market с таблицей products.

    records: [{'vk_id': int, 'article': str | None, 'name': str, 'price': int,
               'sizes': [str, ...]}, ...]
    Артикул — только SKU из карточки VK. id товара VK как артикул НЕ используется.

    Для каждого товара ищем строку в таблице: по vk_id -> по артикулу -> по названию
    (так старые товары, загруженные из Excel, не дублируются, а «привязываются» к VK).
    Найденному обновляем название, цену и размеры. Размер нужен (has_size = 1) только если
    в карточке VK указаны размеры. Если у товара нет артикула и его нет в таблице —
    он пропускается.
    Возвращает (добавлено, изменено, пропущено).
    """
    added = changed = skipped = 0
    with connect() as conn:
        cur = conn.cursor()
        rows = cur.execute(
            'SELECT id, article, name, price, vk_id, sizes, has_size FROM products'
        ).fetchall()
        by_vk = {r[4]: r[0] for r in rows if r[4] is not None}
        by_article = {r[1]: r[0] for r in rows}
        art_of = {r[0]: r[1] for r in rows}
        vk_of = {r[0]: r[4] for r in rows}
        data = {r[0]: (r[2], r[3], r[5], r[6]) for r in rows}
        by_name = {}
        for pid, _, name, _, vk, _, _ in rows:
            if vk is None:
                by_name.setdefault(_norm_name(name), pid)

        for rec in records:
            vk_id, article, name, price = rec['vk_id'], rec.get('article'), rec['name'], rec['price']
            sizes = rec.get('sizes') or []
            sizes_json = json.dumps(sizes, ensure_ascii=False) if sizes else None
            has_size = 1 if sizes else 0

            pid = by_vk.get(vk_id)
            if pid is None and article:
                cand = by_article.get(article)
                if cand is not None and vk_of.get(cand) in (None, vk_id):
                    pid = cand
            if pid is None:
                pid = by_name.pop(_norm_name(name), None)
                if pid is not None and vk_of.get(pid) is not None:
                    pid = None  # эта строка уже привязана к другому товару VK

            if pid is not None:
                old_article = art_of[pid]
                # Исправляем артикул, который раньше подставили как id товара VK
                if (article and article != old_article and old_article == str(vk_id)
                        and article not in by_article):
                    cur.execute('UPDATE products SET article = ? WHERE id = ?', (article, pid))
                    cur.execute('UPDATE order_items SET article = ? WHERE article = ?',
                                (article, old_article))
                    cur.execute('UPDATE cart_items SET article = ? WHERE article = ?',
                                (article, old_article))
                    by_article.pop(old_article, None)
                    by_article[article] = pid
                    art_of[pid] = article
                    changed += 1
                if data[pid] != (name, price, sizes_json, has_size) or vk_of.get(pid) != vk_id:
                    cur.execute(
                        'UPDATE products SET vk_id = ?, name = ?, price = ?, sizes = ?, has_size = ? '
                        'WHERE id = ?', (vk_id, name, price, sizes_json, has_size, pid))
                    changed += 1
                data[pid] = (name, price, sizes_json, has_size)
            else:
                if not article or article in by_article:
                    skipped += 1    # нет артикула (или SKU занят другим товаром VK)
                    continue
                cur.execute(
                    'INSERT INTO products (article, name, price, has_size, vk_id, sizes) '
                    'VALUES (?, ?, ?, ?, ?, ?)',
                    (article, name, price, has_size, vk_id, sizes_json)
                )
                pid = cur.lastrowid
                by_article[article] = pid
                art_of[pid] = article
                data[pid] = (name, price, sizes_json, has_size)
                added += 1
            by_vk[vk_id] = pid
            vk_of[pid] = vk_id
    return added, changed, skipped


def get_all_products():
    """Весь каталог для экспорта: [(article, name, price, has_size, [sizes], vk_id), ...]."""
    with connect() as conn:
        rows = conn.execute(
            'SELECT article, name, price, has_size, sizes, vk_id FROM products ORDER BY id'
        ).fetchall()
    result = []
    for article, name, price, has_size, sizes, vk_id in rows:
        try:
            size_list = json.loads(sizes) if sizes else []
        except json.JSONDecodeError:
            size_list = []
        result.append((article, name, price, has_size, size_list, vk_id))
    return result


# ---------------------------------------------------------------- корзина

def get_cart(user_id):
    """[(article, size, quantity), ...] в порядке добавления."""
    with connect() as conn:
        return conn.execute(
            'SELECT article, size, quantity FROM cart_items WHERE user_id = ? ORDER BY id',
            (user_id,)
        ).fetchall()


def add_to_cart(user_id, article, size, quantity):
    """Добавляет товар. Если такой же артикул с таким же размером уже есть — суммирует количество."""
    with connect() as conn:
        cursor = conn.cursor()
        cursor.execute('INSERT OR IGNORE INTO carts (user_id) VALUES (?)', (user_id,))
        row = cursor.execute(
            'SELECT id FROM cart_items '
            'WHERE user_id = ? AND article = ? AND LOWER(size) = LOWER(?)',
            (user_id, article, size)
        ).fetchone()
        if row:
            cursor.execute(
                'UPDATE cart_items SET quantity = quantity + ? WHERE id = ?',
                (quantity, row[0])
            )
        else:
            cursor.execute(
                'INSERT INTO cart_items (user_id, article, size, quantity) VALUES (?, ?, ?, ?)',
                (user_id, article, size, quantity)
            )


def clear_cart(user_id):
    with connect() as conn:
        conn.execute('DELETE FROM cart_items WHERE user_id = ?', (user_id,))
        conn.execute('DELETE FROM carts WHERE user_id = ?', (user_id,))


def remove_cart_item(user_id, position):
    """Удаляет товар по номеру в корзине (с 1). Возвращает True/False."""
    with connect() as conn:
        rows = conn.execute(
            'SELECT id FROM cart_items WHERE user_id = ? ORDER BY id', (user_id,)
        ).fetchall()
        if position < 1 or position > len(rows):
            return False
        conn.execute('DELETE FROM cart_items WHERE id = ?', (rows[position - 1][0],))
        return True


def update_cart_size(user_id, position, new_size):
    """Меняет размер товара по номеру в корзине (с 1).
    Если в корзине уже есть этот артикул с новым размером — количества суммируются.
    Возвращает 'ok', 'bad_position' или 'no_size' (у товара нет размера)."""
    with connect() as conn:
        rows = conn.execute(
            'SELECT id, article, size, quantity FROM cart_items WHERE user_id = ? ORDER BY id',
            (user_id,)
        ).fetchall()
        if position < 1 or position > len(rows):
            return 'bad_position'
        item_id, article, size, qty = rows[position - 1]
        if not size or size == '-':
            return 'no_size'

        other = conn.execute(
            'SELECT id FROM cart_items '
            'WHERE user_id = ? AND article = ? AND LOWER(size) = LOWER(?) AND id != ?',
            (user_id, article, new_size, item_id)
        ).fetchone()
        if other:
            conn.execute('UPDATE cart_items SET quantity = quantity + ? WHERE id = ?',
                         (qty, other[0]))
            conn.execute('DELETE FROM cart_items WHERE id = ?', (item_id,))
        else:
            conn.execute('UPDATE cart_items SET size = ? WHERE id = ?', (new_size, item_id))
        return 'ok'


def update_cart_quantity(user_id, position, new_qty):
    """Меняет количество товара по номеру в корзине (с 1). Возвращает True/False."""
    with connect() as conn:
        rows = conn.execute(
            'SELECT id FROM cart_items WHERE user_id = ? ORDER BY id', (user_id,)
        ).fetchall()
        if position < 1 or position > len(rows):
            return False
        conn.execute('UPDATE cart_items SET quantity = ? WHERE id = ?',
                     (new_qty, rows[position - 1][0]))
        return True


# ---------------------------------------------------------------- заказы

def checkout_cart(user_id, user_name, payment_photo):
    """Превращает корзину в заказ одной транзакцией и очищает корзину.
    Цены фиксируются в заказе. Возвращает order_id или None, если корзина пуста."""
    with connect() as conn:
        cursor = conn.cursor()
        items = cursor.execute('''
            SELECT c.article, c.size, c.quantity, p.price
            FROM cart_items c
            LEFT JOIN products p ON p.article = c.article
            WHERE c.user_id = ?
            ORDER BY c.id
        ''', (user_id,)).fetchall()
        if not items:
            return None

        cursor.execute(
            'INSERT INTO orders (user_id, status, user_name, payment_photo) VALUES (?, ?, ?, ?)',
            (user_id, STATUS_PROCESSING, user_name, payment_photo)
        )
        order_id = cursor.lastrowid
        for article, size, qty, price in items:
            cursor.execute(
                'INSERT INTO order_items (order_id, article, size, quantity, price) '
                'VALUES (?, ?, ?, ?, ?)',
                (order_id, article, size, qty, price or 0)
            )
        cursor.execute('DELETE FROM cart_items WHERE user_id = ?', (user_id,))
        cursor.execute('DELETE FROM carts WHERE user_id = ?', (user_id,))
        return order_id


def get_order(order_id):
    """Возвращает словарь с данными заказа или None."""
    with connect() as conn:
        row = conn.execute(
            'SELECT id, user_id, status, user_name, payment_photo FROM orders WHERE id = ?',
            (order_id,)
        ).fetchone()
    if not row:
        return None
    return {'id': row[0], 'user_id': row[1], 'status': row[2],
            'user_name': row[3], 'payment_photo': row[4]}


def _order_items(cursor, order_id):
    rows = cursor.execute('''
        SELECT p.name, i.article, i.size, i.quantity, COALESCE(i.price, p.price, 0)
        FROM order_items i
        LEFT JOIN products p ON p.article = i.article
        WHERE i.order_id = ?
        ORDER BY i.id
    ''', (order_id,)).fetchall()
    return [{'name': r[0], 'article': r[1], 'size': r[2], 'quantity': r[3], 'price': r[4]}
            for r in rows]


def get_order_items(order_id):
    """[{'name', 'article', 'size', 'quantity', 'price'}, ...]. name = None, если товара нет в каталоге."""
    with connect() as conn:
        return _order_items(conn.cursor(), order_id)


def get_order_total(order_id):
    """Итоговая сумма заказа по ценам, зафиксированным при оформлении."""
    return sum(i['price'] * i['quantity'] for i in get_order_items(order_id))


def set_status(order_id, status):
    """Меняет статус заказа. Возвращает True, если заказ найден."""
    with connect() as conn:
        cursor = conn.execute('UPDATE orders SET status = ? WHERE id = ?', (status, order_id))
        return cursor.rowcount > 0


def get_orders_for_export():
    """Все заказы (новые сверху) со списком товаров и суммой:
    [{'id', 'user_id', 'user_name', 'status', 'payment_photo', 'items', 'total'}, ...]"""
    with connect() as conn:
        cursor = conn.cursor()
        orders = cursor.execute(
            'SELECT id, user_id, user_name, status, payment_photo FROM orders ORDER BY id DESC'
        ).fetchall()
        result = []
        for order_id, user_id, user_name, status, photo in orders:
            items = _order_items(cursor, order_id)
            result.append({
                'id': order_id,
                'user_id': user_id,
                'user_name': user_name,
                'status': status,
                'payment_photo': photo,
                'items': items,
                'total': sum(i['price'] * i['quantity'] for i in items),
            })
        return result