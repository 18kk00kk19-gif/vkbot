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
        ]:
            try:
                cursor.execute(col_sql)
            except sqlite3.OperationalError:
                pass  # колонка уже существует

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


def get_product_by_name(name):
    """Ищет товар по названию (без учёта регистра и пробелов). Возвращает (article, price, has_size)."""
    with connect() as conn:
        return conn.execute(
            'SELECT article, price, has_size FROM products '
            'WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))', (name,)
        ).fetchone()


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