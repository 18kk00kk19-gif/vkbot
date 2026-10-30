from vk_api.keyboard import VkKeyboard, VkKeyboardColor


def get_main_keyboard():
    """Начальная."""
    keyboard = VkKeyboard(one_time=True)
    keyboard.add_openlink_button('Каталог', link='https://vk.ru/market-242321155?screen=group')
    keyboard.add_line()
    keyboard.add_button('Корзина', color=VkKeyboardColor.PRIMARY)
    keyboard.add_line()
    keyboard.add_button('Оплатить Заказ', color=VkKeyboardColor.POSITIVE)
    return keyboard


def get_cart_keyboard():
    """Меню корзины."""
    keyboard = VkKeyboard(one_time=False)
    keyboard.add_button('Очистить корзину', color=VkKeyboardColor.NEGATIVE)
    keyboard.add_button('Добавить товар', color=VkKeyboardColor.POSITIVE)
    keyboard.add_line()
    keyboard.add_button('Изменить корзину', color=VkKeyboardColor.PRIMARY)
    keyboard.add_line()
    keyboard.add_button('Назад', color=VkKeyboardColor.PRIMARY)
    return keyboard


def get_item_input_keyboard():
    """Клавиатура во время ввода товара."""
    keyboard = VkKeyboard(one_time=False)
    keyboard.add_button('Отмена', color=VkKeyboardColor.NEGATIVE)
    return keyboard


def get_cart_edit_keyboard(show_change_size=False):
    """Клавиатура для редактирования корзины. «Изменить размер» — только если в корзине есть товары с размером."""
    keyboard = VkKeyboard(one_time=False)
    keyboard.add_button('Удалить товар', color=VkKeyboardColor.NEGATIVE)
    keyboard.add_line()
    keyboard.add_button('Изменить количество')
    if show_change_size:
        keyboard.add_line()
        keyboard.add_button('Изменить размер')
    keyboard.add_line()
    keyboard.add_button('Назад')
    return keyboard


def get_payment_keyboard():
    """Клавиатура для выбора способа оплаты."""
    keyboard = VkKeyboard(one_time=True)
    keyboard.add_button('Оплатить по реквизитам', color=VkKeyboardColor.POSITIVE)
    keyboard.add_line()
    keyboard.add_button('Оплатить через СБП', color=VkKeyboardColor.POSITIVE)
    keyboard.add_line()
    keyboard.add_button('Назад в меню', color=VkKeyboardColor.SECONDARY)
    return keyboard


def get_payment_confirmation_keyboard():
    """Клавиатура для подтверждения оплаты."""
    keyboard = VkKeyboard(one_time=True)
    keyboard.add_button('Я оплатил(а)', color=VkKeyboardColor.POSITIVE)
    keyboard.add_button('Отмена', color=VkKeyboardColor.NEGATIVE)
    return keyboard


def get_market_confirm_keyboard(article):
    """Под карточкой товара, присланного из VK Market: «В корзину» / «Отмена»."""
    keyboard = VkKeyboard(one_time=True)
    keyboard.add_button('В корзину', color=VkKeyboardColor.POSITIVE, payload={'article': article})
    keyboard.add_button('Отмена', color=VkKeyboardColor.NEGATIVE)
    return keyboard