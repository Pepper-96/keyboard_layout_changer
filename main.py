import threading
import sys
import os
import pystray
from pystray import MenuItem as Item, Menu
from PIL import Image
import time
import keyboard
import pywintypes
import win32clipboard as cb
import win32con
import win32event
import win32api
import winerror
from pynput.keyboard import Key, Controller


# Имя mutex для защиты от запуска нескольких экземпляров
MUTEX_NAME = "keyboard_layout_changer_single_instance_mutex"

# Наборы букв
RU_CHARS = set("йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ")
EN_CHARS = set("qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM")

# Таблица замен
RU = "йцукенгшщзхъфывапролджэячсмитьбюё,.\"№;:?"
EN = "qwertyuiop[]asdfghjkl;'zxcvbnm,.`?/@#$^&"
EN_UPPER = "QWERTYUIOP{}ASDFGHJKL:\"ZXCVBNM<>~?/@#$^&"
ru_to_en = str.maketrans(RU + RU.upper(), EN + EN_UPPER)
en_to_ru = str.maketrans(EN + EN_UPPER, RU + RU.upper())

# Словарь подстановки клавиш
HOTKEY_PH = {
    "ctrl": Key.ctrl,
    "lctrl": Key.ctrl_l,
    "rctrl": Key.ctrl_r,
    "alt": Key.alt,
    "lalt": Key.alt_l,
    "ralt": Key.alt_r,
    "shift": Key.shift,
    "lshift": Key.shift_l,
    "rshift": Key.shift_r,
    "space": Key.space,
    "backspace": Key.backspace,
    "enter": Key.enter,
    "end": Key.end,
}

def ensure_single_instance() -> int:
    """Запуск единственного экземпляра"""
    # создаём / открываем именованный mutex
    handle = win32event.CreateMutex(None, False, MUTEX_NAME)
    # проверяем, не существует ли он уже
    last_error = win32api.GetLastError()
    if last_error == winerror.ERROR_ALREADY_EXISTS:
        # другой экземпляр уже создал mutex -> выходим
        # можно сначала показать messagebox, если хочешь
        sys.exit(0)
    return handle

def resource_path(relative_path: str) -> str:
    """Путь к ресурсу, работает и в dev, и в exe."""
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_clipboard_text(retries: int = 5, delay: float = 0.05) -> str:
    """
    Чтение буфера

    Получает текст из буфера обмена

    Args:
        retries(int): количество попыток чтения буфера
        delay(float): задержка между попытками в секундах

    Returns:
        str: текст из буфера
    """
    for i in range(retries):
        try:
            cb.OpenClipboard()
            try:
                if cb.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                    data = cb.GetClipboardData(win32con.CF_UNICODETEXT)
                elif cb.IsClipboardFormatAvailable(win32con.CF_TEXT):
                    data = cb.GetClipboardData(win32con.CF_TEXT)
                else:
                    data = ""
            finally:
                cb.CloseClipboard()
            return data
        except pywintypes.error as e:
            # 5 = Access denied -> буфер занят другим процессом
            if e.args[0] == 5:
                time.sleep(delay)
                continue
            else:
                raise
    # Не смогли открыть после всех попыток
    return ""

def set_clipboard_text(text: str, retries: int = 5, delay: float = 0.05) -> None:
    """
    Перезапись буфера обмена

    Args:
        text(str): текст, копируемый в буфер
        retries(int): количество попыток записи в буфер
        delay(float): задержка между попытками в секундах
    """
    for i in range(retries):
        try:
            cb.OpenClipboard()
            try:
                cb.EmptyClipboard()
                cb.SetClipboardText(text, win32con.CF_UNICODETEXT)
            finally:
                cb.CloseClipboard()
            return
        except pywintypes.error as e:
            if e.args[0] == 5:
                time.sleep(delay)
                continue
            else:
                raise

def detect_direction(text: str) -> str:
    """
    Определение требуемой раскладки по соотношению символов одной раскладки другой

    Если кириллических букв больше, чем букв на латинице, значит, скорее всего текст должен был быть
    написан на латинице.

    Если букв на латинице больше, значит текст должен был быть написан на кириллице.
    Иначе оставляем текст нетронутым.

    Args:
        text(str): выделенный текст

    Returns:
        str: требуемая раскладка
    """
    ru_count = sum(1 for ch in text if ch in RU_CHARS)
    en_count = sum(1 for ch in text if ch in EN_CHARS)

    if en_count > ru_count:
        return "ru"
    elif ru_count > en_count:
        return "en"
    else:
        # Непонятно (цифры, символы, смесь) — можно оставить как есть
        return "none"

def fix_layout(text: str) -> str:
    """
    Замена символов на символы другой раскладки

    Args:
        text(str): изначальный текст

    Returns:
        str: текст после замены символов
    """
    direction = detect_direction(text)
    if direction == "ru":
        return text.translate(en_to_ru)
    elif direction == "en":
        return text.translate(ru_to_en)
    else:
        return text

def on_hotkey(hotkey: str) -> None:
    """
    Отклик на сочетание клавиш

    Args:
        hotkey(str): сочетание клавиш

    Вызывает функцию замены символов одной раскладки на другую
    """
    # Сохраняем старый буфер в переменную и очищаем буфер
    old = get_clipboard_text()
    set_clipboard_text("")
    # Создаём переменную для клавиатуры
    kb = Controller()
    # Отпускаем сочетание клавиш для вызова функции
    for key in hotkey.split("+"):
        if key in HOTKEY_PH.keys():
            kb.release(HOTKEY_PH[key])
        else:
            kb.release(key)
    # Выполняем копирование выделенного текста через Ctrl+C
    kb.press(Key.ctrl)
    kb.press('c')
    # Отпускаем сочетание клавиш для копирования
    kb.release('c')
    kb.release(Key.ctrl)
    # Задержка для обновления буфера
    time.sleep(0.05)
    # Сохраняем текст из буфера в переменную
    selected = get_clipboard_text()
    if not selected:
        # Ничего не выделено или не текст
        set_clipboard_text(old)
        return

    # Исправляем символы на символы другой раскладки
    fixed = fix_layout(selected)
    # Заменяем буфер исправленным текстом
    set_clipboard_text(fixed)
    # Выполняем вставку через Ctrl+V
    kb.press(Key.ctrl)
    kb.press('v')
    # Отпускаем сочетание клавиш для вставки
    kb.release('v')
    kb.release(Key.ctrl)
    # Задержка для обновления буфера
    time.sleep(0.05)
    # Восстановление старого буфера
    set_clipboard_text(old)

def start_hotkeys() -> None:
    """Добавление сочетания клавиш"""
    hotkey = "ctrl+shift+q"
    keyboard.add_hotkey(hotkey, lambda: on_hotkey(hotkey))
    keyboard.wait()  # блокирует поток, поэтому запускаем в отдельном

def on_exit(icon, item) -> None:
    """Останавливает иконку и завершает процесс"""
    icon.visible = False
    icon.stop()
    sys.exit(0)

def run_tray() -> None:
    """Запуск программы в трее"""
    icon_path = resource_path("icon.png")
    icon = pystray.Icon(
        "LayoutFixer",
        icon=Image.open(icon_path),
        title="Сменщик раскладки", # Layout fixer
        menu=Menu(
            Item("Завершить", on_exit)   # пункт меню
        )
    )
    icon.run()  # блокирует поток

if __name__ == "__main__":
    mutex = ensure_single_instance()
    # Хоткеи в отдельном потоке
    t = threading.Thread(target=start_hotkeys, daemon=True)
    t.start()
    # Главный поток – иконка в трее
    run_tray()