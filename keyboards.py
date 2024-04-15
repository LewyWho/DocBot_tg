import sqlite3

from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

conn = sqlite3.connect('database.db')
cursor = conn.cursor()


async def main_keyboard(user_id):

    role = cursor.execute("SELECT role FROM users WHERE user_id =?", (user_id,)).fetchone()[0]

    keyboard = InlineKeyboardMarkup()

    if role == 1:
        button_create_event = InlineKeyboardButton("Создать мероприятие", callback_data="create_event")
        button_check_my_stats = InlineKeyboardButton("Проверить статистику", callback_data="check_my_stats")
        button_add_result = InlineKeyboardButton("Добавить результат", callback_data="add_result")
        keyboard.add(button_create_event)
        keyboard.add(button_check_my_stats)
        keyboard.add(button_add_result)
    elif role == 2:
        button_check_stats_staff = InlineKeyboardButton("Проверить статистику", callback_data="check_stats_staff")
        button_create_report = InlineKeyboardButton("Создать отчет сотрудников", callback_data="create_report_staff")
        keyboard.add(button_check_stats_staff)
        keyboard.add(button_create_report)
    elif role == 3:
        button_create_report = InlineKeyboardButton("Создать отчет сотрудников", callback_data="create_report_staff")
        keyboard.add(button_create_report)
    elif role == 4:
        button_add_new_staff = InlineKeyboardButton("Добавить нового сотрудника", callback_data="add_new_staff")
        button_del_staff = InlineKeyboardButton("Удалить сотрудника", callback_data="del_staff")
        button_change_staff = InlineKeyboardButton("Изменить сотрудника", callback_data="change_staff")
        keyboard.add(button_add_new_staff)
        keyboard.add(button_del_staff)
        keyboard.add(button_change_staff)

    return keyboard


