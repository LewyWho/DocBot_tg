from main import cursor
from main import conn


def my_profile_in_main_menu(user_id):
    role = cursor.execute("SELECT role FROM users WHERE user_id =?", (user_id,)).fetchone()[0]
    if role == 1:
        role = 'Сотрудник'
    elif role == 2:
        role = 'Руководство'
    elif role == 3:
        role = 'Аналитик'
    elif role == 4:
        role = 'Администратор бота'
    return f"""
Ваш ID: {user_id}
Ваш роль: {role}"""
