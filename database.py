from main import cursor
from main import conn

cursor.execute('''
    CREATE TABLE IF NOT EXISTS events (
        event_id INTEGER PRIMARY KEY, -- Идентификатор мероприятия
        event_name TEXT, -- Название мероприятия
        event_description TEXT, -- Описание мероприятия
        file_id TEXT, -- Идентификатор прикрепленного файла
        event_date date, -- Дате мероприятия
        responsible_user_id INTEGER -- Идентификатор ответственного за мероприятие
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        user_id INTEGER PRIMARY KEY, -- Идентификатор пользователя
        full_name TEXT, -- ФИО пользователя
        role INTEGER, -- Роль пользователя: 1 - Сотрудник, 2 - Администрация НГИЭУ, 3 - Аналитик, 4 - Администратор бота
        additional_info TEXT -- Дополнительная информация о пользователе
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS results (
        result_id INTEGER PRIMARY KEY, -- Идентификатор результата
        event_id INTEGER, -- Идентификатор мероприятия
        level TEXT, -- Уровень мероприятия (муниципальный, региональный и т.д.)
        result TEXT, -- Результат (1 место, 2 место, 3 место, участие и т.д.)
        points INTEGER -- Баллы, назначаемые за результат
    )
''')

# Сохранение изменений и закрытие соединения
conn.commit()