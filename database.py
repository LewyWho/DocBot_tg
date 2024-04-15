import sqlite3

conn = sqlite3.connect('database.db')
cursor = conn.cursor()

cursor.execute('''
    CREATE TABLE IF NOT EXISTS events (
        event_id INTEGER PRIMARY KEY, -- Идентификатор мероприятия
        event_name TEXT, -- Название мероприятия
        event_description TEXT, -- Описание мероприятия
        event_level TEXT, -- Уровень мероприятия
        file_id TEXT, -- Идентификатор прикрепленного файла
        event_date date, -- Дате мероприятия
        responsible_user_id INTEGER -- Идентификатор ответственного за мероприятие
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        user_id INTEGER PRIMARY KEY, -- Идентификатор пользователя
        full_name TEXT, -- ФИО пользователя
        department TEXT, -- Название кафедры
        role INTEGER, -- Роль пользователя: 1 - Сотрудник, 2 - Администрация НГИЭУ, 3 - Аналитик, 4 - Администратор бота
        additional_info TEXT -- Дополнительная информация о пользователе
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS results (
        result_id INTEGER PRIMARY KEY, -- Идентификатор результата
        event_id INTEGER, -- Идентификатор мероприятия
        result TEXT, -- Результат (1 место, 2 место, 3 место, участие и т.д.)
        points INTEGER -- Баллы, назначаемые за результат
    )
''')

# Сохранение изменений и закрытие соединения
conn.commit()