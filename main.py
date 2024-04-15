import asyncio
import os
import sqlite3
import time
import uuid
import zipfile
from zipfile import ZipFile

import openpyxl
from aiogram.utils.exceptions import BadRequest
from aiogram import types, Dispatcher, Bot
from aiogram.contrib.middlewares.logging import LoggingMiddleware
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import state
from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.types import ParseMode, InlineKeyboardButton, InlineKeyboardMarkup
from aiogram.utils import executor
import pytz
from openpyxl.styles import Alignment, Font

import config
import keyboards
import register
import sms
from datetime import datetime
from aiogram.dispatcher import FSMContext

bot = Bot(token=config.BOT_TOKEN, parse_mode=ParseMode.HTML)
dp = Dispatcher(bot, storage=MemoryStorage())
dp.middleware.setup(LoggingMiddleware())
conn = sqlite3.connect('database.db')
cursor = conn.cursor()


class AddResult(StatesGroup):
    event_id = State()
    result = State()
    points = State()


class CreateEvent(StatesGroup):
    name = State()
    description = State()
    level = State()
    file_id = State()
    date = State()


################################ START ################################
@dp.message_handler(commands=['start'])
async def handler_start(message: types.Message, state: FSMContext):
    verify = cursor.execute("SELECT * FROM users WHERE user_id = ?", (message.from_user.id,)).fetchone()
    if verify:
        await state.finish()
        await bot.send_message(chat_id=message.from_user.id,
                               text=sms.my_profile_in_main_menu(message.from_user.id),
                               reply_markup=await keyboards.main_keyboard(message.from_user.id))


# ROLE 1 CREATE MP
@dp.callback_query_handler(text='create_event')
async def create_event(call: types.CallbackQuery):
    await call.message.answer('Введите название мероприятия:')
    await CreateEvent.name.set()


@dp.message_handler(state=CreateEvent.name)
async def set_description(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['name'] = message.text
    await message.answer('Введите описание мероприятия:')
    await CreateEvent.next()


@dp.message_handler(state=CreateEvent.description)
async def set_file_id(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['description'] = message.text
    await message.answer('Введите уровень мероприятия (Региональный, муниципальный, государственный и т.д.):')
    await CreateEvent.next()


@dp.message_handler(state=CreateEvent.level)
async def set_level(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['level'] = message.text
    await message.answer('Отправьте файл мероприятия (если есть, если нет, то напишите "Нет"):')
    await CreateEvent.next()


@dp.message_handler(state=CreateEvent.file_id, content_types=[
    types.ContentType.DOCUMENT,
    types.ContentType.PHOTO,
    types.ContentType.TEXT
])
async def set_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        if message.content_type == types.ContentType.TEXT:
            data['file_id'] = message.text
        elif message.document:
            file_id = message.document.file_id

            file_extension = message.document.file_name.split('.')[-1]
            file_path = os.path.join('docs', f'{file_id}.{file_extension}')

            data['file_id'] = f"{file_id}.{file_extension}"

            await message.document.download(destination=file_path)
        elif message.photo:
            file_id = message.photo[-1].file_id

            unique_filename = str(uuid.uuid4())

            photo_info = await bot.get_file(file_id)

            file_extension = photo_info.file_path.split('.')[-1]

            file_path = os.path.join('docs', f'{unique_filename}.{file_extension}')

            await bot.download_file(photo_info.file_path, destination=file_path)

            data['file_id'] = f"{unique_filename}.{file_extension}"

    await message.answer('Введите дату мероприятия в формате ДД-ММ-ГГГГ:')
    await CreateEvent.next()


@dp.message_handler(state=CreateEvent.date)
async def save_event(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        try:
            event_date = datetime.strptime(message.text, '%d-%m-%Y')
            data['event_date'] = event_date.strftime('%d-%m-%Y')
        except ValueError:
            await message.answer('Неверный формат даты. Попробуйте снова.')
            return

        cursor.execute('''
            INSERT INTO events (event_name, event_description, event_level, file_id, event_date, responsible_user_id)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            data['name'], data['description'], data['level'], data['file_id'], data['event_date'],
            message.from_user.id))
        conn.commit()
    await message.answer('Мероприятие сохранено!')
    await state.finish()


# ROLE 1 check_my_stats
@dp.callback_query_handler(text='check_my_stats')
async def check_my_stats(call: types.CallbackQuery):
    user_id = call.from_user.id

    events_query = """
        SELECT events.event_id, events.event_name, events.event_description, events.event_level, events.file_id, events.event_date, users.full_name
        FROM events
        JOIN users ON events.responsible_user_id = users.user_id
        WHERE events.responsible_user_id = ?
    """
    events = cursor.execute(events_query, (user_id,)).fetchall()

    if not events:
        await call.message.answer("У вас нет мероприятий.")
        return

    await call.message.answer('Данные загружаются, пожалуйста, подождите...')

    zip_filename = f"events_{user_id}.zip"
    filename = f"events_{user_id}.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Мероприятия"
    bold_font = openpyxl.styles.Font(bold=True)
    ws.append(
        ["ID", "Название мероприятия", "Описание мероприятия", "Уровень мероприятия", "Файл", "Дата мероприятия",
         "Ответственный за мероприятие"])
    for cell in ws[1]:
        cell.font = bold_font

    for event in events:
        event_list = list(event)
        if event_list[4] and (event_list[4].lower() != 'нет'):
            file_path = f'files/{event_list[4]}'
            event_list[4] = f'=HYPERLINK("{file_path}", "Есть файлы (Кликабельно)")'
        ws.append(event_list)

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    results_query = """
        SELECT events.event_name, events.event_description, results.result, results.points
        FROM events
        JOIN results ON events.event_id = results.event_id
        WHERE events.responsible_user_id = ?
    """
    results = cursor.execute(results_query, (user_id,)).fetchall()

    ws_results = wb.create_sheet(title="Результаты мероприятий")
    ws_results.append(["Название мероприятия", "Описание мероприятия", "Результат", "Баллы"])
    for result in results:
        ws_results.append(result)

    for row in ws_results.iter_rows():
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    wb.save(filename)

    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        zipf.writestr('Мои мероприятия/', '')

        zipf.write(filename, os.path.join('Мои мероприятия', filename))

        zipf.writestr(os.path.join('Мои мероприятия', 'files/'), '')

        docs_folder = 'docs'
        for root, dirs, files in os.walk(docs_folder):
            for file in files:
                file_id, _ = os.path.splitext(file)
                file_path = os.path.join(root, file)
                zipf.write(file_path,
                           arcname=os.path.join('Мои мероприятия', 'files', os.path.relpath(file_path, docs_folder)))

    with open(zip_filename, "rb") as file:
        await call.message.answer_document(file)

    os.remove(filename)
    os.remove(zip_filename)

# ROLE 1 add_result
@dp.callback_query_handler(text='add_result')
async def add_result(call: types.CallbackQuery):
    await call.message.answer('Для добавления результата введите ID мероприятия:')
    await AddResult.event_id.set()


@dp.message_handler(state=AddResult.event_id)
async def save_result(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        try:
            event_id = int(message.text)
            data['event_id'] = event_id
        except ValueError:
            await message.answer('Неверный формат ID. Попробуйте снова.')
            return
        await message.answer('Введите результат:')
        await AddResult.next()


@dp.message_handler(state=AddResult.result)
async def save_result(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        result = message.text
        data['result'] = result
    await message.answer('Введите баллы, назначаемые за результат: ')
    await AddResult.next()


@dp.message_handler(state=AddResult.points)
async def save_result(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        points = message.text
        data['points'] = points

    cursor.execute("INSERT INTO results (event_id, result, points) VALUES (?,?,?)",
                   (data['event_id'], data['result'], data['points']))

    conn.commit()

    await message.answer("Результат успешно добавлен!")
    await state.finish()


if __name__ == '__main__':
    dp.register_message_handler(handler_start, commands=['start'], state='*')
    register.all_callback(dp)
    executor.start_polling(dp, skip_updates=True)
