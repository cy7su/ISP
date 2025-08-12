from __future__ import annotations

import asyncio
import os
import re
from datetime import datetime, timedelta
from typing import Any

import httpx
import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
from aiogram.utils.keyboard import InlineKeyboardBuilder
from html2image import Html2Image
import io
from PIL import Image

TABLE_URL = "https://ppk.sstu.ru/doc/rasp/Горького,%209/stud.xls"
RESULT_FILE = "ИСП-11.xls"
TEMP_FILE = "temp.xls"
HTML_FILE = "ИСП-11.html"
DAY_HTML_FILE = "day_ИСП-11.html"

bot = Bot(token='----------------------------')
dp = Dispatcher()

def get_main_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.add(InlineKeyboardButton(text="📅 Расписание на сегодня", callback_data="today"))
    builder.add(InlineKeyboardButton(text="📊 Расписание на неделю", callback_data="week"))
    builder.add(InlineKeyboardButton(text="🔄 Обновить расписание", callback_data="update"))
    builder.add(InlineKeyboardButton(text="📄 HTML расписание", callback_data="html"))
    builder.adjust(2)
    return builder.as_markup()

def get_back_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.add(InlineKeyboardButton(text="⬅️ Назад", callback_data="back"))
    return builder.as_markup()

async def find_last_bot_message(chat_id: int) -> types.Message | None:
    """Ищет последнее сообщение от бота в чате"""
    try:
        # Получаем последние сообщения в чате
        messages = await bot.get_chat_history(chat_id, limit=10)
        for message in messages:
            if message.from_user.id == bot.id:
                return message
    except Exception:
        pass
    return None

async def send_or_edit_message(chat_id: int, text: str, reply_markup=None) -> types.Message:
    """Отправляет новое сообщение или редактирует существующее"""
    last_message = await find_last_bot_message(chat_id)
    
    if last_message:
        # Редактируем существующее сообщение
        await last_message.edit_text(text, reply_markup=reply_markup)
        return last_message
    else:
        # Отправляем новое сообщение
        return await bot.send_message(chat_id, text, reply_markup=reply_markup)

@dp.message(Command("start"))
async def start(message: types.Message) -> None:
    sent_message = await message.answer(
        "Привет! Я бот для расписания ИСП-11.\n"
        "Используй кнопки внизу экрана для управления:",
        reply_markup=get_main_keyboard()
    )
    
    # Закрепляем сообщение
    try:
        await sent_message.pin()
    except Exception as e:
        print(f"Не удалось закрепить сообщение: {e}")

@dp.callback_query()
async def handle_callback(callback: types.CallbackQuery) -> None:
    """Обрабатываем callback'и от inline кнопок"""
    if callback.data == "today":
        await get_today_schedule(callback.message)
    elif callback.data == "week":
        await get_week_schedule(callback.message)
    elif callback.data == "update":
        await download_schedule(callback.message)
    elif callback.data == "html":
        await send_html_file(callback.message)
    elif callback.data == "back":
        await show_main_menu(callback.message)
    
    await callback.answer()

async def show_main_menu(message: types.Message) -> None:
    """Показывает главное меню"""
    await send_or_edit_message(
        message.chat.id,
        "Привет! Я бот для расписания ИСП-11.\n"
        "Используй кнопки внизу экрана для управления:",
        get_main_keyboard()
    )

async def download_schedule(message: types.Message) -> None:
    try:
        # Отправляем или редактируем сообщение о загрузке
        status_message = await send_or_edit_message(message.chat.id, "⏳ Загружаю новое расписание...", get_back_keyboard())
        
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.get(TABLE_URL)
            response.raise_for_status()

        with open(TEMP_FILE, 'wb') as f:
            f.write(response.content)
        
        df = pd.read_excel(TEMP_FILE, engine='xlrd')
        
        # Ищем начало ИСП-11
        target_row = None
        for row in range(df.shape[0]):
            cell_value = str(df.iloc[row, 0])
            if "Группа - ИСП-11" in cell_value:
                target_row = row
                break

        if target_row is None:
            await status_message.edit_text("❌ Группа ИСП-11 не найдена в таблице.", reply_markup=get_back_keyboard())
            return

        # Ищем конец ИСП-11 (до следующей группы)
        end_row = None
        for row in range(target_row + 1, df.shape[0]):
            cell_value = str(df.iloc[row, 0])
            if "Группа -" in cell_value and "ИСП-11" not in cell_value:
                end_row = row
                break
        
        if end_row is None:
            end_row = df.shape[0]
            
        # Извлекаем только данные ИСП-11
        extracted_data = df.iloc[target_row + 1:end_row, :]
        
        # Заменяем NaN на пробелы
        extracted_data = extracted_data.fillna(' ')
        
        extracted_data.to_excel(RESULT_FILE, index=False, header=False, engine='openpyxl')

        # Сразу конвертируем в HTML (полное расписание)
        await convert_to_html_and_save(extracted_data)
        
        # Создаем HTML для дня
        await create_day_html(extracted_data)
        
        # Удаляем старое сообщение
        await status_message.delete()
        
        # Отправляем новое сообщение с результатом
        await bot.send_message(
            message.chat.id,
            "✅ Расписание успешно обновлено и конвертировано в HTML!",
            reply_markup=get_main_keyboard()
        )

    except httpx.HTTPError as e:
        await send_or_edit_message(message.chat.id, f"❌ Ошибка загрузки: {e}", get_back_keyboard())
    except Exception as e:
        await send_or_edit_message(message.chat.id, f"❌ Произошла ошибка: {e}", get_back_keyboard())
    finally:
        if os.path.exists(TEMP_FILE):
            os.remove(TEMP_FILE)

async def convert_to_html_and_save(df: pd.DataFrame) -> None:
    # Фильтруем пустые строки и убираем субботу
    filtered_data = []
    for row_idx in range(df.shape[0]):
        row_data = df.iloc[row_idx, :]
        # Проверяем есть ли данные в строке (исключая первые 2 колонки - номер и время)
        has_data = any(str(cell).strip() not in ['nan', ' ', ''] for cell in row_data[2:7])  # Пн-Пт (колонки 2-7)
        if has_data:
            filtered_data.append(row_data)
    
    # Создаем новый DataFrame только с непустыми строками
    filtered_df = pd.DataFrame(filtered_data)
    
    # Убираем строку "Дисциплина, вид занятия, преподаватель"
    if filtered_df.shape[0] > 1:
        # Удаляем вторую строку (индекс 1)
        filtered_df = filtered_df.drop(filtered_df.index[1])
    
    # Создаем красивый HTML с полным расписанием (только Пн-Пт)
    html_content = filtered_df.iloc[:, :7].to_html(index=False, header=False, classes='table table-striped')
    
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Расписание ИСП-11</title>
        <style>
            body {{ 
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans', Helvetica, Arial, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #0d1117;
                color: #c9d1d9;
                line-height: 1.6;
                min-height: 100vh;
            }}
            .container {{
                width: 1380px;
                margin: 0 auto;
                background-color: #161b22;
                border: 1px solid #30363d;
                border-radius: 6px;
                box-shadow: 0 8px 24px rgba(1, 4, 9, 0.12);
                overflow: hidden;
                display: flex;
                flex-direction: column;
            }}
            .table {{ 
                border-collapse: collapse; 
                width: 100%; 
                background-color: #161b22;
                font-size: 16px;
            }}
            .table td, .table th {{ 
                border: 1px solid #30363d; 
                padding: 20px 16px; 
                text-align: left; 
                vertical-align: top;
            }}
            .table th {{ 
                background-color: #21262d;
                color: #f0f6fc;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                font-size: 12px;
            }}
            .table tr:nth-child(even) {{ 
                background-color: #1c2128; 
            }}
            .table tr:nth-child(odd) {{
                background-color: #161b22;
            }}
            .table tr:hover {{
                background-color: #21262d;
                transition: background-color 0.2s ease;
            }}
            .table tr:first-child {{
                background-color: #1f6feb !important;
                color: #f0f6fc;
                font-weight: 600;
                font-size: 14px;
            }}
            .table tr:nth-child(2) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(3) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(4) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(5) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(6) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(7) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(8) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(9) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(10) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(11) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(12) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(13) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(14) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(15) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(16) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(17) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(18) {{
                background-color: #161b22;
            }}
            .table tr:nth-child(19) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(20) {{
                background-color: #161b22;
            }}
            .update-time {{
                text-align: center;
                color: #8b949e;
                font-style: italic;
                background-color: #21262d;
                border-top: 1px solid #30363d;
                font-size: 14px;
            }}
            .time-column {{
                background-color: #21262d !important;
                color: #58a6ff;
                font-weight: 600;
                font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
            }}
            .empty-cell {{
                color: #484f58;
                font-style: italic;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            {html_content}
            <div class="update-time">
                Обновлено: {datetime.now().strftime('%d.%m.%Y %H:%M')}
            </div>
        </div>
    </body>
    </html>
    """
    
    # Сохраняем HTML файл
    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(full_html)

async def create_day_html(df: pd.DataFrame) -> None:
    """Создаем отдельный HTML файл для расписания на день в виде таблицы"""
    # Умно определяем дату
    target_date, reason = get_smart_date_for_schedule()
    weekday_names = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница"]
    
    # Ищем расписание на целевую дату
    day_col = None
    for col in range(2, df.shape[1]):
        cell_value = str(df.iloc[0, col])
        if weekday_names[target_date.weekday()] in cell_value and target_date.strftime('%d.%m.%Y') in cell_value:
            day_col = col
            break
    
    # Если не найдено, ищем последнее доступное
    if day_col is None:
        for col in range(2, df.shape[1]):
            cell_value = str(df.iloc[0, col])
            if weekday_names[target_date.weekday()] in cell_value:
                date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', cell_value)
                if date_match:
                    found_date = datetime.strptime(date_match.group(1), '%d.%m.%Y').date()
                    if found_date <= target_date:
                        day_col = col
                        target_date = found_date
                        reason = f"последнее доступное ({found_date.strftime('%d.%m.%Y')})"
                        break
    
    if day_col is None:
        return

    current_weekday_name = weekday_names[target_date.weekday()]
    
    # Собираем пары для дня
    pairs = []
    first_pair_time = None
    
    for row in range(1, df.shape[0]):
        time_slot = str(df.iloc[row, 1]).strip()
        if not time_slot or time_slot == 'nan' or time_slot == ' ':
            continue
            
        cell_value = str(df.iloc[row, day_col]).strip()
        if not cell_value or cell_value == 'nan' or cell_value == ' ':
            continue
            
        if first_pair_time is None:
            first_pair_time = time_slot.split('-')[0].strip()
            
        lines = cell_value.split('\n')
        subject = lines[0].strip() if lines else ""
        teacher = ""
        classroom = ""
        
        if len(lines) > 1:
            second_line = lines[1].strip()
            if "аудитория" in second_line:
                parts = second_line.split("аудитория")
                teacher = parts[0].strip()
                classroom = parts[1].strip() if len(parts) > 1 else ""
                if classroom and not re.match(r'^\d{3}$', classroom):
                    classroom = ""
            else:
                teacher = second_line
                
        pairs.append({
            'time': time_slot,
            'subject': subject,
            'teacher': teacher,
            'classroom': classroom
        })

    # Создаем HTML для дня в виде таблицы
    if not pairs:
        # Создаем пустую таблицу для дня без пар
        table_rows = ""
        for i in range(8):  # 8 строк для заполнения высоты
            table_rows += f"""
                <tr>
                    <td style="text-align: center; font-weight: 600; color: #58a6ff;">{i + 1 if i < 4 else ''}</td>
                    <td style="text-align: center; color: #8b949e;">{['08.00-09.30', '09.40-11.10', '11.20-12.50', '13.20-14.50'][i] if i < 4 else ''}</td>
                    <td colspan="5" style="text-align: center; color: #238636; font-size: 1.5em; padding: 40px;">
                        На {current_weekday_name}, {target_date.strftime('%d.%m.%Y')} пар нет!
                    </td>
                </tr>
            """
    else:
        # Создаем таблицу с парами
        table_rows = ""
        
        # Определяем, есть ли первая пара (8:00)
        has_first_pair = any(pair['time'].startswith('08.00') for pair in pairs)
        
        if has_first_pair:
            # Если первая пара есть, нумеруем как обычно
            for i, pair in enumerate(pairs):
                table_rows += f"""
                    <tr>
                        <td style="text-align: center; font-weight: 600; color: #58a6ff;">{i + 1}</td>
                        <td style="text-align: center; color: #8b949e;">{pair['time']}</td>
                        <td colspan="5">
                            <div style="font-weight: 600; color: #f0f6fc; margin-bottom: 8px; font-size: 1.1em;">{pair['subject']}</div>
                            {f'<div style="color: #8b949e; margin-bottom: 5px;">{pair["teacher"]}</div>' if pair['teacher'] else ''}
                            {f'<div style="color: #238636; font-weight: 600;">Ауд. {pair["classroom"]}</div>' if pair['classroom'] else ''}
                        </td>
                    </tr>
                """
        else:
            # Если первой пары нет, первая строка пустая с номером 1
            table_rows += f"""
                <tr>
                    <td style="text-align: center; font-weight: 600; color: #58a6ff;">1</td>
                    <td style="text-align: center; color: #8b949e;">08.00-09.30</td>
                    <td colspan="5"></td>
                </tr>
            """
            
            # Остальные пары нумеруем начиная с 2
            for i, pair in enumerate(pairs):
                table_rows += f"""
                    <tr>
                        <td style="text-align: center; font-weight: 600; color: #58a6ff;">{i + 2}</td>
                        <td style="text-align: center; color: #8b949e;">{pair['time']}</td>
                        <td colspan="5">
                            <div style="font-weight: 600; color: #f0f6fc; margin-bottom: 8px; font-size: 1.1em;">{pair['subject']}</div>
                            {f'<div style="color: #8b949e; margin-bottom: 5px;">{pair["teacher"]}</div>' if pair['teacher'] else ''}
                            {f'<div style="color: #238636; font-weight: 600;">Ауд. {pair["classroom"]}</div>' if pair['classroom'] else ''}
                        </td>
                    </tr>
                """
        
        # Добавляем пустые строки для заполнения высоты
        total_rows = len(pairs) + (0 if has_first_pair else 1)
        for i in range(total_rows, 8):
            table_rows += f"""
                <tr>
                    <td style="text-align: center; font-weight: 600; color: #58a6ff;">{i + 1 if i < 4 else ''}</td>
                    <td style="text-align: center; color: #8b949e;">{['08.00-09.30', '09.40-11.10', '11.20-12.50', '13.20-14.50'][i] if i < 4 else ''}</td>
                    <td colspan="5"></td>
                </tr>
            """
    
    day_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Расписание на {current_weekday_name}</title>
        <style>
            body {{ 
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans', Helvetica, Arial, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #0d1117;
                color: #c9d1d9;
                line-height: 1.6;
            }}
            .container {{
                width: 680px;
                margin: 0 auto;
                background-color: #161b22;
                border: 1px solid #30363d;
                border-radius: 6px;
                box-shadow: 0 8px 24px rgba(1, 4, 9, 0.12);
                overflow: hidden;
                display: flex;
                flex-direction: column;
            }}
            .time-info {{
                background-color: #1c2128;
                padding: 20px;
                text-align: center;
                color: #58a6ff;
                font-size: 1.1em;
                border-bottom: 1px solid #30363d;
            }}
            .table {{
                border-collapse: collapse;
                width: 100%;
                background-color: #161b22;
                font-size: 16px;
                margin: 0;
                flex: 1;
            }}
            .table td, .table th {{
                border: 1px solid #30363d;
                padding: 20px 16px;
                text-align: left;
                vertical-align: top;
            }}
            .table tr:nth-child(even) {{
                background-color: #1c2128;
            }}
            .table tr:nth-child(odd) {{
                background-color: #161b22;
            }}
            .table tr:first-child {{
                background-color: #1f6feb !important;
                color: #f0f6fc;
                font-weight: 600;
                font-size: 14px;
            }}
            .update-time {{
                text-align: center;
                color: #8b949e;
                font-style: italic;
                margin: 0;
                background-color: #21262d;
                border-top: 1px solid #30363d;
                font-size: 14px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="time-info">
                Приходить к: {first_pair_time}
            </div>
            <table class="table">
                <tr>
                    <td style="text-align: center; font-weight: 600;">№</td>
                    <td style="text-align: center; font-weight: 600;">Время</td>
                    <td colspan="5" style="text-align: center; font-weight: 600;">{current_weekday_name}, {target_date.strftime('%d.%m.%Y')}</td>
                </tr>
                {table_rows}
            </table>
            <div class="update-time">
                Обновлено: {datetime.now().strftime('%d.%m.%Y %H:%M')}
            </div>
        </div>
    </body>
    </html>
    """
    
    # Сохраняем HTML файл для дня
    with open(DAY_HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(day_html)

def get_saratov_time() -> datetime:
    """Получаем текущее время в Саратове (UTC+4)"""
    from datetime import timezone, timedelta
    utc_now = datetime.now(timezone.utc)
    saratov_tz = timezone(timedelta(hours=4))
    return utc_now.astimezone(saratov_tz)

def get_smart_date_for_schedule() -> tuple[datetime.date, str]:
    """Умно определяем дату для расписания"""
    saratov_now = get_saratov_time()
    current_time = saratov_now.time()
    
    # Если уже после 18:00, показываем завтра
    if current_time.hour >= 18:
        target_date = saratov_now.date() + timedelta(days=1)
        reason = "завтра (после 18:00)"
    else:
        target_date = saratov_now.date()
        reason = "сегодня"
    
    return target_date, reason

def crop_bottom_200px(image_bytes: bytes, target_height: int) -> bytes:
    """Обрезает 200px снизу с фото"""
    try:
        # Открываем изображение из байтов
        image = Image.open(io.BytesIO(image_bytes))
        
        # Получаем размеры
        width, height = image.size
        
        # Обрезаем снизу 200px
        cropped_image = image.crop((0, 0, width, target_height))
        
        # Конвертируем обратно в байты
        output = io.BytesIO()
        cropped_image.save(output, format='PNG')
        return output.getvalue()
        
    except Exception as e:
        print(f"Ошибка обрезки фото: {e}")
        return image_bytes

def create_today_image() -> bytes:
    """Создаем фото из HTML файла дня с полными стилями"""
    if not os.path.exists(DAY_HTML_FILE):
        return None
    
    try:
        hti = Html2Image()
        hti.output_path = '.'
        
        # Указываем путь к Chrome
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Users\{}\AppData\Local\Google\Chrome\Application\chrome.exe".format(os.getenv('USERNAME', '')),
        ]
        
        for path in chrome_paths:
            if os.path.exists(path):
                hti.browser_executable = path
                break
        
        # Читаем HTML файл
        with open(DAY_HTML_FILE, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Создаем фото с увеличенным размером
        hti.screenshot(html_str=html_content, save_as='today_schedule.png', size=(680, 1040))
        
        # Читаем созданное фото
        with open('today_schedule.png', 'rb') as f:
            image_bytes = f.read()
        
        # Удаляем временный файл
        if os.path.exists('today_schedule.png'):
            os.remove('today_schedule.png')
        
        # Обрезаем 200px снизу (940px - 200px = 740px)
        return crop_bottom_200px(image_bytes, 740)
        
    except Exception as e:
        print(f"Ошибка создания фото дня: {e}")
        return None

def create_week_image() -> bytes:
    """Создаем фото из HTML файла недели с полными стилями"""
    if not os.path.exists(HTML_FILE):
        return None
    
    try:
        hti = Html2Image()
        hti.output_path = '.'
        
        # Указываем путь к Chrome
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Users\{}\AppData\Local\Google\Chrome\Application\chrome.exe".format(os.getenv('USERNAME', '')),
        ]
        
        for path in chrome_paths:
            if os.path.exists(path):
                hti.browser_executable = path
                break
        
        # Читаем HTML файл
        with open(HTML_FILE, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Создаем фото с увеличенным размером
        hti.screenshot(html_str=html_content, save_as='week_schedule.png', size=(1380, 1110))
        
        # Читаем созданное фото
        with open('week_schedule.png', 'rb') as f:
            image_bytes = f.read()
        
        # Удаляем временный файл
        if os.path.exists('week_schedule.png'):
            os.remove('week_schedule.png')
        
        # Обрезаем 50px снизу (970px - 50px = 920px)
        return crop_bottom_200px(image_bytes, 920)
        
    except Exception as e:
        print(f"Ошибка создания фото недели: {e}")
        return None

async def send_html_file(message: types.Message) -> None:
    try:
        if not os.path.exists(HTML_FILE):
            await send_or_edit_message(message.chat.id, "❌ HTML файл не найден. Сначала выполните /update", get_back_keyboard())
            return

        # Читаем HTML файл
        with open(HTML_FILE, 'rb') as f:
            html_content = f.read()
        
        # Находим последнее сообщение бота и редактируем его, добавляя документ
        last_message = await find_last_bot_message(message.chat.id)
        if last_message:
            # Удаляем старое сообщение
            await last_message.delete()
        
        # Отправляем новое сообщение с документом
        await bot.send_document(
            message.chat.id,
            types.BufferedInputFile(html_content, filename="Расписание_ИСП-11.html"),
            caption="📄 HTML версия расписания",
            reply_markup=get_main_keyboard()
        )

    except Exception as e:
        await send_or_edit_message(message.chat.id, f"❌ Произошла ошибка: {e}", get_back_keyboard())

async def get_today_schedule(message: types.Message) -> None:
    try:
        if not os.path.exists(DAY_HTML_FILE):
            await send_or_edit_message(message.chat.id, "❌ HTML файл дня не найден. Сначала выполните /update", get_back_keyboard())
            return

        # Отправляем или редактируем сообщение о создании фото
        status_message = await send_or_edit_message(message.chat.id, "⏳ Создаю фото расписания на сегодня...", get_back_keyboard())
        
        # Создаем фото используя html2image
        image_bytes = create_today_image()
        
        if image_bytes is None:
            await status_message.edit_text("❌ Ошибка создания фото", reply_markup=get_back_keyboard())
            return
        
        # Удаляем старое сообщение
        await status_message.delete()
        
        # Отправляем новое сообщение с фото
        await bot.send_photo(
            message.chat.id,
            types.BufferedInputFile(image_bytes, filename="today_schedule.png"),
            caption="📅 Расписание ИСП-11 на сегодня",
            reply_markup=get_main_keyboard()
        )

    except Exception as e:
        await send_or_edit_message(message.chat.id, f"❌ Произошла ошибка: {e}", get_back_keyboard())

async def get_week_schedule(message: types.Message) -> None:
    try:
        if not os.path.exists(HTML_FILE):
            await send_or_edit_message(message.chat.id, "❌ HTML файл не найден. Сначала выполните /update", get_back_keyboard())
            return

        # Отправляем или редактируем сообщение о создании фото
        status_message = await send_or_edit_message(message.chat.id, "⏳ Создаю фото расписания на неделю...", get_back_keyboard())
        
        # Создаем фото используя html2image
        image_bytes = create_week_image()
        
        if image_bytes is None:
            await status_message.edit_text("❌ Ошибка создания фото", reply_markup=get_back_keyboard())
            return
        
        # Удаляем старое сообщение
        await status_message.delete()
        
        # Отправляем новое сообщение с фото
        await bot.send_photo(
            message.chat.id,
            types.BufferedInputFile(image_bytes, filename="week_schedule.png"),
            caption="📊 Расписание ИСП-11 на неделю",
            reply_markup=get_main_keyboard()
        )

    except Exception as e:
        await send_or_edit_message(message.chat.id, f"❌ Произошла ошибка: {e}", get_back_keyboard())

async def main() -> None:
    await dp.start_polling(bot)

if __name__ == '__main__':
    asyncio.run(main())