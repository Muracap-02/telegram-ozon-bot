import os
import pandas as pd
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, MessageHandler, filters,
    ContextTypes, CommandHandler, ConversationHandler
)
import tempfile
import logging
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import zipfile

# Логирование
logging.basicConfig(format='[LOG] %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TEMPLATE_FILENAME = "AllPackageEC_.xlsx"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)

# Стейты
CHOOSE_MODE, WAIT_FILE = range(2)
user_mode = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [["Обработать на части", "Макрос Пасспорт"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("Выберите режим:", reply_markup=reply_markup)
    return CHOOSE_MODE

async def choose_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    choice = update.message.text
    user_id = update.message.from_user.id

    if choice not in ["Обработать на части", "Макрос Пасспорт"]:
        await update.message.reply_text("Пожалуйста, выберите один из вариантов.")
        return CHOOSE_MODE

    user_mode[user_id] = choice
    await update.message.reply_text("Отправьте Excel-файл для обработки.")
    return WAIT_FILE

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    user_id = user.id
    document = update.message.document
    logger.info(f"[LOG] Файл от @{user.username}: {document.file_name}")

    file = await context.bot.get_file(document.file_id)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        await file.download_to_drive(f.name)
        data_file = f.name

    mode = user_mode.get(user_id)

    if mode == "Обработать на части":
        await process_in_parts(update, context, data_file)
    elif mode == "Макрос Пасспорт":
        await process_passport_macro(update, context, data_file)
    else:
        await update.message.reply_text("Ошибка: режим обработки не выбран.")
        return ConversationHandler.END

    return ConversationHandler.END

async def process_in_parts(update, context, data_file):
    logger.info("[LOG] Обработка: разбивка на части")
    df = pd.read_excel(data_file, header=None, skiprows=3)

    def fix_code(x):
        try:
            s = str(int(float(x)))
            if len(s) == 5:
                return "0" + s
            return x
        except:
            return x

    df[10] = df[10].apply(fix_code)

    seen = set()
    for idx, val in df[0].items():
        val = str(val).strip()
        if val and val in seen:
            df.loc[idx, 0:7] = None
        else:
            seen.add(val)

    chunk_size = 1000
    parts = [df[i:i + chunk_size] for i in range(0, len(df), chunk_size)]

    output_files = []
    for idx, part in enumerate(parts):
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        for r_idx, row in enumerate(dataframe_to_rows(part, index=False, header=False), start=4):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        start_range = idx * chunk_size
        filename = f"AllPackageEC_{start_range}.xlsx"
        output_path = os.path.join(tempfile.gettempdir(), filename)
        wb.save(output_path)
        output_files.append(output_path)
        logger.info(f"[LOG] Сохранён файл: {filename}")

    zip_path = os.path.join(tempfile.gettempdir(), f"AllPackageEC_{user.username}.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file_path in output_files:
            zipf.write(file_path, os.path.basename(file_path))

    await update.message.reply_text("Обработка завершена. Архив отправляется...")
    await context.bot.send_document(chat_id=update.message.chat_id, document=open(zip_path, 'rb'))

async def process_passport_macro(update, context, data_file):
    logger.info("[LOG] Выполняется макрос 'Пасспорт'")
    wb = load_workbook(data_file)
    ws = wb.active

    valid_start = "123456789MRTGKZECUVFBNDGHJLKQIP"
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        val = str(row[4].value).strip() if row[4].value else ""
        if val and val[0].upper() in valid_start:
            row[4].value = "AB0663236"
            row[5].value = "23,12,1988"

    output_path = os.path.join(tempfile.gettempdir(), f"PassportUpdated_{update.message.from_user.username}.xlsx")
    wb.save(output_path)

    await update.message.reply_text("Макрос выполнен. Файл отправляется...")
    await context.bot.send_document(chat_id=update.message.chat_id, document=open(output_path, 'rb'))

def main():
    app = ApplicationBuilder().token("YOUR_TOKEN_HERE").build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            CHOOSE_MODE: [MessageHandler(filters.TEXT, choose_mode)],
            WAIT_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
        },
        fallbacks=[],
    )

    app.add_handler(conv_handler)
    logger.info("[LOG] Бот запущен...")
    app.run_polling()

if __name__ == '__main__':
    main()
