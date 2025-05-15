import os
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes, CommandHandler
import tempfile
import logging
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import zipfile

# Настройка логирования
logging.basicConfig(format='[LOG] %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TEMPLATE_FILENAME = "AllPackageEC_.xlsx"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)

# НЕ проверяем наличие шаблона через raise
if not os.path.exists(TEMPLATE_PATH):
    logger.warning(f"[WARN] Шаблон {TEMPLATE_FILENAME} не найден!")

allowed_prefixes = ("AB", "AC", "AA", "AD", "FA", "XS")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отправьте основной Excel-файл. Шаблон и база уже находятся рядом со скриптом.")

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    document = update.message.document
    logger.info(f"[LOG] Получен файл от пользователя @{user}: {document.file_name}")

    file = await context.bot.get_file(document.file_id)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        await file.download_to_drive(f.name)
        data_file = f.name

    logger.info("[LOG] Чтение входного файла...")
    df = pd.read_excel(data_file, header=None, skiprows=3)

    # --- Задача 1: замена паспортов и дат рождения ---
    def should_replace(val):
        if pd.isna(val):
            return True
        val = str(val).strip().upper()
        return not any(val.startswith(prefix) for prefix in allowed_prefixes)

    for idx, val in df[4].items():
        if should_replace(val):
            df.at[idx, 4] = "AB0663236"
            df.at[idx, 5] = "23,12,1988"

    # --- Задача 2: добавление 0 к 5-значным кодам в колонке K (index 10) ---
    def fix_code(x):
        try:
            s = str(int(float(x)))
            if len(s) == 5:
                return "0" + s
            return x
        except:
            return x

    df[10] = df[10].apply(fix_code)

    # --- Задача 3: удаление дубликатов по колонке A ---
    seen = set()
    for idx, val in df[0].items():
        val = str(val).strip()
        if val and val in seen:
            df.loc[idx, 0:7] = None
        else:
            seen.add(val)

    # --- Разделение по чанкам и запись в шаблон ---
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

    # --- Архивация всех файлов в один ZIP ---
    zip_path = os.path.join(tempfile.gettempdir(), f"AllPackageEC_{user}.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file_path in output_files:
            zipf.write(file_path, os.path.basename(file_path))

    await update.message.reply_text(f"Обработка завершена. Файлы упакованы в архив.")
    await context.bot.send_document(chat_id=update.message.chat_id, document=open(zip_path, 'rb'))

def main():
    app = ApplicationBuilder().token("7872241701:AAF633V3rjyXTJkD8F0lEW13nDtAqHoqeic").build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))

    logger.info("[LOG] Бот запущен и готов к работе...")
    app.run_polling()

if __name__ == '__main__':
    main()
