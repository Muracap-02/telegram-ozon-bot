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

# Проверим, есть ли шаблон в папке
if not os.path.exists(TEMPLATE_PATH):
    raise FileNotFoundError(f"Шаблон {TEMPLATE_FILENAME} не найден рядом со скриптом!")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отправьте основной Excel-файл. Шаблон будет использован автоматически из папки со скриптом.")

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    document = update.message.document
    logger.info(f"[LOG] Получен файл от пользователя @{user}: {document.file_name}")

    file = await context.bot.get_file(document.file_id)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        await file.download_to_drive(f.name)
        data_file = f.name

    logger.info("[LOG] Загрузка шаблона...")
    df = pd.read_excel(data_file, header=None, skiprows=3)
    logger.info("[LOG] Чтение данных из 4-й строки")

    # Задача 1: Замена значений в колонках E и F
    def replace_passport_and_birthdate(value):
        allowed_prefixes = ("AD", "AB", "FA", "XS", "AE", "AC", "AA")
        if isinstance(value, str) and not value.startswith(allowed_prefixes):
            return "AB0663236"
        return value

    df[4] = df[4].apply(replace_passport_and_birthdate)
    df[5] = ["23,12,1988" if val == "AB0663236" else old for val, old in zip(df[4], df[5])]

    # Задача 2: Добавление 0 к 5-значным кодам в колонке K
    df[10] = df[10].apply(lambda x: f"0{x}" if pd.notna(x) and isinstance(x, (int, float, str)) and len(str(int(float(x)))) == 5 else x)

    # Задача 3: Удаление повторяющихся посылок по колонке A
    seen = set()
    for idx, val in df[0].items():
        val = str(val).strip()
        if val and val in seen:
            df.loc[idx, 0:7] = None
        else:
            seen.add(val)

    # Разделение на чанки по 1000 строк
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

    # Создание ZIP архива
    zip_path = os.path.join(tempfile.gettempdir(), "AllPackages.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file_path in output_files:
            arcname = os.path.basename(file_path)
            zipf.write(file_path, arcname)
            logger.info(f"[LOG] Добавлен в архив: {arcname}")

    await update.message.reply_text(f"Обработка завершена. Отправляю архив с {len(output_files)} файлами.")
    await context.bot.send_document(chat_id=update.message.chat_id, document=open(zip_path, 'rb'))

def main():
    app = ApplicationBuilder().token("7872241701:AAF633V3rjyXTJkD8F0lEW13nDtAqHoqeic").build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))

    logger.info("[LOG] Бот запущен и готов к работе...")
    app.run_polling()

if __name__ == '__main__':
    main()
