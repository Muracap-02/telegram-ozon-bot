import os
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, MessageHandler, filters,
    ContextTypes, CommandHandler, CallbackQueryHandler
)
import tempfile
import logging
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import zipfile

# Путь к лог-файлу
LOG_FILE = "bot.log"
logging.basicConfig(format='[LOG] %(message)s', level=logging.INFO, filename=LOG_FILE)
logger = logging.getLogger(__name__)

# Шаблон Excel
TEMPLATE_FILENAME = "AllPackageEC_.xlsx"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)

# Логотип
LOGO_URL = "https://sdmntprnortheu.oaiusercontent.com/files/00000000-e354-61f4-96fa-e3575a0560e9/raw?se=2025-05-18T17%3A42%3A35Z&sp=r&sv=2024-08-04&sr=b&scid=00000000-0000-0000-0000-000000000000&skoid=b32d65cd-c8f1-46fb-90df-c208671889d4&sktid=a48cca56-e6da-484e-a814-9c849652bcb3&skt=2025-05-18T09%3A21%3A08Z&ske=2025-05-19T09%3A21%3A08Z&sks=b&skv=2024-08-04&sig=jryFrwnA9%2BlNVxH%2B7pMu1GRs2SeldZaRWxZgXWiiVx4%3D"

MODE_CHOICE = {}

# Основные кнопки
def get_main_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("▶️ Обработать на части", callback_data="chunk")],
        [InlineKeyboardButton("📄 Макрос Пасспорт", callback_data="passport")],
        [InlineKeyboardButton("📞 Поддержка: +998334743434", url="tel:+998334743434")],
        [InlineKeyboardButton("🗑 Удалить лог", callback_data="clear_log")]
    ])

# Старт
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_photo(
        photo=LOGO_URL,
        caption="👋 Добро пожаловать!\nВыберите, что вы хотите сделать:",
        reply_markup=get_main_keyboard()
    )

# Выбор режима
async def mode_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    MODE_CHOICE[user_id] = query.data

    await context.bot.send_message(
        chat_id=query.message.chat.id,
        text="📎 Пожалуйста, отправьте Excel-файл для выбранной обработки."
    )

# Обработка файла
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    mode = MODE_CHOICE.get(user_id)

    if not mode:
        await update.message.reply_text("Сначала выберите режим через /start.")
        return

    document = update.message.document
    file = await context.bot.get_file(document.file_id)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        await file.download_to_drive(f.name)
        data_file = f.name

    if mode == "chunk":
        await process_in_parts(update, context, data_file)
    elif mode == "passport":
        await process_passport_macro(update, context, data_file)

# Обработка на части
async def process_in_parts(update, context, data_file):
    logger.info("[LOG] Обработка: разбивка на части")
    df = pd.read_excel(data_file, header=None, skiprows=3)

    def fix_code(x):
        try:
            s = str(int(float(x)))
            return "0" + s if len(s) == 5 else x
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

        filename = f"AllPackageEC_{idx * chunk_size}.xlsx"
        output_path = os.path.join(tempfile.gettempdir(), filename)
        wb.save(output_path)
        output_files.append(output_path)
        logger.info(f"[LOG] Сохранён файл: {filename}")

    zip_path = os.path.join(tempfile.gettempdir(), f"AllPackageEC_{update.message.from_user.username}.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file_path in output_files:
            zipf.write(file_path, os.path.basename(file_path))

    await update.message.reply_text("✅ Обработка завершена. Отправляю архив...")
    await context.bot.send_document(chat_id=update.message.chat_id, document=open(zip_path, 'rb'))
    await send_final_buttons(update, context)

# Макрос Пасспорт
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

    await update.message.reply_text("✅ Макрос выполнен. Отправляю файл...")
    await context.bot.send_document(chat_id=update.message.chat_id, document=open(output_path, 'rb'))
    await send_final_buttons(update, context)

# Кнопки после завершения
async def send_final_buttons(update, context):
    chat_id = update.effective_chat.id
    await context.bot.send_message(
        chat_id=chat_id,
        text="Что вы хотите сделать дальше? 👇",
        reply_markup=get_main_keyboard()
    )

# Очистка логов
async def handle_clear_log(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    try:
        if os.path.exists(LOG_FILE):
            os.remove(LOG_FILE)
            await query.message.reply_text("🗑 Лог-файл успешно удалён.")
        else:
            await query.message.reply_text("ℹ️ Лог-файл уже был удалён.")
    except Exception as e:
        await query.message.reply_text(f"❌ Ошибка при удалении лога: {e}")

# Запуск бота
def main():
    app = ApplicationBuilder().token("7872241701:AAF633V3rjyXTJkD8F0lEW13nDtAqHoqeic").build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(mode_selected, pattern="^(chunk|passport)$"))
    app.add_handler(CallbackQueryHandler(handle_clear_log, pattern="^clear_log$"))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))

    logger.info("[LOG] Бот запущен...")
    app.run_polling()

if __name__ == '__main__':
    main()
