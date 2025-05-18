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

logging.basicConfig(format='[LOG] %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TEMPLATE_FILENAME = "AllPackageEC_.xlsx"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)

MODE_CHOICE = {}
LOG_FILE_PATH = "bot.log"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("▶️ Обработать на части", callback_data="chunk")],
        [InlineKeyboardButton("📄 Макрос Пасспорт", callback_data="passport")],
        [InlineKeyboardButton("📞 Поддержка: +998334743434", url="tel:+998334743434")],
        [InlineKeyboardButton("🗑 Удалить лог", callback_data="clear_log")]
    ])
    await update.message.reply_text(
        "👋 Добро пожаловать!\nВыберите режим обработки:",
        reply_markup=keyboard
    )

async def mode_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    MODE_CHOICE[user_id] = query.data

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("📞 Поддержка: +998334743434", url="tel:+998334743434")],
        [InlineKeyboardButton("🗑 Удалить лог", callback_data="clear_log")]
    ])

    await query.message.reply_text(
        "📎 Пожалуйста, отправьте Excel-файл для выбранной обработки.",
        reply_markup=keyboard
    )

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    user_id = user.id
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

    zip_path = os.path.join(tempfile.gettempdir(), f"AllPackageEC_{update.message.from_user.username}.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file_path in output_files:
            zipf.write(file_path, os.path.basename(file_path))

    await update.message.reply_text("✅ Обработка завершена. Архив отправляется...")
    await context.bot.send_document(chat_id=update.message.chat_id, document=open(zip_path, 'rb'))
    await show_main_buttons(update, context)

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

    await update.message.reply_text("✅ Макрос выполнен. Файл отправляется...")
    await context.bot.send_document(chat_id=update.message.chat_id, document=open(output_path, 'rb'))
    await show_main_buttons(update, context)

async def show_main_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("▶️ Обработать на части", callback_data="chunk")],
        [InlineKeyboardButton("📄 Макрос Пасспорт", callback_data="passport")],
        [InlineKeyboardButton("📞 Поддержка: +998334743434", url="tel:+998334743434")],
        [InlineKeyboardButton("🗑 Удалить лог", callback_data="clear_log")]
    ])
    await context.bot.send_message(chat_id=update.message.chat_id, text="Что вы хотите сделать дальше?", reply_markup=keyboard)

async def clear_log(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if os.path.exists(LOG_FILE_PATH):
        os.remove(LOG_FILE_PATH)
        await query.message.reply_text("🗑 Лог успешно удалён.")
    else:
        await query.message.reply_text("📁 Лог уже пуст.")
    await show_main_buttons(query, context)

def main():
    app = ApplicationBuilder().token("7872241701:AAF633V3rjyXTJkD8F0lEW13nDtAqHoqeic").build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(mode_selected, pattern="^(chunk|passport)$"))
    app.add_handler(CallbackQueryHandler(clear_log, pattern="^clear_log$"))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))

    logger.info("[LOG] Бот запущен...")
    app.run_polling()

if __name__ == '__main__':
    main()
