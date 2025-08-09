import os
import pandas as pd
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
    ConversationHandler,
)

# States for conversation
WAITING_FOR_FILLED = 1
WAITING_FOR_EMPTY = 2

# Temporary internal filenames
FILLED_FILE = "filledexcel.xlsx"
EMPTY_FILE = "emptyone.xlsx"
OUTPUT_FILE = "new.xlsx"

PRODUCT_COL = "Product Family Name"
CATEGORY_COL = "Category"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Send me the Excel file with the fully categorized products first (any file name is OK)."
    )
    return WAITING_FOR_FILLED

async def receive_filled_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document:
        await update.message.reply_text("Please send a file. Try again.")
        return WAITING_FOR_FILLED
    
    # Download to temp file
    file = await document.get_file()
    temp_path = "temp_filled_file"
    await file.download_to_drive(temp_path)
    
    # Check if it's a valid Excel file by trying to load it
    try:
        pd.read_excel(temp_path)
    except Exception:
        await update.message.reply_text(
            "This file is not a valid Excel file. Please send a valid Excel file."
        )
        os.remove(temp_path)
        return WAITING_FOR_FILLED
    
    # Rename to internal filename
    if os.path.exists(FILLED_FILE):
        os.remove(FILLED_FILE)
    os.rename(temp_path, FILLED_FILE)
    
    await update.message.reply_text(
        "Received the categorized Excel. Now send me the Excel file with missing categories."
    )
    return WAITING_FOR_EMPTY

async def receive_empty_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document:
        await update.message.reply_text("Please send a file. Try again.")
        return WAITING_FOR_EMPTY
    
    file = await document.get_file()
    temp_path = "temp_empty_file"
    await file.download_to_drive(temp_path)
    
    try:
        pd.read_excel(temp_path)
    except Exception:
        await update.message.reply_text(
            "This file is not a valid Excel file. Please send a valid Excel file."
        )
        os.remove(temp_path)
        return WAITING_FOR_EMPTY
    
    if os.path.exists(EMPTY_FILE):
        os.remove(EMPTY_FILE)
    os.rename(temp_path, EMPTY_FILE)
    
    await update.message.reply_text("Processing... Please wait.")
    
    try:
        fill_categories_from_source()
    except Exception as e:
        await update.message.reply_text(f"Error processing files: {e}")
        return ConversationHandler.END
    
    # Send back the filled Excel file
    await update.message.reply_document(document=open(OUTPUT_FILE, 'rb'))
    await update.message.reply_text("Done! Here's your filled Excel file.")
    
    # Cleanup
    for f in [FILLED_FILE, EMPTY_FILE, OUTPUT_FILE]:
        if os.path.exists(f):
            os.remove(f)
    
    return ConversationHandler.END

def fill_categories_from_source():
    # Load source and target Excel files
    df_source = pd.read_excel(FILLED_FILE)
    df_target = pd.read_excel(EMPTY_FILE)

    # Create mapping from product name to category
    product_to_category = dict(zip(df_source[PRODUCT_COL], df_source[CATEGORY_COL]))

    # Fill category in target file
    df_target[CATEGORY_COL] = df_target[PRODUCT_COL].map(product_to_category)

    # Save output
    df_target.to_excel(OUTPUT_FILE, index=False)

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Cancelled.")
    return ConversationHandler.END

if __name__ == "__main__":
    TOKEN = "8404087668:AAFRabaGH_Gfc-ljJVlT3ZxScwlkHDNRG6c"  # Replace with your bot token

    app = ApplicationBuilder().token(TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAITING_FOR_FILLED: [MessageHandler(filters.Document.ALL, receive_filled_file)],
            WAITING_FOR_EMPTY: [MessageHandler(filters.Document.ALL, receive_empty_file)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv_handler)

    print("Bot is running...")
    app.run_polling()
