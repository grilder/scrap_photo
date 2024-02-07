from telegram import Update, ForceReply
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackContext
import os
from openpyxl import Workbook

# Create a new Excel workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "Hashtags"
ws.append(["Image Name", "Hashtags"])
wb.save("hashtags.xlsx")

# Define the bot's token
TOKEN = '6914361513:AAE1PVuiEAZcSlGQB6xQC2IPdzR0mCEqK00'  # Замените на ваш новый токен

# Define the base path for saving images
BASE_IMAGE_PATH = 'images/'

# Check if directory exists and create if not
if not os.path.exists(BASE_IMAGE_PATH):
    os.makedirs(BASE_IMAGE_PATH)

# Define the start command handler
def start(update: Update, context: CallbackContext) -> None:
    user = update.effective_user
    update.message.reply_markdown_v2(
        fr'Здравствуйте, {user.mention_markdown_v2()}\! Пожалуйста, отправьте мне фотографию с хештегами в подписи\.',
        reply_markup=ForceReply(selective=True),
    )

# Define the help command handler
def help_command(update: Update, context: CallbackContext) -> None:
    update.message.reply_text('Отправьте мне изображение с хештегами в подписи, и я сохраню их.')

# Define the image and hashtag handler
def image_handler(update: Update, context: CallbackContext) -> None:
    # Extract the image and its hashtags
    photo = update.message.photo[-1]
    message_caption = update.message.caption or ""
    hashtags = [ht.strip() for ht in message_caption.split() if ht.startswith('#')]
    
    # Download the image
    photo_file = photo.get_file()
    file_path = f'{BASE_IMAGE_PATH}{photo.file_id}.jpg'
    photo_file.download(file_path)
    
    # Save the hashtags in the Excel file
    ws.append([photo.file_id, " ".join(hashtags)])
    wb.save("hashtags.xlsx")

    update.message.reply_text(f'Изображение и хештеги сохранены.\nПуть к файлу: {file_path}')

# Main function to start the bot
def main() -> None:
    # Create the Updater and pass it your bot's token.
    updater = Updater(TOKEN)

    # Get the dispatcher to register handlers
    dispatcher = updater.dispatcher

    # on different commands - answer in Telegram
    dispatcher.add_handler(CommandHandler("start", start))
    dispatcher.add_handler(CommandHandler("help", help_command))

    # on non-command i.e message - save image and hashtags
    dispatcher.add_handler(MessageHandler(Filters.photo & Filters.caption, image_handler))

    # Start the Bot
    updater.start_polling()

    updater.idle()

if __name__ == '__main__':
    main()