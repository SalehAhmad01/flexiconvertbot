import os
import tempfile
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, CallbackQueryHandler, ContextTypes
from PIL import Image
from gtts import gTTS
import pytesseract
from fpdf import FPDF
from telegram.constants import ParseMode
from docx import Document

# Bot Token
TOKEN = '7659657213:AAELMh_KQiGlupgZTF4jAmnQsxpuyOAmpqg'  # Replace with your bot token
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Set up Application
app = ApplicationBuilder().token(TOKEN).build()

# Conversion functions
def docx_to_pdf(input_path, output_path):
    try:
        # Use python-docx to read the document
        doc = Document(input_path)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Add paragraphs from the DOCX file to the PDF
        for paragraph in doc.paragraphs:
            pdf.multi_cell(0, 10, paragraph.text)

        pdf.output(output_path)
    except Exception as e:
        print(f"Error converting DOCX to PDF: {e}")

def images_to_pdf(image_paths, output_path):
    try:
        valid_images = [img for img in image_paths if img and img.lower().endswith(('jpg', 'jpeg', 'png'))]
        images = [Image.open(img) for img in valid_images]
        if images:
            images[0].save(output_path, save_all=True, append_images=images[1:])
        else:
            print("No valid images found to convert.")
    except Exception as e:
        print(f"Error converting images to PDF: {e}")

def image_to_text(image_path):
    try:
        image = Image.open(image_path)
        width, height = 800, 600
        image = image.convert('L')
        image = image.resize((width, height), Image.Resampling.LANCZOS)
        return pytesseract.image_to_string(image)
    except Exception as e:
        print(f"Error extracting text from image: {e}")
        return ""

def text_to_speech(text, output_path):
    try:
        tts = gTTS(text=text, lang='en')
        tts.save(output_path)
    except Exception as e:
        print(f"Error converting text to speech: {e}")

def pdf_to_text(pdf_path):
    try:
        from fitz import open as fitz_open
        with fitz_open(pdf_path) as doc:
            text = "".join([page.get_text() for page in doc])
        return text
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")
        return ""

def pdf_to_docx(pdf_text, output_path):
    try:
        doc = Document()
        doc.add_paragraph(pdf_text)
        doc.save(output_path)
    except Exception as e:
        print(f"Error converting PDF to DOCX: {e}")

def text_to_pdf(text, output_path):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for line in text.split('\n'):
            pdf.multi_cell(0, 10, txt=line, align='L')
        pdf.output(output_path)
    except Exception as e:
        print(f"Error converting text to PDF: {e}")

# Callback for conversion options
async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    file_path = context.user_data.get('file_path')
    text = context.user_data.get('text', "")

    try:
        if query.data == 'Speech' and text:
            output_path = os.path.join(tempfile.gettempdir(), "speech.mp3")
            text_to_speech(text, output_path)
            await query.message.reply_voice(open(output_path, 'rb'))
            cleanup(output_path)

        elif query.data == 'PDF' and text:
            output_path = os.path.join(tempfile.gettempdir(), "output.pdf")
            text_to_pdf(text, output_path)
            await query.message.reply_document(open(output_path, 'rb'))
            cleanup(output_path)

        elif query.data == 'DOCX' and text:
            output_path = os.path.join(tempfile.gettempdir(), "output.docx")
            doc = Document()
            doc.add_paragraph(text)
            doc.save(output_path)
            await query.message.reply_document(open(output_path, 'rb'))
            cleanup(output_path)

        elif query.data == 'DOCX' and file_path:
            pdf_text = pdf_to_text(file_path)
            output_path = os.path.join(tempfile.gettempdir(), "output.docx")
            pdf_to_docx(pdf_text, output_path)
            await query.message.reply_document(open(output_path, 'rb'))
            cleanup(output_path)

    except Exception as e:
        await query.message.reply_text(f"Error: {e}")

def cleanup(file_path):
    try:
        os.remove(file_path)
    except Exception as e:
        print(f"Error cleaning up file {file_path}: {e}")

# Start command
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    greeting_text = (
        "Hello! I am a bot that can:\n"
        "- Convert docs to PDF\n"
        "- Convert images to PDF\n"
        "- Convert images to text\n"
        "- Convert text to speech\n"
        "- Convert PDF to text\n"
        "- Convert PDF to DOCX\n"
        "Upload a file to start your conversion."
    )
    await update.message.reply_text(greeting_text)

# Ask user to select conversion options
async def ask_conversion_options(update: Update, context: ContextTypes.DEFAULT_TYPE, options):
    keyboard = [[InlineKeyboardButton(option, callback_data=option) for option in options]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text('Please choose a conversion option:', reply_markup=reply_markup)

# Handlers for documents, images, and text
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    context.user_data['text'] = text
    await ask_conversion_options(update, context, ['Speech', 'PDF', 'DOCX'])

async def handle_image(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.photo:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        file_path = os.path.join(tempfile.gettempdir(), "image.jpg")
        await file.download_to_drive(file_path)

        text = image_to_text(file_path)
        context.user_data['text'] = text
        context.user_data['file_path'] = file_path
        os.remove(file_path)
    else:
        await update.message.reply_text("No image found. Please upload a valid image file.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if document.mime_type == 'application/pdf':
        file = await context.bot.get_file(document.file_id)
        file_path = os.path.join(tempfile.gettempdir(), document.file_name)
        await file.download_to_drive(file_path)
        context.user_data['file_path'] = file_path
        await ask_conversion_options(update, context, ['Text', 'DOCX'])

# Add handlers to the application
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT, handle_text))
app.add_handler(MessageHandler(filters.PHOTO, handle_image))
app.add_handler(MessageHandler(filters.Document.ALL, handle_pdf))
app.add_handler(CallbackQueryHandler(button_callback))

# Start the bot
if __name__ == "__main__":
    print("Starting bot...")
    app.run_polling()
