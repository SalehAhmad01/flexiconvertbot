import os
import tempfile
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, CallbackQueryHandler, ContextTypes
from PIL import Image
from gtts import gTTS
import pytesseract
from comtypes import client
from fpdf import FPDF
from telegram.constants import ParseMode
from docx import Document  # Import python-docx for DOCX handling

# Bot Token
TOKEN = '7659657213:AAELMh_KQiGlupgZTF4jAmnQsxpuyOAmpqg'  # Replace with your bot token
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Set up Application
app = ApplicationBuilder().token(TOKEN).build()

# Conversion functions
def docx_to_pdf(input_path, output_path):
    try:
        word = client.CreateObject('Word.Application')
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)
        doc.Close()
        word.Quit()
    except Exception as e:
        print(f"Error converting DOCX to PDF: {e}")

def images_to_pdf(image_paths, output_path):
    try:
        # Filter out None and non-image paths
        valid_images = [img for img in image_paths if img and img.lower().endswith(('jpg', 'jpeg', 'png'))]
        
        # Open images and save as PDF
        images = [Image.open(img) for img in valid_images]
        if images:
            images[0].save(output_path, save_all=True, append_images=images[1:])
        else:
            print("No valid images found to convert.")
    except Exception as e:
        print(f"Error converting images to PDF: {e}")

def image_to_text(image_path):
    try:
        image = Image.open(image_path)  # Use the image_path argument
        width, height = 800, 600
        # Convert image to grayscale for better OCR accuracy
        image = image.convert('L')
        # Resize image for better recognition
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

# def pdf_to_text(pdf_path):
#     try:
#         from fitz import open as fitz_open
#         doc = fitz_open(pdf_path)
#         return "".join([page.get_text() for page in doc])
#     except Exception as e:
#         print(f"Error extracting text from PDF: {e}")
#         return ""
def pdf_to_text(pdf_path):
    try:
        from fitz import open as fitz_open
        # Open the PDF document
        with fitz_open(pdf_path) as doc:
            # Extract text from all pages
            text = "".join([page.get_text() for page in doc])
        return text
    except FileNotFoundError:
        print(f"Error: The file at {pdf_path} was not found.")
        return ""
    except Exception as e:
        print(f"Error extracting text from PDF ({pdf_path}): {e}")
        return ""

def pdf_to_docx(pdf_text, output_path):
    try:
        # Create a new Document
        doc = Document()
        doc.add_paragraph(pdf_text)  # Add the extracted PDF text to the document
        doc.save(output_path)  # Save the document as DOCX
    except Exception as e:
        print(f"Error converting PDF to DOCX: {e}")

def text_to_pdf(text, output_path):
    try:
        # Create a PDF instance
        pdf = FPDF()
        pdf.add_page()
        
        # Set font to Arial, with UTF-8 encoding support
        pdf.set_font("Arial", size=12)

        # Write text to PDF, ensuring it handles all characters
        # Split the text into lines and write them
        for line in text.split('\n'):
            pdf.multi_cell(200, 10, txt=line, align='L')
        
        # Output the PDF to a file
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
            text_to_pdf(text, output_path)  # Call the new function
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
            pdf_text = pdf_to_text(file_path)  # Extract text from PDF
            output_path = os.path.join(tempfile.gettempdir(), "output.docx")
            pdf_to_docx(pdf_text, output_path)  # Convert PDF text to DOCX
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
    # Check if the message contains a photo
    if update.message.photo:
        # Get the largest resolution of the photo
        photo = update.message.photo[-1]
        
        # Download the photo file
        file = await context.bot.get_file(photo.file_id)
        file_path = os.path.join(tempfile.gettempdir(), "image.jpg")
        await file.download_to_drive(file_path)

        # Log the file path for debugging
        print(f"Downloaded image to: {file_path}")
        
        # Extract text from image
        text = image_to_text(file_path)

        # Save the extracted text for later use
        context.user_data['text'] = text
        context.user_data['file_path'] = file_path  # Store the file path for PDF conversion

        

        # Clean up the temporary file
        os.remove(file_path)
    else:
        # If the message doesn't contain an image
        await update.message.reply_text("No image found. Please upload a valid image file.")


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if document.mime_type == 'application/pdf':
        file = await context.bot.get_file(document.file_id)
        file_path = os.path.join(tempfile.gettempdir(), document.file_name)
        await file.download_to_drive(file_path)
        context.user_data['file_path'] = file_path
        await ask_conversion_options(update, context, ['Text', 'DOCX'])

# Callback for conversion options
async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()  # Acknowledge the callback query quickly
    await query.message.reply_text("Processing your request, please wait...")  # Intermediate response
    
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
        await query.message.reply_text(f"Error: {e}", parse_mode=ParseMode.MARKDOWN)

# Add handlers to the application
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT, handle_text))
app.add_handler(MessageHandler(filters.PHOTO, handle_image))
app.add_handler(MessageHandler(filters.Document.ALL, handle_pdf))
app.add_handler(CallbackQueryHandler(button_callback))


# Start the bot
if __name__ == "__main__":
    print("Starting bot...")  # Added print statement
    app.run_polling()
