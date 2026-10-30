import os
import comtypes.client
from PIL import Image
from PyPDF2 import PdfReader, PdfMerger
from pdf2docx import Converter
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    filters,
    ContextTypes,
)

TOKEN = "9269963148:AAFFdHNvFUDwteS9KIF775Te5x_-8E7QGwQ"

# --- دوال التحويل والمعالجة ---

def convert_ppt_to_pdf(input_path, output_path):
    comtypes.CoInitialize()
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(os.path.abspath(input_path))
    deck.SaveAs(os.path.abspath(output_path), 32)
    deck.Close()
    powerpoint.Quit()

def convert_word_to_pdf(input_path, output_path):
    comtypes.CoInitialize()
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 0
    doc = word.Documents.Open(os.path.abspath(input_path))
    doc.SaveAs(os.path.abspath(output_path), 17)
    doc.Close()
    word.Quit()

def convert_pdf_to_word(input_path, output_path):
    cv = Converter(input_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()

def convert_img_to_pdf(input_path, output_path):
    image = Image.open(input_path)
    if image.mode != "RGB":
        image = image.convert("RGB")
    image.save(output_path, "PDF", resolution=100.0)

def merge_pdf_files(file_paths, output_path):
    merger = PdfMerger()
    for path in file_paths:
        merger.append(path)
    merger.write(output_path)
    merger.close()

# --- القائمة الرئيسية والأزرار ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["merge_files"] = []
    
    keyboard = [
        [InlineKeyboardButton("📚 دمج ملفات PDF", callback_data="merge_pdfs")],
        [InlineKeyboardButton("📊 PowerPoint ➔ PDF", callback_data="ppt_to_pdf")],
        [InlineKeyboardButton("📝 Word ➔ PDF", callback_data="word_to_pdf")],
        [InlineKeyboardButton("📄 PDF ➔ Word", callback_data="pdf_to_word")],
        [InlineKeyboardButton("🖼️ صورة ➔ PDF", callback_data="img_to_pdf")],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    text = (
        "👋 أهلاً بك في بوت **PDFMasterDZ**!\n\n"
        "📑 **اختر نوع العملية المطلوبة من الأزرار أدناه:**\n"
        "(أو أرسل أي ملف PDF مباشرة لفحصه)"
    )
    if update.message:
        await update.message.reply_text(text, reply_markup=reply_markup, parse_mode="Markdown")
    elif update.callback_query:
        await update.callback_query.message.reply_text(text, reply_markup=reply_markup, parse_mode="Markdown")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    choice = query.data

    if choice == "merge_pdfs":
        context.user_data["action"] = "merge_pdfs"
        context.user_data["merge_files"] = []
        await query.edit_message_text(
            "📚 **وضع دمج ملفات الـ PDF:**\n\n"
            "أرسل لي الآن ملفات الـ PDF واحداً تلو الآخر بالترتيب المطلوب.\n"
            "عند الانتهاء من إرسال الملفات، اضغط على زر **'إتمام الدمج الآن'** في الأسفل."
        )
    elif choice == "finish_merge":
        files_to_merge = context.user_data.get("merge_files", [])
        if len(files_to_merge) < 2:
            await query.message.reply_text("⚠️ يجب إرسال ملفين PDF على الأقل لإتمام الدمج!")
            return

        status_msg = await query.message.reply_text("⏳ جاري دمج الملفات معاً في ملف واحد...")
        output_file = os.path.abspath("Merged_Document.pdf")
        try:
            merge_pdf_files(files_to_merge, output_file)
            await status_msg.edit_text("✅ تم دمج الملفات بنجاح! جاري إرسال الملف...")
            with open(output_file, "rb") as f:
                await query.message.reply_document(document=f, filename="Merged_Document.pdf")
        except Exception:
            await status_msg.edit_text("❌ حدث خطأ أثناء دمج الملفات.")
        finally:
            for p in files_to_merge:
                if os.path.exists(p): os.remove(p)
            if os.path.exists(output_file): os.remove(output_file)
            context.user_data["merge_files"] = []
            context.user_data["action"] = None

    elif choice == "ppt_to_pdf":
        context.user_data["action"] = "ppt_to_pdf"
        await query.edit_message_text("📥 أرسل لي الآن ملف الـ **PowerPoint (.pptx أو .ppt)**:")
    elif choice == "word_to_pdf":
        context.user_data["action"] = "word_to_pdf"
        await query.edit_message_text("📥 أرسل لي الآن ملف الـ **Word (.docx أو .doc)**:")
    elif choice == "pdf_to_word":
        context.user_data["action"] = "pdf_to_word"
        await query.edit_message_text("📥 أرسل لي الآن ملف الـ **PDF** لتحويله إلى مستند Word:")
    elif choice == "img_to_pdf":
        context.user_data["action"] = "img_to_pdf"
        await query.edit_message_text("📥 أرسل لي الآن **الصورة**:")

# --- معالجة الملفات ---

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    action = context.user_data.get("action")
    doc = update.message.document
    filename = doc.file_name.lower()
    base_name = os.path.splitext(doc.file_name)[0]

    # حالة دمج ملفات PDF
    if action == "merge_pdfs":
        if not filename.endswith(".pdf"):
            await update.message.reply_text("⚠️ يرجى إرسال ملفات بصيغة PDF فقط في وضع الدمج.")
            return

        current_files = context.user_data.setdefault("merge_files", [])
        temp_path = os.path.abspath(f"merge_part_{len(current_files)}_{doc.file_name}")
        
        tg_file = await context.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(temp_path)
        current_files.append(temp_path)

        keyboard = [[InlineKeyboardButton("✅ إتمام الدمج الآن", callback_data="finish_merge")]]
        reply_markup = InlineKeyboardMarkup(keyboard)

        await update.message.reply_text(
            f"📥 استلمت الملف رقم ({len(current_files)}): **{doc.file_name}**\n\n"
            f"أرسل الملف التالي، أو اضغط على الزر أدناه لإتمام الدمج:",
            reply_markup=reply_markup,
            parse_mode="Markdown"
        )
        return

    # باقي التحويلات
    input_file = os.path.abspath(f"temp_{doc.file_name}")
    status_msg = await update.message.reply_text("⏳ جاري استلام ومعالجة الملف...")

    try:
        tg_file = await context.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(input_file)

        if action == "ppt_to_pdf":
            if not (filename.endswith(".pptx") or filename.endswith(".ppt")):
                await status_msg.edit_text("⚠️ يرجى إرسال ملف PowerPoint بصيغة (.pptx أو .ppt).")
                return
            output_file = os.path.abspath(f"{base_name}.pdf")
            convert_ppt_to_pdf(input_file, output_file)
            await status_msg.edit_text("✅ تم التحويل بنجاح! جاري إرسال الملف...")
            with open(output_file, "rb") as f:
                await update.message.reply_document(document=f, filename=f"{base_name}.pdf")
            if os.path.exists(output_file): os.remove(output_file)

        elif action == "word_to_pdf":
            if not (filename.endswith(".docx") or filename.endswith(".doc")):
                await status_msg.edit_text("⚠️ يرجى إرسال ملف Word بصيغة (.docx أو .doc).")
                return
            output_file = os.path.abspath(f"{base_name}.pdf")
            convert_word_to_pdf(input_file, output_file)
            await status_msg.edit_text("✅ تم التحويل بنجاح! جاري إرسال الملف...")
            with open(output_file, "rb") as f:
                await update.message.reply_document(document=f, filename=f"{base_name}.pdf")
            if os.path.exists(output_file): os.remove(output_file)

        elif action == "pdf_to_word":
            if not filename.endswith(".pdf"):
                await status_msg.edit_text("⚠️ يرجى إرسال ملف بصيغة PDF فقط.")
                return
            output_file = os.path.abspath(f"{base_name}.docx")
            convert_pdf_to_word(input_file, output_file)
            await status_msg.edit_text("✅ تم التحويل بنجاح! جاري إرسال مستند Word...")
            with open(output_file, "rb") as f:
                await update.message.reply_document(document=f, filename=f"{base_name}.docx")
            if os.path.exists(output_file): os.remove(output_file)

        elif action == "img_to_pdf":
            output_file = os.path.abspath(f"{base_name}.pdf")
            convert_img_to_pdf(input_file, output_file)
            await status_msg.edit_text("✅ تم تحويل الصورة إلى PDF بنجاح!")
            with open(output_file, "rb") as f:
                await update.message.reply_document(document=f, filename=f"{base_name}.pdf")
            if os.path.exists(output_file): os.remove(output_file)

        elif filename.endswith(".pdf"):
            reader = PdfReader(input_file)
            pages_count = len(reader.pages)
            snippet = reader.pages[0].extract_text() or "لا يحتوي نصاً قابلاً للقراءة."
            report = (
                f"✅ **تم فحص ملف الـ PDF**\n\n"
                f"📁 **الاسم:** {doc.file_name}\n"
                f"📄 **الصفحات:** {pages_count}\n\n"
                f"📝 **المقتطف:**\n{snippet[:200]}..."
            )
            await status_msg.edit_text(report)

        else:
            await status_msg.edit_text("يرجى اختيار نوع التحويل أولاً عبر إرسال /start.")

        context.user_data["action"] = None

    except Exception:
        await status_msg.edit_text("❌ حدث خطأ أثناء المعالجة، يرجى التأكد من سلامة الملف.")
    finally:
        if os.path.exists(input_file):
            os.remove(input_file)

# معالجة الصور العادية المرسلة كصورة
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    action = context.user_data.get("action")
    if action == "img_to_pdf":
        status_msg = await update.message.reply_text("⏳ جاري تحويل الصورة إلى PDF...")
        photo = update.message.photo[-1]
        input_file = os.path.abspath("temp_photo.jpg")
        output_file = os.path.abspath("converted_photo.pdf")

        try:
            tg_file = await context.bot.get_file(photo.file_id)
            await tg_file.download_to_drive(input_file)

            convert_img_to_pdf(input_file, output_file)

            await status_msg.edit_text("✅ تم تحويل الصورة إلى PDF بنجاح!")
            with open(output_file, "rb") as f:
                await update.message.reply_document(document=f, filename="Image_Document.pdf")

            context.user_data["action"] = None
        except Exception:
            await status_msg.edit_text("❌ حدث خطأ أثناء تحويل الصورة.")
        finally:
            if os.path.exists(input_file): os.remove(input_file)
            if os.path.exists(output_file): os.remove(output_file)
    else:
        await update.message.reply_text("لتحويل هذه الصورة إلى PDF، اضغط /start واختر 'صورة ➔ PDF'.")

if __name__ == "__main__":
    print("البوت يعمل الآن بميزة دمج الملفات وكافة التحويلات...")
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.run_polling()
