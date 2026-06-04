import os
import re
import telebot
import logging
from threading import Thread
from flask import Flask
from docx import Document
from docx.shared import Mm, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

# --- 1. HỆ THỐNG GHI LOGS (BẮT LỖI) ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)

# --- 2. CẤU HÌNH TOKEN TELEGRAM CỦA BẠN ---
API_TOKEN = '8772669126:AAFRYcViBI3dG_2MfDkA6VZtGIi-nSm0eBM'
bot = telebot.TeleBot(API_TOKEN)

# --- 3. TẠO WEB SERVER PHỤ (Giúp giữ bot luôn chạy online trên Render) ---
app = Flask('')

@app.route('/')
def home():
    return "Bot định dạng Word đang chạy trực tuyến 24/7!"

def run_web_server():
    port = int(os.environ.get("PORT", 8080))
    app.run(host='0.0.0.0', port=port)

# --- 4. HÀM ĐỊNH DẠNG FILE WORD THEO YÊU CẦU ĐỒ ÁN ---
def set_run_font_times_new_roman(run, size_pt, bold=False):
    """Ép font chữ chuẩn Times New Roman cho hiển thị tiếng Việt không bị lỗi font"""
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')

def format_docx(input_path, output_path):
    """Hàm căn lề, chỉnh font, giãn dòng đúng chuẩn yêu cầu đồ án và sửa lỗi giãn chữ"""
    logging.info(f"Bắt đầu xử lý file: {input_path}")
    doc = Document(input_path)
    
    # Cấu hình khổ giấy A4 và Căn lề chuẩn: Trái 30mm, Phải 15mm, Trên 20mm, Dưới 20mm
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)

    # Duyệt qua từng đoạn văn bản trong file để ép định dạng
    for paragraph in doc.paragraphs:
        # Loại bỏ ký tự xuống dòng mềm (\v, \n) ở đầu/cuối đoạn để trị dứt điểm bệnh giãn chữ của Word
        for run in paragraph.runs:
            if run.text:
                run.text = run.text.replace('\v', ' ').replace('\n', ' ')
                
        text = paragraph.text.strip()
        if not text:
            continue
            
        # --- BỘ LỌC THÔNG MINH NHẬN DIỆN TIÊU ĐỀ CHƯƠNG / VẤN ĐỀ ---
        is_chapter = False
        # Trường hợp 1: Bắt đầu bằng chữ "CHƯƠNG"
        if text.upper().startswith("CHƯƠNG"):
            is_chapter = True
        # Trường hợp 2: File gốc đang được chủ ý Căn giữa sẵn
        elif paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            is_chapter = True
        # Trường hợp 3: Dòng ngắn, viết hoa toàn bộ hoặc có cài hiệu ứng All Caps trong Word
        elif len(text) < 100:
            has_all_caps_run = any(run.font.all_caps for run in paragraph.runs)
            if text.isupper() or has_all_caps_run:
                is_chapter = True

        # --- BỘ LỌC NHẬN DIỆN MỤC CON 1, 2, 1.1, 1.2... ---
        is_section = False
        if not is_chapter and re.match(r'^(\d+(\.\d+)*\.?)\s+', text):
            is_section = True

        # --- TIẾN HÀNH ÁP ĐỊNH DẠNG CHUẨN ---
        if is_chapter:
            # Tên chương/Vấn đề: Chữ in hoa, cao 16, đậm (Bold), căn lề giữa (center)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.text = text.upper()  # Ép ký tự thành viết hoa thật sự thay vì dùng hiệu ứng
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(6)
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            paragraph.paragraph_format.line_spacing = 1.3
            for run in paragraph.runs:
                set_run_font_times_new_roman(run, 16, bold=True)
                
        elif is_section:
            # Các mục 1, 2...: Chữ thường, cỡ 13, đậm (Bold), bảo hiểm dòng ngắn để tránh dãn chữ
            if len(text) < 50:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(3)
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            paragraph.paragraph_format.line_spacing = 1.3
            for run in paragraph.runs:
                set_run_font_times_new_roman(run, 13, bold=True)
                
        else:
            # Các nội dung khác: Cỡ chữ 13, Thụt lề (first line) 1cm, Giãn dòng Multiple 1.3, spacing before 6pt
            # Nếu dòng text quá ngắn (dòng đơn độc lập), để lề Trái để tuyệt đối không bị lỗi kéo dãn chữ
            if len(text) < 50:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
            paragraph.paragraph_format.first_line_indent = Cm(1)
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            paragraph.paragraph_format.line_spacing = 1.3
            for run in paragraph.runs:
                set_run_font_times_new_roman(run, 13, bold=False)

    doc.save(output_path)
    logging.info(f"Đã xử lý lỗi giãn chữ thành công: {output_path}")

# --- 5. ĐIỀU HƯỚNG BOT TELEGRAM ---
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    welcome_text = (
        "👋 Xin chào! Tôi là Bot tự động căn lề báo cáo Đồ án Viễn thông (Bản nâng cấp sửa lỗi giãn chữ).\n\n"
        "📥 Hãy gửi cho tôi file Word (.docx) của bạn. Tôi sẽ tự động chỉnh sửa đúng chuẩn đồ án."
    )
    bot.reply_to(message, welcome_text)

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    input_filename = ""
    output_filename = ""
    try:
        file_name = message.document.file_name
        logging.info(f"Nhận file '{file_name}' từ Chat ID: {message.chat.id}")
        
        if not file_name.endswith('.docx'):
            bot.reply_to(message, "❌ Vui lòng chỉ gửi file định dạng Word (.docx) thôi nhé bạn ơi!")
            return
            
        bot.reply_to(message, "🔄 Đang tiến hành bóc tách lỗi giãn chữ và căn chỉnh lại toàn bộ file, vui lòng đợi vài giây...")
        
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        input_filename = f"input_{message.chat.id}_{file_name}"
        output_filename = f"Formatted_{file_name}"
        
        with open(input_filename, 'wb') as new_file:
            new_file.write(downloaded_file)
            
        format_docx(input_filename, output_filename)
        
        with open(output_filename, 'rb') as formatted_doc:
            bot.send_document(message.chat.id, formatted_doc, caption="✅ Đã sửa lỗi giãn chữ và căn lề xong xuôi theo chuẩn đồ án!")
            
    except Exception as e:
        logging.error(f"LỖI HỆ THỐNG: {str(e)}", exc_info=True)
        bot.reply_to(message, "❌ Có lỗi cấu trúc đặc biệt xảy ra trong file. Bạn hãy thử lưu lại file Word đó rồi gửi lại nhé.")
        
    finally:
        if input_filename and os.path.exists(input_filename):
            os.remove(input_filename)
        if output_filename and os.path.exists(output_filename):
            os.remove(output_filename)

# --- 6. CHẠY ĐỒNG THỜI WEB SERVER VÀ BOT TELEGRAM ---
if __name__ == '__main__':
    logging.info("--- HỆ THỐNG BOT BẮT ĐẦU KÍCH HOẠT ---")
    server_thread = Thread(target=run_web_server)
    server_thread.start()
    bot.infinity_polling()
