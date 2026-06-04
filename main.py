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
        logging.StreamHandler()  # Hiển thị trực tiếp trên màn hình máy chủ Render để dễ check lỗi
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
    """Hàm căn lề, chỉnh font, giãn dòng đúng chuẩn yêu cầu đồ án trong ảnh"""
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
        text = paragraph.text.strip()
        if not text:
            continue
            
        # Nhận diện tiêu đề Chương (Bắt đầu bằng chữ chương hoặc viết hoa toàn bộ)
        is_chapter = False
        if text.upper().startswith("CHƯƠNG") or (text.isupper() and len(text) < 100):
            is_chapter = True
            paragraph.text = text.upper()

        # Nhận diện mục con bằng số (Ví dụ: 1., 2., 1.1, 1.2...)
        is_section = False
        if not is_chapter and re.match(r'^(\d+(\.\d+)*\.?)\s+', text):
            is_section = True

        # Tiến hành áp định dạng chuẩn vào từng loại văn bản
        if is_chapter:
            # Tên chương: Chữ in hoa, cao 16, đậm (Bold), căn lề giữa (center)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(6)
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            paragraph.paragraph_format.line_spacing = 1.3
            for run in paragraph.runs:
                set_run_font_times_new_roman(run, 16, bold=True)
                
        elif is_section:
            # Các mục 1, 2...: Chữ thường, cỡ 13, đậm (Bold), để justify
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(3)
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            paragraph.paragraph_format.line_spacing = 1.3
            for run in paragraph.runs:
                set_run_font_times_new_roman(run, 13, bold=True)
                
        else:
            # Các nội dung khác: Cỡ chữ 13, Justify, first line (thụt lề) 1cm
            # Giãn dòng Multiple 1.3, spacing before 6pt
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Cm(1)
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            paragraph.paragraph_format.line_spacing = 1.3
            for run in paragraph.runs:
                set_run_font_times_new_roman(run, 13, bold=False)

    doc.save(output_path)
    logging.info(f"Đã định dạng xong và lưu file sạch: {output_path}")

# --- 5. ĐIỀU HƯỚNG BOT TELEGRAM ---
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    welcome_text = (
        "👋 Xin chào! Tôi là Bot tự động căn lề báo cáo Đồ án Viễn thông.\n\n"
        "📥 Hãy gửi cho tôi file Word (.docx) của bài báo cáo. Tôi sẽ tự động chỉnh sửa đúng chuẩn lề, font chữ, thụt dòng đầu, và kích thước mục con theo yêu cầu!"
    )
    bot.reply_to(message, welcome_text)

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    input_filename = ""
    output_filename = ""
    try:
        file_name = message.document.file_name
        logging.info(f"Nhận yêu cầu từ Chat ID {message.chat.id}, file: '{file_name}'")
        
        # Chỉ chấp nhận file đuôi .docx
        if not file_name.endswith('.docx'):
            bot.reply_to(message, "❌ Vui lòng chỉ gửi file định dạng Word (.docx) thôi nhé bạn ơi!")
            return
            
        bot.reply_to(message, "🔄 Đang nhận file và tự động căn chỉnh lề + font chữ, vui lòng chờ giây lát...")
        
        # Tải file từ Telegram về máy chủ tạm thời
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        input_filename = f"input_{message.chat.id}_{file_name}"
        output_filename = f"Formatted_{file_name}"
        
        with open(input_filename, 'wb') as new_file:
            new_file.write(downloaded_file)
            
        # Tiến hành căn chỉnh lại file Word
        format_docx(input_filename, output_filename)
        
        # Gửi trả file sạch về cho người dùng trên Telegram
        with open(output_filename, 'rb') as formatted_doc:
            bot.send_document(message.chat.id, formatted_doc, caption="✅ Đã định dạng xong xuôi theo đúng yêu cầu đồ án trường bạn rồi nhé!")
            
    except Exception as e:
        logging.error(f"LỖI HỆ THỐNG: {str(e)}", exc_info=True)
        bot.reply_to(message, "❌ Có lỗi xảy ra trong quá trình xử lý file. Bạn hãy kiểm tra lại cấu trúc file nhé.")
        
    finally:
        # Xóa các file rác tạm thời để máy chủ không bị đầy dung lượng
        if input_filename and os.path.exists(input_filename):
            os.remove(input_filename)
        if output_filename and os.path.exists(output_filename):
            os.remove(output_filename)

# --- 6. CHẠY ĐỒNG THỜI WEB SERVER VÀ BOT TELEGRAM ---
if __name__ == '__main__':
    logging.info("--- HỆ THỐNG BOT BẮT ĐẦU KÍCH HOẠT ---")
    # Tạo luồng chạy web server phụ độc lập để giữ kết nối với Render không bị ngắt
    server_thread = Thread(target=run_web_server)
    server_thread.start()
    
    # Bật cổng lắng nghe của Telegram Bot liên tục
    bot.infinity_polling()
