import os
import sys
import subprocess
import json
import time
from datetime import datetime

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

import telebot
import openpyxl
import threading
import logging
from flask import Flask, jsonify, render_template_string


# ==============================================================================
# CẤU HÌNH THÔNG TIN BOT TELEGRAM
# Bạn có thể điền trực tiếp vào đây hoặc tạo file 'telegram_config.json'
# ==============================================================================
BOT_TOKEN = "ĐIỀN_BOT_TOKEN_VÀO_ĐÂY"
ALLOWED_CHAT_ID = "ĐIỀN_CHAT_ID_VÀO_ĐÂY"  # Dãy số ID lấy từ @userinfobot (ví dụ: "123456789")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "telegram_config.json")

def is_valid_token(tok):
    return bool(tok and ":" in tok and not tok.startswith("DIEN_") and not tok.startswith("ĐIỀN_"))

# Tự động nạp từ Biến môi trường (Cloud Server / Render / VPS) hoặc file telegram_config.json
BOT_TOKEN = os.environ.get("BOT_TOKEN", BOT_TOKEN)
ALLOWED_CHAT_ID = os.environ.get("ALLOWED_CHAT_ID", ALLOWED_CHAT_ID)

if os.path.exists(CONFIG_FILE):
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            if not is_valid_token(BOT_TOKEN):
                BOT_TOKEN = cfg.get("bot_token", BOT_TOKEN)
            if ALLOWED_CHAT_ID == "ĐIỀN_CHAT_ID_VÀO_ĐÂY":
                ALLOWED_CHAT_ID = str(cfg.get("allowed_chat_id", ALLOWED_CHAT_ID))
    except Exception as e:
        print(f"Lỗi đọc config: {e}")

if not is_valid_token(BOT_TOKEN):
    print("\n" + "=" * 65)
    print("⚠️  CHƯA CẤU HÌNH BOT_TOKEN HOẶC TOKEN KHÔNG HỢP LỆ!")
    print(f"👉 Vui lòng mở file '{os.path.basename(CONFIG_FILE)}' và dán Token vào:")
    print('   {"bot_token": "123456789:ABCdefGhIJKlmNoPQRsTUVwxyZ", ...}')
    print("=" * 65 + "\n")
    sys.exit(1)

bot = telebot.TeleBot(BOT_TOKEN)
user_files = {}

def check_permission(message):
    """Kiểm tra quyền truy cập để đảm bảo chỉ có bạn mới dùng được bot"""
    if "ĐIỀN_" in ALLOWED_CHAT_ID or not ALLOWED_CHAT_ID:
        return True # Nếu chưa set chat id thì tạm thời cho phép để lấy ID
    if str(message.chat.id) != str(ALLOWED_CHAT_ID):
        bot.reply_to(message, f"⛔ Bạn không có quyền sử dụng Bot này.\n(Chat ID của bạn: `{message.chat.id}`)", parse_mode="Markdown")
        return False
    return True

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    chat_id = str(message.chat.id)
    msg = (
        "👋 **Chào mừng chị Cẩm Giang!**\n\n"
        "🤖 **Bot Báo Cáo Doanh Số Team (CVS + BHX)** đã sẵn sàng.\n\n"
        f"🆔 Chat ID của bạn: `{chat_id}`\n\n"
        "📌 **Cách sử dụng rất đơn giản:**\n"
        "1️⃣ Từ điện thoại, hãy gửi 2 file Excel (`.xlsb` hoặc `.xlsx`):\n"
        "   • **File Siêu thị BHX**: tên có chứa `ST_KD6` hoặc `ST`\n"
        "   • **File CVS & NPP**: tên có chứa `KD6` hoặc `KD06`\n"
        "2️⃣ Bot sẽ tự động tải về, tính toán và trả về ngay file `Team_CamGiang_Report.xlsx` kèm tóm tắt KPI trên màn hình.\n\n"
        "📋 **Các lệnh tiện ích:**\n"
        "• `/status` : Kiểm tra các file hiện có\n"
        "• `/chay`   : Bắt buộc chạy báo cáo với các file hiện tại\n"
        "• `/reset`  : Xóa dữ liệu tạm để gửi lại từ đầu"
    )
    bot.reply_to(message, msg, parse_mode='Markdown')

@bot.message_handler(commands=['status'])
def check_status(message):
    if not check_permission(message):
        return
    st_f = user_files.get('st')
    cvs_f = user_files.get('cvs')
    
    st_text = f"✅ `{os.path.basename(st_f)}`" if (st_f and os.path.exists(st_f)) else "❌ Chưa nhận"
    cvs_text = f"✅ `{os.path.basename(cvs_f)}`" if (cvs_f and os.path.exists(cvs_f)) else "❌ Chưa nhận"

    msg = (
        "📊 **Tình trạng file nguồn hiện tại:**\n"
        f"• File Siêu thị (BHX): {st_text}\n"
        f"• File CVS & NPP (CK, GS25...): {cvs_text}\n\n"
    )
    if st_f and cvs_f:
        msg += "👉 **Đã đủ 2 file!** Gõ /chay nếu muốn tạo lại báo cáo ngay."
    else:
        msg += "👉 Hãy gửi tiếp file còn thiếu vào đây nhé."
    bot.reply_to(message, msg, parse_mode='Markdown')

@bot.message_handler(commands=['reset'])
def reset_session(message):
    if not check_permission(message):
        return
    user_files.clear()
    bot.reply_to(message, "🔄 Đã reset danh sách file tạm. Bạn có thể gửi lại 2 file mới.")

@bot.message_handler(commands=['chay'])
def trigger_run(message):
    if not check_permission(message):
        return
    run_pipeline(message)

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    if not check_permission(message):
        return

    doc = message.document
    file_name = doc.file_name

    if not file_name.lower().endswith(('.xlsb', '.xlsx')):
        bot.reply_to(message, "⚠️ Vui lòng chỉ gửi file Excel có định dạng `.xlsb` hoặc `.xlsx`.")
        return

    status_msg = bot.reply_to(message, f"📥 Đang tải file `{file_name}` về máy...", parse_mode='Markdown')

    try:
        file_info = bot.get_file(doc.file_id)
        downloaded = bot.download_file(file_info.file_path)
        save_path = os.path.join(BASE_DIR, file_name)

        with open(save_path, 'wb') as f:
            f.write(downloaded)

        fn_upper = file_name.upper()
        detected_type = ""

        # Nhận diện loại file
        if 'HNTRINH_ST' in fn_upper or 'ST_KD' in fn_upper or 'ST' in fn_upper or 'SIEU THI' in fn_upper:
            user_files['st'] = save_path
            detected_type = "Siêu thị (BHX)"
        elif 'HNTRINH_KD' in fn_upper or 'KD6' in fn_upper or 'KD06' in fn_upper or 'CVS' in fn_upper:
            user_files['cvs'] = save_path
            detected_type = "CVS & NPP"
        else:
            if 'st' not in user_files:
                user_files['st'] = save_path
                detected_type = "Siêu thị (BHX) [Tạm gán]"
            else:
                user_files['cvs'] = save_path
                detected_type = "CVS & NPP [Tạm gán]"

        bot.edit_message_text(
            f"✅ **Đã tải xong:** `{file_name}`\n🏷️ **Phân loại:** {detected_type}",
            chat_id=message.chat.id,
            message_id=status_msg.message_id,
            parse_mode='Markdown'
        )

        # Kiểm tra xem đủ 2 file chưa
        if 'st' in user_files and 'cvs' in user_files:
            run_pipeline(message)
        else:
            bot.send_message(
                message.chat.id,
                "⏳ **Đã nhận 1/2 file.** Vui lòng gửi tiếp file còn lại để bot tự động tính toán!",
                parse_mode='Markdown'
            )

    except Exception as e:
        bot.reply_to(message, f"❌ Lỗi tải file: {str(e)}")

def extract_summary_kpi():
    """Đọc tóm tắt các chỉ số KPI từ file Report vừa tạo"""
    report_path = os.path.join(BASE_DIR, "Team_CamGiang_Report.xlsx")
    if not os.path.exists(report_path):
        return None

    try:
        wb = openpyxl.load_workbook(report_path, data_only=True)
        ws = wb["Dashboard"] if "Dashboard" in wb.sheetnames else wb.active

        # Đọc dữ liệu thẻ KPI
        target = ws["B5"].value or 0
        actual_bhx = ws["C5"].value or 0
        actual_cvs = ws["D5"].value or 0
        actual_total = ws["E5"].value or 0
        pct_achieved = ws["F5"].value or 0
        status_text = ws["B8"].value or ""
        
        return {
            "target": target,
            "bhx": actual_bhx,
            "cvs": actual_cvs,
            "total": actual_total,
            "pct": pct_achieved,
            "status": str(status_text)
        }
    except Exception as e:
        print(f"Lỗi trích xuất KPI: {e}")
        return None

def run_pipeline(message):
    """Thực thi chuỗi xử lý: update_data -> team_leader_report"""
    chat_id = message.chat.id
    bot.send_message(
        chat_id,
        "🚀 **Đã nhận đủ 2 file nguồn! Đang tiến hành xử lý:**\n"
        "1️⃣ Cập nhật dữ liệu BHX & CVS...\n"
        "2️⃣ Tính toán phân bổ 232 CH BHX & chi tiết cửa hàng CVS...\n"
        "3️⃣ Xuất báo cáo Excel chuyên nghiệp 6 sheets...\n\n"
        "⏱️ _Vui lòng đợi khoảng 15 - 30 giây..._",
        parse_mode='Markdown'
    )

    try:
        # Bước 1: Chạy update_data_and_config.py
        p1 = subprocess.run(
            [sys.executable, "update_data_and_config.py"],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )

        if p1.returncode != 0:
            err_snippet = (p1.stderr or p1.stdout)[-500:]
            bot.send_message(chat_id, f"❌ **Lỗi ở bước cập nhật dữ liệu:**\n```{err_snippet}```", parse_mode='Markdown')
            return

        # Bước 2: Chạy team_leader_report_v4.py
        p2 = subprocess.run(
            [sys.executable, "team_leader_report_v4.py"],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )

        if p2.returncode != 0:
            err_snippet = (p2.stderr or p2.stdout)[-500:]
            bot.send_message(chat_id, f"❌ **Lỗi ở bước tạo báo cáo Excel:**\n```{err_snippet}```", parse_mode='Markdown')
            return

        report_file = os.path.join(BASE_DIR, "Team_CamGiang_Report.xlsx")
        if not os.path.exists(report_file):
            bot.send_message(chat_id, "❌ Không tìm thấy file `Team_CamGiang_Report.xlsx` sau khi xuất.")
            return

        # Bước 3: Đọc tóm tắt KPI
        kpi = extract_summary_kpi()
        time_str = datetime.now().strftime("%H:%M ngày %d/%m/%Y")

        summary_msg = f"🎉 **BÁO CÁO DOANH SỐ TEAM ĐÃ HOÀN TẤT!**\n"
        summary_msg += f"🕒 Thời gian: {time_str}\n"
        summary_msg += "━━━━━━━━━━━━━━━━━━━━━━\n"

        if kpi:
            def fmt(v):
                try:
                    return f"{int(round(float(v))):,d}".replace(",", ".")
                except:
                    return str(v)

            pct_val = kpi['pct']
            if isinstance(pct_val, (int, float)):
                pct_str = f"{pct_val * 100:.1f}%" if pct_val <= 2 else f"{pct_val:.1f}%"
            else:
                pct_str = str(pct_val)

            summary_msg += f"🎯 **Target Team:** `{fmt(kpi['target'])}` đ\n"
            summary_msg += f"📦 **Thực hiện BHX:** `{fmt(kpi['bhx'])}` đ\n"
            summary_msg += f"🏪 **Thực hiện CVS:** `{fmt(kpi['cvs'])}` đ\n"
            summary_msg += f"💰 **TỔNG THỰC HIỆN:** `{fmt(kpi['total'])}` đ\n"
            summary_msg += f"📈 **% ĐẠT:** `{pct_str}`\n"
            if kpi['status']:
                summary_msg += f"🚦 **Trạng thái:** {kpi['status']}\n"
            summary_msg += "━━━━━━━━━━━━━━━━━━━━━━\n"

        summary_msg += "📥 _File báo cáo chi tiết Excel được gửi đính kèm ngay dưới đây:_"

        bot.send_message(chat_id, summary_msg, parse_mode='Markdown')

        with open(report_file, 'rb') as f:
            bot.send_document(
                chat_id,
                f,
                caption=f"📊 Team_CamGiang_Report_{datetime.now().strftime('%d%m_%H%M')}.xlsx"
            )

        # Xóa phiên để sẵn sàng cho lần báo cáo tiếp theo
        user_files.clear()

    except Exception as e:
        bot.send_message(chat_id, f"❌ Có lỗi không mong muốn: {str(e)}")

# ==============================================================================
# FLASK WEB SERVER (Phục vụ Render Web Service & UptimeRobot Ping 24/7)
# ==============================================================================
flask_app = Flask(__name__)
logging.getLogger('werkzeug').setLevel(logging.ERROR)

BOT_START_TIME = datetime.now()

HTML_PAGE = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CVS & BHX Telegram Bot Service</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
            color: #333;
        }
        .card {
            background: rgba(255, 255, 255, 0.96);
            backdrop-filter: blur(10px);
            border-radius: 20px;
            padding: 40px 30px;
            max-width: 500px;
            width: 100%;
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.25);
            text-align: center;
        }
        .badge {
            display: inline-flex;
            align-items: center;
            background: #e8f5e9;
            color: #2e7d32;
            padding: 8px 18px;
            border-radius: 50px;
            font-size: 14px;
            font-weight: 700;
            letter-spacing: 0.5px;
            margin-bottom: 20px;
        }
        .pulse-dot {
            width: 10px;
            height: 10px;
            background-color: #2e7d32;
            border-radius: 50%;
            margin-right: 8px;
            animation: pulse 1.8s infinite;
        }
        @keyframes pulse {
            0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(46, 125, 50, 0.7); }
            70% { transform: scale(1); box-shadow: 0 0 0 10px rgba(46, 125, 50, 0); }
            100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(46, 125, 50, 0); }
        }
        h1 {
            font-size: 22px;
            color: #1a202c;
            margin-bottom: 12px;
        }
        p.subtitle {
            color: #4a5568;
            font-size: 14px;
            line-height: 1.6;
            margin-bottom: 20px;
        }
        .info-box {
            background: #f7fafc;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 16px;
            text-align: left;
            font-size: 13px;
            color: #4a5568;
            line-height: 1.8;
        }
        .info-box code {
            background: #edf2f7;
            padding: 2px 6px;
            border-radius: 4px;
            color: #2b6cb0;
            font-family: Consolas, monospace;
        }
        .footer {
            margin-top: 20px;
            font-size: 12px;
            color: #a0aec0;
        }
    </style>
</head>
<body>
    <div class="card">
        <div class="badge"><span class="pulse-dot"></span> ĐANG HOẠT ĐỘNG 24/7</div>
        <h1>🤖 Bot Báo Cáo Doanh Số CVS + BHX</h1>
        <p class="subtitle">Máy chủ Flask Web Server đang chạy song song hỗ trợ UptimeRobot Ping giữ thức và giám sát trạng thái trên Render.</p>
        
        <div class="info-box">
            <div>🕒 <b>Khởi động lúc:</b> {{ start_time }}</div>
            <div>🌐 <b>Cổng lắng nghe (PORT):</b> <code>{{ port }}</code></div>
            <div>📡 <b>Healthcheck URL:</b> <code>/ping</code> hoặc <code>/health</code></div>
            <div>💬 <b>Telegram Bot:</b> <span style="color: #2e7d32; font-weight: bold;">ONLINE</span></div>
        </div>

        <div class="footer">
            CVS & BHX Tracker • Flask Server + TeleBot
        </div>
    </div>
</body>
</html>
"""

@flask_app.route("/")
def home():
    port = os.environ.get("PORT", "8080")
    return render_template_string(
        HTML_PAGE,
        start_time=BOT_START_TIME.strftime("%H:%M:%S ngày %d/%m/%Y"),
        port=port
    ), 200

@flask_app.route("/ping")
@flask_app.route("/health")
def ping():
    return jsonify({
        "status": "ok",
        "service": "CVS & BHX Telegram Bot",
        "bot_online": True,
        "started_at": BOT_START_TIME.isoformat(),
        "current_time": datetime.now().isoformat()
    }), 200

def start_flask_server():
    port = int(os.environ.get("PORT", 8080))
    try:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] 🌐 Đang khởi động Flask Web Server trên cổng {port} cho Render...")
        flask_app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
    except Exception as e:
        print(f"❌ Lỗi khởi động Flask Web Server: {e}")

if __name__ == "__main__":
    if "ĐIỀN_" in BOT_TOKEN:
        print("⚠️ Vui lòng cấu hình BOT_TOKEN trước khi chạy bot.")
        print(f"File config tại: {CONFIG_FILE}")
        sys.exit(1)

    # 1. Khởi động Flask Web Server chạy trên luồng phụ (daemon thread) song song với Telegram Bot
    web_thread = threading.Thread(target=start_flask_server, daemon=True)
    web_thread.start()

    # 2. Khởi động vòng lặp polling nhận tin nhắn từ Telegram
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 🤖 Bot Telegram đang hoạt động và lắng nghe tin nhắn...")
    while True:
        try:
            bot.infinity_polling(timeout=20, long_polling_timeout=20)
        except Exception as e:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Cảnh báo kết nối Telegram: {e}")
            print("Đang tự động thử kết nối lại sau 5 giây...")
            time.sleep(5)

