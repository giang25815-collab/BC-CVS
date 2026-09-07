import os
import sys
import subprocess
import json
import time
import csv
import re
from datetime import datetime
from collections import defaultdict

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
user_states = {}  # chat_id -> {'mode': 'AWAITING_TARGET', 'set_at': timestamp}

def get_current_config():
    cfg = {}
    p = os.path.join(BASE_DIR, "team_config.csv")
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                next(reader, None)
                for r in reader:
                    if len(r) >= 2:
                        cfg[r[0].strip()] = r[1].strip()
        except Exception as e:
            print(f"Lỗi đọc config: {e}")
    return cfg

def save_current_config(cfg_dict):
    target_dirs = [BASE_DIR, os.path.join(BASE_DIR, "bot_cvs_deploy"), r'C:\Users\giang\.gemini\config\skills\team-cvs-bhx-tracker']
    rows = [(k, str(v)) for k, v in cfg_dict.items()]
    for td in set(target_dirs):
        if os.path.exists(td):
            p = os.path.join(td, "team_config.csv")
            try:
                with open(p, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["key", "value"])
                    writer.writerows(rows)
            except Exception as e:
                print(f"Lỗi lưu config vào {p}: {e}")

def get_current_employees():
    emp_list = []
    p = os.path.join(BASE_DIR, "team_employees.csv")
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    name = (r.get("employee_name") or "").strip()
                    if name:
                        emp_list.append({
                            "name": name,
                            "bhx_stores": int(float(r.get("bhx_stores_managed", 0) or 0)),
                            "target": int(float(r.get("target", 0) or 0))
                        })
        except Exception as e:
            print(f"Lỗi đọc employees: {e}")
    return emp_list

def save_employees(emp_list):
    target_dirs = [BASE_DIR, os.path.join(BASE_DIR, "bot_cvs_deploy"), r'C:\Users\giang\.gemini\config\skills\team-cvs-bhx-tracker']
    fieldnames = ["employee_name", "bhx_stores_managed", "target"]
    for td in set(target_dirs):
        if os.path.exists(td):
            p = os.path.join(td, "team_employees.csv")
            try:
                with open(p, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    for emp in emp_list:
                        writer.writerow({
                            "employee_name": emp["name"],
                            "bhx_stores_managed": emp["bhx_stores"],
                            "target": emp["target"]
                        })
            except Exception as e:
                print(f"Lỗi lưu employees vào {p}: {e}")

def parse_and_update_target_file(file_path, file_name):
    """
    Phân tích file Target (Excel hoặc CSV):
    - Tự động nhận diện Tháng (từ tên file hoặc sheet)
    - Trích xuất Target của từng nhân viên & số lượng CH BHX
    - Tính tổng Target BHX
    - Cập nhật team_employees.csv và team_config.csv
    """
    fn_lower = file_name.lower()
    detected_month = None
    m = re.search(r'(?:t|thang|tháng)[ _-]?(\d{1,2})', fn_lower)
    if m:
        val = int(m.group(1))
        if 1 <= val <= 12:
            detected_month = val

    cur_emps = {e['name'].strip().lower(): e for e in get_current_employees()}
    updated_emps = {}
    detected_bhx_total = 0.0

    if fn_lower.endswith('.csv'):
        with open(file_path, 'r', encoding='utf-8-sig', errors='replace') as f:
            reader = csv.DictReader(f)
            field_map = {}
            for col in reader.fieldnames or []:
                c_up = col.strip().upper()
                if any(k in c_up for k in ['EMPLOYEE', 'TÊN', 'NAME', 'PG']):
                    field_map['name'] = col
                elif any(k in c_up for k in ['BHX', 'STORE']):
                    field_map['bhx'] = col
                elif any(k in c_up for k in ['TARGET', 'CHỈ TIÊU', 'CHI TIEU']):
                    field_map['target'] = col

            for row in reader:
                name = row.get(field_map.get('name', 'employee_name'), '').strip()
                if not name:
                    continue
                try:
                    bhx = int(float(row.get(field_map.get('bhx', 'bhx_stores_managed'), 0)))
                except:
                    bhx = 0
                try:
                    tgt = int(float(row.get(field_map.get('target', 'target'), 0)))
                except:
                    tgt = 0
                updated_emps[name.lower()] = {
                    'name': name,
                    'bhx_stores': bhx,
                    'target': tgt
                }

    elif fn_lower.endswith(('.xlsx', '.xlsb')):
        import openpyxl
        wb = openpyxl.load_workbook(file_path, data_only=True)
        if not detected_month:
            for sname in wb.sheetnames:
                m2 = re.search(r'(?:t|thang|tháng)[ _-]?(\d{1,2})', sname.lower())
                if m2:
                    val = int(m2.group(1))
                    if 1 <= val <= 12:
                        detected_month = val
                        break

        for sname in wb.sheetnames:
            ws = wb[sname]
            header_row_idx = None
            headers = []
            for r_idx, r in enumerate(ws.iter_rows(values_only=True)):
                str_r = [str(c).strip().upper() if c is not None else '' for c in r]
                if any('PG' in c or 'NHÂN VIÊN' in c or 'HỌ TÊN' in c for c in str_r) and any('CHỈ TIÊU' in c or 'TARGET' in c for c in str_r):
                    header_row_idx = r_idx + 1
                    headers = str_r
                    break

            if header_row_idx:
                pg_col = next((i for i, h in enumerate(headers) if 'PG' in h or 'NHÂN VIÊN' in h or 'HỌ TÊN' in h), None)
                st_col = next((i for i, h in enumerate(headers) if 'TÊN ST' in h or 'STORE' in h or 'CỬA HÀNG' in h), None)
                target_col = next((i for i, h in enumerate(headers) if 'CHỈ TIÊU' in h or 'TARGET' in h), None)
                leader_col = next((i for i, h in enumerate(headers) if 'LEADER' in h or 'QUẢN LÝ' in h), None)

                if pg_col is not None and target_col is not None:
                    pg_targets = defaultdict(float)
                    pg_bhx = defaultdict(int)
                    for r in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
                        if len(r) <= max(pg_col, target_col):
                            continue
                        pg_val = r[pg_col]
                        tgt_val = r[target_col]
                        st_val = r[st_col] if (st_col is not None and len(r) > st_col) else ''
                        lead_val = r[leader_col] if (leader_col is not None and len(r) > leader_col) else ''

                        if pg_val and str(pg_val).strip() and not str(pg_val).strip().upper().startswith(('TỔNG', 'TỔNG', 'TOTAL')):
                            pg_clean = str(pg_val).strip()
                            is_cg = False
                            if lead_val and 'CẨM GIANG' in str(lead_val).upper():
                                is_cg = True
                            elif pg_clean.lower() in cur_emps:
                                is_cg = True
                            elif not leader_col:
                                is_cg = True

                            if is_cg:
                                try:
                                    amt = float(tgt_val) if tgt_val is not None else 0.0
                                except:
                                    amt = 0.0
                                pg_targets[pg_clean] += amt
                                if st_val and ('BÁCH HOÁ XANH' in str(st_val).upper() or 'BHX' in str(st_val).upper()):
                                    pg_bhx[pg_clean] += 1
                                    detected_bhx_total += amt

                    if pg_targets:
                        for pg, tot_t in pg_targets.items():
                            updated_emps[pg.lower()] = {
                                'name': pg,
                                'bhx_stores': pg_bhx[pg],
                                'target': int(round(tot_t))
                            }
                        break

    if not updated_emps:
        return False, "Không tìm thấy dữ liệu Target/Chỉ tiêu hợp lệ của Team trong file.", None

    final_emps = []
    for key, old_e in cur_emps.items():
        if key in updated_emps:
            final_emps.append(updated_emps.pop(key))
        else:
            matched_key = None
            for uk in list(updated_emps.keys()):
                if uk in key or key in uk:
                    matched_key = uk
                    break
            if matched_key:
                m_emp = updated_emps.pop(matched_key)
                m_emp['name'] = old_e['name']
                final_emps.append(m_emp)
            else:
                final_emps.append(old_e)

    for uk, new_e in updated_emps.items():
        final_emps.append(new_e)

    save_employees(final_emps)

    cfg = get_current_config()
    if detected_month:
        cfg['month'] = str(detected_month)
    else:
        detected_month = int(cfg.get('month', datetime.now().month))

    cur_year = int(cfg.get('year', datetime.now().year))
    if detected_bhx_total > 0:
        cfg['total_target_bhx'] = str(int(round(detected_bhx_total)))

    try:
        from team_leader_report_v4 import get_realtime_timegone
        cfg['time_percentage'] = str(get_realtime_timegone(detected_month, cur_year))
    except Exception:
        pass

    save_current_config(cfg)

    return True, {
        'month': detected_month,
        'year': cur_year,
        'employees': final_emps,
        'target_bhx': int(float(cfg.get('total_target_bhx', 0))),
        'timegone': float(cfg.get('time_percentage', 0))
    }, None

def check_permission(message):
    """Kiểm tra quyền truy cập để đảm bảo chỉ có bạn mới dùng được bot"""
    if "ĐIỀN_" in ALLOWED_CHAT_ID or not ALLOWED_CHAT_ID:
        return True
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
        "📌 **1. CHẠY BÁO CÁO HÀNG NGÀY:**\n"
        "Gửi trực tiếp 2 file Excel doanh số (`.xlsb` hoặc `.xlsx`):\n"
        "• **File Siêu thị BHX**: tên có chứa `ST_KD6` hoặc `ST`\n"
        "• **File CVS & NPP**: tên có chứa `KD6` hoặc `KD06`\n"
        "👉 Bot sẽ tự động tính toán và trả file `Team_CamGiang_Report.xlsx` trong 15-30s.\n\n"
        "📌 **2. CẬP NHẬT TARGET / CHỈ TIÊU THÁNG MỚI:**\n"
        "• `/up_target` : Bật chế độ nạp file Target/Chỉ tiêu tháng mới (tránh nhầm lẫn)\n"
        "• `/kpi`       : Xem bảng Target & % Timegone hiện tại của 10 nhân viên\n"
        "• `/thang 9`   : Chuyển nhanh chu kỳ báo cáo sang Tháng 9\n\n"
        "📋 **Các lệnh tiện ích khác:**\n"
        "• `/status` : Kiểm tra tình trạng 2 file doanh số hiện tại\n"
        "• `/chay`   : Bắt buộc chạy báo cáo với các file hiện có\n"
        "• `/reset`  : Xóa dữ liệu tạm để gửi lại từ đầu\n"
        "• `/cancel` : Hủy chế độ nạp Target nếu đang bật"
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

@bot.message_handler(commands=['up_target', 'target'])
def cmd_up_target(message):
    if not check_permission(message):
        return
    chat_id = str(message.chat.id)
    user_states[chat_id] = {'mode': 'AWAITING_TARGET', 'set_at': time.time()}
    msg = (
        "🎯 **CHẾ ĐỘ NẠP FILE TARGET / CHỈ TIÊU THÁNG MỚI ĐÃ BẬT!**\n\n"
        "👉 **Chị hãy gửi file Excel hoặc CSV Target mới vào đây.**\n"
        "_(Ví dụ: `DOANH SỐ T9.xlsx`, `Target_T9.xlsx`, `team_employees.csv`...)_\n\n"
        "💡 **Ưu điểm khi dùng lệnh này:**\n"
        "• **Không sợ nhầm lẫn:** File này sẽ được nhận diện riêng biệt, KHÔNG đè vào file doanh số hàng ngày!\n"
        "• **Tự động chuyển tháng:** Tự nhận diện Tháng 9 và tính % Timegone theo 30 ngày chuẩn xác.\n"
        "• **Tự động trích xuất:** Cập nhật Target và số lượng CH BHX của từng nhân viên.\n"
        "• **Hoạt động ngay lập tức:** Không cần lên GitHub, không cần deploy lại máy chủ!\n\n"
        "*(Nếu đổi ý, chị chỉ cần gõ /cancel để quay lại)*"
    )
    bot.reply_to(message, msg, parse_mode='Markdown')

@bot.message_handler(commands=['cancel'])
def cmd_cancel(message):
    if not check_permission(message):
        return
    chat_id = str(message.chat.id)
    if chat_id in user_states:
        user_states.pop(chat_id, None)
        bot.reply_to(message, "🔄 Đã hủy chế độ nạp Target. Bot quay lại trạng thái nhận file doanh số hàng ngày bình thường.")
    else:
        bot.reply_to(message, "ℹ️ Hiện không có lệnh nào đang chờ xử lý.")

@bot.message_handler(commands=['kpi'])
def cmd_kpi(message):
    if not check_permission(message):
        return
    cfg = get_current_config()
    emps = get_current_employees()
    m = cfg.get('month', '8')
    y = cfg.get('year', '2026')
    t_bhx = int(float(cfg.get('total_target_bhx', 0)))
    time_pct = float(cfg.get('time_percentage', 0))

    def fmt_vnd(n):
        return f"{int(round(float(n))):,d}".replace(",", ".")

    tot_emp_tgt = sum(e['target'] for e in emps)

    msg = f"📊 **CẤU HÌNH TARGET & KPI HIỆN TẠI**\n"
    msg += f"🗓️ **Chu kỳ báo cáo:** Tháng {m}/{y}\n"
    msg += f"⏳ **% Timegone:** `{time_pct:.1f}%`\n"
    msg += f"🎯 **Tổng Target BHX:** `{fmt_vnd(t_bhx)}` đ\n"
    msg += f"💰 **Tổng Target Team ({len(emps)} NV):** `{fmt_vnd(tot_emp_tgt)}` đ\n"
    msg += "━━━━━━━━━━━━━━━━━━━━━━\n"
    msg += "👥 **Chỉ tiêu từng nhân viên:**\n"
    for i, e in enumerate(emps, 1):
        msg += f"{i}. {e['name']}: `{fmt_vnd(e['target'])}` đ ({e['bhx_stores']} CH BHX)\n"
    msg += "━━━━━━━━━━━━━━━━━━━━━━\n"
    msg += "💡 _Để nạp file Target mới: gõ `/up_target` rồi gửi file._\n"
    msg += "💡 _Để đổi tháng nhanh: gõ `/thang <số tháng>` (VD: `/thang 9`)._"
    bot.reply_to(message, msg, parse_mode='Markdown')

@bot.message_handler(commands=['thang', 'set_month'])
def cmd_set_month(message):
    if not check_permission(message):
        return
    args = message.text.strip().split()
    if len(args) < 2:
        bot.reply_to(message, "⚠️ Vui lòng nhập số tháng muốn đổi. Ví dụ: `/thang 9` hoặc `/thang 9/2026`", parse_mode='Markdown')
        return
    val = args[1].strip()
    new_m = None
    new_y = None
    if '/' in val:
        parts = val.split('/')
        try:
            new_m = int(parts[0])
            new_y = int(parts[1])
        except:
            pass
    else:
        try:
            new_m = int(val)
        except:
            pass

    if not new_m or not (1 <= new_m <= 12):
        bot.reply_to(message, "⚠️ Tháng không hợp lệ (phải từ 1 đến 12). Ví dụ: `/thang 9`", parse_mode='Markdown')
        return

    cfg = get_current_config()
    cfg['month'] = str(new_m)
    if new_y:
        cfg['year'] = str(new_y)
    else:
        new_y = int(cfg.get('year', datetime.now().year))

    try:
        from team_leader_report_v4 import get_realtime_timegone
        cfg['time_percentage'] = str(get_realtime_timegone(new_m, new_y))
    except Exception:
        pass

    save_current_config(cfg)
    bot.reply_to(
        message,
        f"✅ **Đã chuyển chu kỳ báo cáo sang Tháng {new_m}/{new_y}!**\n"
        f"⏳ % Timegone hiện tại: `{float(cfg.get('time_percentage', 0)):.1f}%`.\n\n"
        f"👉 Từ bây giờ khi gửi 2 file doanh số, hệ thống sẽ tính toán theo số ngày của Tháng {new_m}!",
        parse_mode='Markdown'
    )

@bot.message_handler(commands=['target_bhx'])
def cmd_target_bhx(message):
    if not check_permission(message):
        return
    args = message.text.strip().split()
    if len(args) < 2:
        bot.reply_to(message, "⚠️ Vui lòng nhập số tiền Target BHX. Ví dụ: `/target_bhx 37382000000`", parse_mode='Markdown')
        return
    num_str = re.sub(r'[,.\sđ]', '', args[1])
    try:
        val = int(float(num_str))
    except:
        bot.reply_to(message, "⚠️ Số tiền không hợp lệ. Ví dụ: `/target_bhx 37382000000`", parse_mode='Markdown')
        return

    cfg = get_current_config()
    cfg['total_target_bhx'] = str(val)
    save_current_config(cfg)
    def fmt_vnd(n):
        return f"{int(round(float(n))):,d}".replace(",", ".")
    bot.reply_to(message, f"✅ **Đã cập nhật Tổng Target BHX:** `{fmt_vnd(val)}` đ.", parse_mode='Markdown')

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

    chat_id = str(message.chat.id)
    doc = message.document
    file_name = doc.file_name
    fn_lower = file_name.lower()
    fn_upper = file_name.upper()

    if not fn_lower.endswith(('.xlsb', '.xlsx', '.csv')):
        bot.reply_to(message, "⚠️ Vui lòng chỉ gửi file Excel (`.xlsx`, `.xlsb`) hoặc `.csv`.")
        return

    # 1. KIỂM TRA XEM CÓ ĐANG TRONG CHẾ ĐỘ NẠP FILE TARGET KHÔNG
    state = user_states.get(chat_id, {})
    is_awaiting_target = (state.get('mode') == 'AWAITING_TARGET')

    if is_awaiting_target:
        user_states.pop(chat_id, None)
        status_msg = bot.reply_to(message, f"📥 Đang tải và phân tích file Target `{file_name}`...", parse_mode='Markdown')
        try:
            file_info = bot.get_file(doc.file_id)
            downloaded = bot.download_file(file_info.file_path)
            save_path = os.path.join(BASE_DIR, file_name)
            with open(save_path, 'wb') as f:
                f.write(downloaded)

            success, data, err = parse_and_update_target_file(save_path, file_name)
            if not success:
                bot.edit_message_text(
                    f"❌ **Không thể nạp Target:** {data}\n\n"
                    "👉 Chị vui lòng kiểm tra lại file (cần có cột Tên nhân viên/PG và Target/Chỉ tiêu), hoặc gõ `/up_target` để thử lại.",
                    chat_id=message.chat.id,
                    message_id=status_msg.message_id,
                    parse_mode='Markdown'
                )
                return

            def fmt_vnd(n):
                return f"{int(round(float(n))):,d}".replace(",", ".")

            res_msg = (
                f"🎉 **ĐÃ CẬP NHẬT TARGET THÁNG {data['month']}/{data['year']} THÀNH CÔNG!**\n"
                f"🗓️ Chu kỳ báo cáo: **Tháng {data['month']}/{data['year']}**\n"
                f"⏳ % Timegone: `{data['timegone']:.1f}%`\n"
                f"🎯 **Tổng Target BHX:** `{fmt_vnd(data['target_bhx'])}` đ\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"👥 **Chỉ tiêu {len(data['employees'])} nhân viên:**\n"
            )
            for i, e in enumerate(data['employees'], 1):
                res_msg += f"{i}. {e['name']}: `{fmt_vnd(e['target'])}` đ ({e['bhx_stores']} CH BHX)\n"
            res_msg += (
                "━━━━━━━━━━━━━━━━━━━━━━\n"
                "✅ Toàn bộ Target mới đã được lưu vào hệ thống.\n"
                "👉 Từ bây giờ, chị cứ gửi 2 file doanh số hàng ngày vào bình thường, bot sẽ tự tính toán theo Target Tháng mới này!"
            )
            bot.edit_message_text(
                res_msg,
                chat_id=message.chat.id,
                message_id=status_msg.message_id,
                parse_mode='Markdown'
            )
        except Exception as e:
            bot.edit_message_text(
                f"❌ Lỗi xử lý file Target: {str(e)}",
                chat_id=message.chat.id,
                message_id=status_msg.message_id
            )
        return

    # 2. NẾU KHÔNG Ở CHẾ ĐỘ UP_TARGET: CẢNH BÁO THÔNG MINH TRÁNH NHẦM FILE TARGET
    is_target_keyword = any(k in fn_upper for k in ['TARGET', 'CHI_TIEU', 'CHI TIEU', 'CHỈ TIÊU', 'DOANH SỐ T', 'DOANH SO T', 'KPI'])
    is_daily_sales = any(k in fn_upper for k in ['HNTRINH', 'ST_KD6', 'KD6', 'KD06', 'SIEU THI'])

    if is_target_keyword and not is_daily_sales:
        bot.reply_to(
            message,
            "⚠️ **Phát hiện file có tên giống File Target / Chỉ tiêu!**\n\n"
            "👉 Để tránh nhầm lẫn với 2 file doanh số hàng ngày, chị vui lòng gõ lệnh `/up_target` trước rồi gửi file này nhé.\n\n"
            "_(Nếu đây là file doanh số hàng ngày, chị hãy kiểm tra lại tên file để có chứa `ST` hoặc `KD6` nhé)_",
            parse_mode='Markdown'
        )
        return

    # 3. LUỒNG XỬ LÝ 2 FILE DOANH SỐ HÀNG NGÀY
    status_msg = bot.reply_to(message, f"📥 Đang tải file `{file_name}` về máy...", parse_mode='Markdown')

    try:
        file_info = bot.get_file(doc.file_id)
        downloaded = bot.download_file(file_info.file_path)
        save_path = os.path.join(BASE_DIR, file_name)

        with open(save_path, 'wb') as f:
            f.write(downloaded)

        detected_type = ""

        # Nhận diện loại file doanh số
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

    # 2. Đăng ký menu lệnh gợi ý trên giao diện Telegram
    try:
        bot.set_my_commands([
            telebot.types.BotCommand("start", "Hướng dẫn sử dụng"),
            telebot.types.BotCommand("status", "Kiểm tra 2 file doanh số"),
            telebot.types.BotCommand("up_target", "Nạp file Target/Chỉ tiêu tháng mới"),
            telebot.types.BotCommand("kpi", "Xem Target & % Timegone của 10 NV"),
            telebot.types.BotCommand("thang", "Đổi tháng báo cáo (VD: /thang 9)"),
            telebot.types.BotCommand("chay", "Tạo báo cáo với file hiện có"),
            telebot.types.BotCommand("reset", "Xóa file tạm để gửi lại"),
            telebot.types.BotCommand("cancel", "Hủy nạp Target")
        ])
    except Exception as e:
        print(f"Lỗi thiết lập Bot Commands: {e}")

    # 3. Khởi động vòng lặp polling nhận tin nhắn từ Telegram
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 🤖 Bot Telegram đang hoạt động và lắng nghe tin nhắn...")
    while True:
        try:
            bot.infinity_polling(timeout=20, long_polling_timeout=20)
        except Exception as e:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Cảnh báo kết nối Telegram: {e}")
            print("Đang tự động thử kết nối lại sau 5 giây...")
            time.sleep(5)

