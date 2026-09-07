"""
TEAM LEADER REPORT V4 - Tran Thi Cam Giang
Description: Theo doi doanh so team = BHX (co dinh) + CVS (chi tiet tung cua hang)
             So sanh % Dat vs % Timegone
             Them sheet chi tiet theo doi Doanh so & SKU FamilyMart cua Nguyen Duc Hoa
Version: 4.1
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.utils import get_column_letter
from datetime import datetime
import calendar
import csv
import sys
import os
import io
import glob

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


# ==========================================
# TIMEGONE HELPER
# ==========================================
def get_realtime_timegone(month=None, year=None):
    """
    Tinh % Timegone tu dong theo thoi gian thuc:
    Vi du: Hom nay ngay 29/8 trong thang 8 co 31 ngay -> (29 / 31) * 100 = 93.55%
    """
    now = datetime.now()
    m = int(month) if month else now.month
    y = int(year) if year else now.year
    _, max_days = calendar.monthrange(y, m)
    if y == now.year and m == now.month:
        day = min(now.day, max_days)
    elif (y, m) < (now.year, now.month):
        day = max_days
    else:
        day = 0
    return round((day / max_days) * 100, 2)


# ==========================================
# CSV READERS
# ==========================================
def read_config(csv_file, auto_realtime_timegone=True):
    config = {}
    with open(csv_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            key = row['key'].strip()
            value = row['value'].strip()
            if key in ['month', 'year', 'total_target_bhx', 'total_actual_bhx', 'total_stores_bhx']:
                config[key] = int(float(value))
            elif key in ['time_percentage']:
                config[key] = float(value)
            else:
                config[key] = value

    if auto_realtime_timegone:
        m = config.get('month')
        y = config.get('year')
        config['time_percentage'] = get_realtime_timegone(m, y)

    return config


def read_team_employees(csv_file):
    employees = []
    with open(csv_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            employees.append({
                'name': row['employee_name'].strip(),
                'bhx_stores': int(row['bhx_stores_managed']),
                'target': int(row['target']),
            })
    return employees


def read_cvs_stores(csv_file):
    """Doc CVS stores chi tiet -> dict {employee: [store_dicts]}"""
    cvs_data = {}
    with open(csv_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            emp = row['employee_name'].strip()
            if emp not in cvs_data:
                cvs_data[emp] = []
            cvs_data[emp].append({
                'chain': row['chain'].strip(),
                'store_code': row['store_code'].strip(),
                'store_address': row['store_address'].strip(),
                'actual': int(float(row['actual'])),
            })
    return cvs_data


# ==========================================
# CORE LOGIC
# ==========================================
def calc_bhx_allocation(total_actual, total_stores, stores_managed):
    return round((total_actual / total_stores) * stores_managed)


def calc_employee_summary(emp, config, cvs_data):
    """Tinh tong ket cua NV"""
    bhx_actual = calc_bhx_allocation(
        config['total_actual_bhx'],
        config['total_stores_bhx'],
        emp['bhx_stores']
    )

    emp_stores = cvs_data.get(emp['name'], [])
    chains = {'GS25': 0, '7-Eleven': 0, 'FamilyMart': 0, 'Circle K': 0, 'Hoàng Đức': 0, 'WinMart+': 0}
    for store in emp_stores:
        ch = store['chain']
        if ch in chains:
            chains[ch] += store['actual']
        elif 'Hoàng Đức' in ch or 'Hoang Duc' in ch:
            chains['Hoàng Đức'] += store['actual']
        elif 'Circle' in ch or 'CK' in ch:
            chains['Circle K'] += store['actual']
        elif '7E' in ch or '7-Eleven' in ch:
            chains['7-Eleven'] += store['actual']
        elif 'Family' in ch or 'FM' in ch:
            chains['FamilyMart'] += store['actual']
        elif 'GS25' in ch:
            chains['GS25'] += store['actual']
        elif 'Win' in ch or 'WMP' in ch:
            chains['WinMart+'] += store['actual']

    cvs_total = sum(s['actual'] for s in emp_stores)
    total_actual = bhx_actual + cvs_total
    percent = (total_actual / emp['target'] * 100) if emp['target'] > 0 else 0

    return {
        'bhx_actual': bhx_actual,
        'cvs_total': cvs_total,
        'chains': chains,
        'total_actual': total_actual,
        'missing': emp['target'] - total_actual,
        'percent': percent,
        'stores': emp_stores,
    }


# ==========================================
# STYLING HELPERS
# ==========================================
def style_header_cell(cell, bg='2E75B6', color='FFFFFF', size=11, bold=True):
    cell.font = Font(bold=bold, color=color, size=size)
    cell.fill = PatternFill('solid', fgColor=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def style_money(cell, negative_red=True):
    if negative_red:
        cell.number_format = '#,##0;[Red](#,##0)'
    else:
        cell.number_format = '#,##0'
    cell.alignment = Alignment(horizontal='right', vertical='center')


def add_borders(ws, cell_range):
    side = Side(border_style='thin', color='B4B4B4')
    for row in ws[cell_range]:
        for cell in row:
            cell.border = Border(top=side, left=side, right=side, bottom=side)


COLOR_RED = 'DD2323'
COLOR_GREEN = '28E11F'
COLOR_ORANGE = 'F5991F'


def apply_timegone_conditional_formatting(ws, cell_range, time_pct):
    """
    Quy tắc màu sắc so sánh với % Timegone theo bảng màu chuẩn của User:
    1. Vượt / Đạt Timegone (>= Timegone): Màu Xanh Lá (#28E11F), chữ đen bold
    2. Cận Timegone ít (>= 75% Timegone và < Timegone): Màu Cam Cảnh Báo (#F5991F), chữ trắng bold
    3. Thấp hơn Timegone nhiều (< 75% Timegone): Màu Đỏ (#DD2323), chữ trắng bold
    """
    tg = round(time_pct / 100, 4)
    warn_threshold = round(tg * 0.75, 4)

    green_fill = PatternFill(start_color=COLOR_GREEN, end_color=COLOR_GREEN, fill_type='solid')
    green_font = Font(color='000000', bold=True)
    orange_fill = PatternFill(start_color=COLOR_ORANGE, end_color=COLOR_ORANGE, fill_type='solid')
    orange_font = Font(color='FFFFFF', bold=True)
    red_fill = PatternFill(start_color=COLOR_RED, end_color=COLOR_RED, fill_type='solid')
    red_font = Font(color='FFFFFF', bold=True)

    rule_green = CellIsRule(operator='greaterThanOrEqual', formula=[str(tg)], stopIfTrue=True, fill=green_fill, font=green_font)
    rule_orange = CellIsRule(operator='between', formula=[str(warn_threshold), str(tg)], stopIfTrue=True, fill=orange_fill, font=orange_font)
    rule_red = CellIsRule(operator='lessThan', formula=[str(warn_threshold)], stopIfTrue=True, fill=red_fill, font=red_font)

    ws.conditional_formatting.add(cell_range, rule_green)
    ws.conditional_formatting.add(cell_range, rule_orange)
    ws.conditional_formatting.add(cell_range, rule_red)


CHAIN_COLORS = {
    'GS25': '4472C4',
    '7-Eleven': '00B050',
    'FamilyMart': 'ED7D31',
    'Circle K': 'C00000',
    'Hoàng Đức Long Khánh': '70AD47',
    'Hoàng Đức Gia Kiệm': '008080',
    'Hoàng Đức': '70AD47',
    'WinMart+': 'D9534F',
    'WMP': 'D9534F',
}


# ==========================================
# SKU DATA PARSER & MATRIX BUILDER
# ==========================================
def get_fm_sku_matrix():
    """
    Lay du lieu chi tiet toan bo CH FamilyMart cua tat ca nhan vien va 46 SKU thuc te da dat hang
    Tu file DOANH SO T8.xlsx va HNTRINH_...xlsb
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    doanhso_candidates = [
        os.path.join(script_dir, 'DOANH SỐ T8.xlsx'),
        'DOANH SỐ T8.xlsx',
    ] + glob.glob(os.path.join(script_dir, '*DOANH S*.xlsx')) + glob.glob('*DOANH S*.xlsx')
    doanhso_candidates = [f for f in doanhso_candidates if not os.path.basename(f).startswith('~$') and not os.path.basename(f).startswith('~')]
    doanhso_path = next((f for f in doanhso_candidates if os.path.exists(f)), None)

    xlsb_candidates = (
        glob.glob(os.path.join(script_dir, '*HNTRINH*.xlsb'))
        + glob.glob('*HNTRINH*.xlsb')
        + glob.glob(os.path.join(os.path.dirname(script_dir), '*HNTRINH*.xlsb'))
        + glob.glob(r'C:\Users\giang\Downloads\*HNTRINH*.xlsb')
        + glob.glob(r'C:\Users\giang\Downloads\DOANH SỐ\*HNTRINH*.xlsb')
    )
    xlsb_candidates = [f for f in set(xlsb_candidates) if not os.path.basename(f).startswith('~$') and not os.path.basename(f).startswith('~')]
    
    valid_xlsb = []
    for f in xlsb_candidates:
        try:
            import pyxlsb
            with pyxlsb.open_workbook(f) as wb:
                if 'NPP' in wb.sheets:
                    with wb.get_sheet('NPP') as s:
                        count = sum(1 for i, r in enumerate(s.rows()) if i < 10 and any(c.v for c in r))
                        if count > 2:
                            valid_xlsb.append((os.path.getmtime(f), f))
        except Exception:
            pass
    valid_xlsb.sort(reverse=True)
    xlsb_path = valid_xlsb[0][1] if valid_xlsb else (sorted(xlsb_candidates, key=lambda x: os.path.getmtime(x), reverse=True)[0] if xlsb_candidates else None)

    all_fm_stores = []
    if doanhso_path and os.path.exists(doanhso_path):
        wb_doanhso = openpyxl.load_workbook(doanhso_path, data_only=True)
        ws1 = wb_doanhso['Trang_tính1']
        for r in range(2, ws1.max_row + 1):
            st_name = ws1.cell(r, 6).value
            if st_name == 'Family Mart':
                code = str(ws1.cell(r, 7).value).strip() if ws1.cell(r, 7).value else ''
                store_code = str(ws1.cell(r, 8).value).strip() if ws1.cell(r, 8).value else ''
                addr = str(ws1.cell(r, 11).value).strip() if ws1.cell(r, 11).value else ''
                location = str(ws1.cell(r, 12).value or 'ST00_CVS_FM').strip()
                ma_pg = str(ws1.cell(r, 13).value or '').strip()
                pg_name = str(ws1.cell(r, 14).value or '').strip()
                vtcv = str(ws1.cell(r, 15).value or 'SR').strip()
                tgt = float(ws1.cell(r, 20).value or 0)
                all_fm_stores.append({
                    'code': code,
                    'store_code': store_code,
                    'addr': addr,
                    'location': location,
                    'ma_pg': ma_pg,
                    'ten_pg': pg_name,
                    'vtcv': vtcv,
                    'target': tgt,
                })

    npp_data = {}
    if xlsb_path and os.path.exists(xlsb_path):
        try:
            import pyxlsb
            with pyxlsb.open_workbook(xlsb_path) as wb_xlsb:
                with wb_xlsb.get_sheet('NPP') as s:
                    for i, row in enumerate(s.rows()):
                        if i == 0:
                            continue
                        otype = str(row[0].v or '')
                        oname = str(row[5].v or '')
                        if 'FM' in otype or 'Family' in oname:
                            ocode = str(row[4].v or '').strip()
                            oaddr = str(row[7].v or '').strip()
                            icode = str(row[8].v or '').strip()
                            iname = str(row[9].v or '').strip()
                            qty = float(row[11].v or 0)
                            amt = float(row[12].v or 0)

                            if ocode not in npp_data:
                                npp_data[ocode] = {'addr': oaddr, 'items': {}, 'total_amt': 0}
                            if icode not in npp_data[ocode]['items']:
                                npp_data[ocode]['items'][icode] = {'name': iname, 'qty': 0, 'amt': 0}
                            npp_data[ocode]['items'][icode]['qty'] += qty
                            npp_data[ocode]['items'][icode]['amt'] += amt
                            npp_data[ocode]['total_amt'] += amt
        except Exception as e:
            print(f"[WARNING] Could not parse xlsb: {e}")

    # Build store-order matrix
    sku_dict = {}
    store_matrix = []

    for s in all_fm_stores:
        code = s['code']
        matched_npp = None
        if code in npp_data:
            matched_npp = npp_data[code]
        else:
            suffix = code[-4:] if len(code) >= 4 else code
            candidates = [k for k in npp_data if k.endswith(suffix)]
            if len(candidates) == 1:
                matched_npp = npp_data[candidates[0]]
            else:
                st_street = s['addr'].split(',')[0].strip().lower()
                if len(st_street) > 3:
                    for n_code, n_val in npp_data.items():
                        if st_street in n_val['addr'].lower() or n_val['addr'].lower() in s['addr'].lower():
                            matched_npp = n_val
                            break

        orders = {}
        if matched_npp:
            for icode, idata in matched_npp['items'].items():
                orders[icode] = idata['amt']
                if icode not in sku_dict:
                    sku_dict[icode] = {'name': idata['name'], 'total': 0}
                sku_dict[icode]['total'] += idata['amt']
        store_matrix.append({'store': s, 'orders': orders, 'actual': sum(orders.values())})

    # Sort store_matrix by ten_pg, then addr
    store_matrix.sort(key=lambda x: (x['store']['ten_pg'] or 'ZZZ', x['store']['addr'] or ''))

    # Group SKUs into logical categories
    categories = [
        ('TRỌNG TÂM / GREEN FARM & CAO ĐẠM', 'C6E0B4', 'E2EFDA', [
            ('07OI20', 'SCU Vinamilk Green Farm 200ml'),
            ('04GP22', 'STT cao đạm ít béo VNM Green Farm 250ml'),
            ('04GI14', 'STTT rất ít đường VNM Green Farm 180ml'),
            ('07FC10', 'SCA GF Topping Hy Lạp Cao Đạm Cà Phê 121g'),
            ('07FM10', 'SCA GF Topping Hy Lạp Cao Đạm Mật Ong 123g'),
            ('07TP11', 'Sữa chua cao đạm Vinamilk 100g'),
        ]),
        ('KEM VINAMILK', 'F8CBAD', 'FCE4D6', [
            ('09VN01', 'Kem viên hương vani Vinamilk 48g'),
            ('09VC01', 'Kem viên sôcôla Vinamilk 48g'),
            ('09NM07', 'Kem mịn Khoai môn VNM 400ml'),
            ('09LD01', 'Kem Gelato Dừa Vinamilk 90ml (12L/T)'),
            ('09LM01', 'Kem Gelato Matcha Vinamilk 90ml (12L/T)'),
            ('09ND08', 'Kem mịn Dừa VNM 400ml'),
        ]),
        ('SỮA CHUA & PROBI', 'FFE699', 'FFF2CC', [
            ('07KD12', 'SCA không đường VNM 100g'),
            ('07ID11', 'SCA ít đường VNM 100g'),
            ('07ND21', 'Sữa chua Nha đam Vinamilk 100g'),
            ('07UI31', 'SCU MS ít đường VNM Probi 130ml'),
            ('07UR31', 'SCU MS vị truyền thống VNM Probi 130ml'),
            ('07TR33', 'SCA Có đường VNM 100g'),
            ('07LI11', 'SCA Lựu đỏ IĐ VNM Collagen 100g (24H/T)'),
            ('07UI40', 'SCU MS ít đường VNM Probi 400ml'),
            ('07VQ11', 'SCA Việt Quất IĐ VNM 100g (24H/T)'),
        ]),
        ('SỮA ĐẶC & CREAMER', 'BDD7EE', 'DDEBF7', [
            ('01SX11', 'Creamer đặc có đường NSPN xanh lá 1284g.'),
            ('01SX07', 'Creamer đặc có đường NSPN Xanh lá HG 380g'),
            ('01TD60', 'SĐCĐ Ông Thọ đỏ tuýp 165g'),
            ('01VD41', 'SĐCĐ Ông Thọ đỏ vỉ 40g.'),
            ('01SX05', 'Creamer đặc có đường NSPN xanh lá 380g.'),
        ]),
        ('SỮA TƯƠI TIỆT TRÙNG', 'B4C6E7', 'D9E1F2', [
            ('04ET04', 'STTT không đường VNM 1L'),
            ('04EC13', 'STTT sôcôla VNM 180ml'),
            ('04ED04', 'STTT có đường VNM 1L'),
            ('04FT2H', 'STTT không đường VNM 100% Sữa tươi F220ml'),
            ('04FD2H', 'STTT có đường VNM 100% Sữa tươi F220ml'),
            ('04FT32', 'SDD không đường Vinamilk F220ml'),
            ('04EI04', 'STTT ít đường VNM 1L'),
            ('04ET13', 'STTT không đường VNM 180ml'),
            ('04EB11', 'STTT giảm béo chuối VNM 180ml'),
            ('04ED13', 'STTT có đường VNM 180ml'),
        ]),
        ('SỮA HẠT & NƯỚC ÉP', 'D9D2E9', 'EAD1DC', [
            ('05AN25', 'Sữa hạt 9 loại hạt Vinamilk 180ml (24H/T)'),
            ('14BA02', 'Nước ép Cam Vinamilk 1L'),
            ('14BT01', 'Nước ép Táo Vinamilk 1L'),
            ('14AL20', 'Nước ép Lựu Táo Collagen VNM 275ml'),
            ('14AK20', 'Nước ép Kiwi Táo Collagen VNM 275ml'),
            ('14AA20', 'Nước ép Cam Đào Collagen VNM 275ml'),
            ('14BN01', 'Nước ép Nho Vinamilk 1L'),
            ('14BD02', 'Necta Đào Vinamilk 1L'),
            ('05DY10', 'Sữa yến mạch Vinamilk hộp 180ml (24H/T)'),
            ('05BG13', 'SĐN Vinamilk hộp 1L'),
        ]),
    ]

    return all_fm_stores, store_matrix, categories


# ==========================================
# SHEET BUILDER: FM - SKU TRACKER
# ==========================================
def build_fm_sku_sheet(wb, config):
    ws = wb.create_sheet("Chi tiết SKU FamilyMart")
    time_pct = config.get('time_percentage', 65)

    all_fm_stores, store_matrix, categories = get_fm_sku_matrix()
    if not store_matrix:
        return

    # 1. Title Banner (Row 1)
    ws.merge_cells('A1:L1')
    ws['A1'] = "🛒 BÁO CÁO CHI TIẾT NHẬP HÀNG THEO CỬA HÀNG & SKU - CHUỖI FAMILYMART"
    style_header_cell(ws['A1'], bg='1F4E78', size=13, bold=True)
    ws.row_dimensions[1].height = 32

    # 2. KPI Summary Cards (Row 2 & 3)
    tot_tgt = sum(m['store']['target'] for m in store_matrix)
    tot_act = sum(m['actual'] for m in store_matrix)
    pct_total = (tot_act / tot_tgt * 100) if tot_tgt > 0 else 0
    passed_count = sum(1 for m in store_matrix if (m['actual'] / m['store']['target'] * 100 if m['store']['target'] > 0 else 0) >= time_pct)

    if pct_total >= time_pct:
        kpi_eval_text = f"🟢 VƯỢT ({pct_total - time_pct:+.1f}%)"
        kpi_eval_bg = COLOR_GREEN
        kpi_eval_fg = '000000'
        pct_bg = COLOR_GREEN
        pct_fg = '000000'
    elif pct_total >= time_pct * 0.75:
        kpi_eval_text = f"🟠 CẬN ({pct_total - time_pct:+.1f}%)"
        kpi_eval_bg = COLOR_ORANGE
        kpi_eval_fg = 'FFFFFF'
        pct_bg = COLOR_ORANGE
        pct_fg = 'FFFFFF'
    else:
        kpi_eval_text = f"🔴 CHẬM ({pct_total - time_pct:+.1f}%)"
        kpi_eval_bg = COLOR_RED
        kpi_eval_fg = 'FFFFFF'
        pct_bg = COLOR_RED
        pct_fg = 'FFFFFF'

    kpis = [
        ("Số Cửa Hàng", f"{len(store_matrix)} CH", '1F4E78', 'FFFFFF'),
        ("Tổng Chỉ Tiêu", f"{tot_tgt:,.0f} đ", '2E75B6', 'FFFFFF'),
        ("Tổng Thực Hiện", f"{tot_act:,.0f} đ", '70AD47', 'FFFFFF'),
        ("% Đạt", f"{pct_total:.1f}%", pct_bg, pct_fg),
        ("% Timegone", f"{time_pct:.1f}%", 'ED7D31', 'FFFFFF'),
        ("Đánh Giá", kpi_eval_text, kpi_eval_bg, kpi_eval_fg),
        ("CH Đạt Tiến Độ", f"{passed_count}/{len(store_matrix)} CH ({passed_count/len(store_matrix)*100:.0f}%)", '4472C4', 'FFFFFF'),
    ]

    col_ptr = 1
    for label, val_str, bg, fg in kpis:
        c1 = ws.cell(row=2, column=col_ptr, value=label)
        c1.font = Font(bold=True, size=9, color='555555')
        c1.fill = PatternFill('solid', fgColor='F2F2F2')
        c1.alignment = Alignment(horizontal='center', vertical='center')

        c2 = ws.cell(row=3, column=col_ptr, value=val_str)
        c2.font = Font(bold=True, size=11, color=fg)
        c2.fill = PatternFill('solid', fgColor=bg)
        c2.alignment = Alignment(horizontal='center', vertical='center')
        col_ptr += 1

    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 24

    # 3. Base Info Columns (Headers Row 4 & 5)
    # Removed 4 columns: Số Điện Thoại Của PG, Tình Trạng Làm Việc, Mã PG Leader, Họ và Tên PG Leader
    base_headers = [
        ("Địa chỉ Store (ship to)", 42),
        ("Tên Store (Location)", 16),
        ("Mã PG", 13),
        ("Họ Tên PG", 22),
        ("VTCV", 8),
        ("Chỉ tiêu Doanh Số Nhập (VND)", 18),
        ("Tổng Doanh Số Nhập Thực Hiện (VND)", 20),
        ("% Đạt", 11),
    ]

    # Row 4: Group Header for Base Info
    ws.merge_cells('A4:E4')
    ws['A4'] = "THÔNG TIN ĐIỂM BÁN (STORE MASTER)"
    style_header_cell(ws['A4'], bg='FFF2CC', color='1F4E78', size=11, bold=True)

    ws.merge_cells('F4:H4')
    ws['F4'] = "CHỈ TIÊU & TỔNG THỰC HIỆN"
    style_header_cell(ws['F4'], bg='FCE4D6', color='C00000', size=11, bold=True)

    # Row 5: Subheaders for Base Info
    for idx, (h_name, width) in enumerate(base_headers, 1):
        cell = ws.cell(row=5, column=idx, value=h_name)
        if idx <= 5:
            style_header_cell(cell, bg='FFE699', color='000000', size=9, bold=True)
        elif idx == 6:
            style_header_cell(cell, bg='F8CBAD', color='000000', size=9, bold=True)
        elif idx == 7:
            style_header_cell(cell, bg='F4B084', color='000000', size=9, bold=True)
        else:
            style_header_cell(cell, bg='C00000', color='FFFFFF', size=9, bold=True)
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = width

    # SKU Group Headers (Row 4 & 5) starting from Column 9 (Col I)
    cur_col = 9
    sku_col_map = {}  # icode -> (col_idx, cat_cell_bg)

    for cat_name, cat_header_bg, cat_cell_bg, sku_list in categories:
        start_c = cur_col
        end_c = cur_col + len(sku_list) - 1
        start_letter = get_column_letter(start_c)
        end_letter = get_column_letter(end_c)

        # Merge group header in Row 4
        ws.merge_cells(f'{start_letter}4:{end_letter}4')
        group_cell = ws.cell(row=4, column=start_c, value=cat_name)
        style_header_cell(group_cell, bg=cat_header_bg, color='000000', size=10, bold=True)

        # Subheaders in Row 5
        for s_idx, (icode, iname) in enumerate(sku_list):
            c_idx = start_c + s_idx
            sku_col_map[icode] = (c_idx, cat_cell_bg)
            sub_cell = ws.cell(row=5, column=c_idx, value=f"{iname}\n({icode})")
            style_header_cell(sub_cell, bg=cat_header_bg, color='1F4E78', size=8, bold=True)
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = 16

        cur_col = end_c + 1

    ws.row_dimensions[4].height = 24
    ws.row_dimensions[5].height = 45

    # 4. Fill Data Rows (Row 6 onwards)
    start_data_row = 6
    last_col_letter = get_column_letter(cur_col - 1)

    for r_offset, m in enumerate(store_matrix):
        r = start_data_row + r_offset
        st = m['store']
        orders = m['orders']

        ws.cell(row=r, column=1, value=st['addr']).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(row=r, column=2, value=st['location']).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=r, column=3, value=st['ma_pg']).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=r, column=4, value=st['ten_pg']).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(row=r, column=5, value=st['vtcv']).alignment = Alignment(horizontal='center', vertical='center')

        # Target (Col F)
        tgt_cell = ws.cell(row=r, column=6, value=st['target'])
        style_money(tgt_cell)
        tgt_cell.font = Font(bold=True)

        # Actual Formula (Col G) = SUM(I{r}:{last_col_letter}{r})
        act_cell = ws.cell(row=r, column=7, value=f"=SUM(I{r}:{last_col_letter}{r})")
        style_money(act_cell)
        act_cell.font = Font(bold=True)

        # % Dat Formula (Col H) = IF(F{r}>0, G{r}/F{r}, 0)
        pct_cell = ws.cell(row=r, column=8, value=f"=IF(F{r}>0, G{r}/F{r}, 0)")
        pct_cell.number_format = '0.0%'
        pct_cell.alignment = Alignment(horizontal='center', vertical='center')
        pct_cell.font = Font(bold=True)

        # Base info fill
        for c in range(1, 6):
            ws.cell(row=r, column=c).fill = PatternFill('solid', fgColor='FAFAFA' if r % 2 == 0 else 'FFFFFF')
            ws.cell(row=r, column=c).font = Font(size=9)
        ws.cell(row=r, column=6).fill = PatternFill('solid', fgColor='FFF9E6')
        ws.cell(row=r, column=7).fill = PatternFill('solid', fgColor='F2F9F2')

        # SKU columns fill
        for icode, (col_idx, cat_cell_bg) in sku_col_map.items():
            sku_amt = orders.get(icode, 0)
            sku_cell = ws.cell(row=r, column=col_idx, value=sku_amt if sku_amt > 0 else 0)
            style_money(sku_cell)
            sku_cell.font = Font(size=9, color='000000' if sku_amt > 0 else 'A0A0A0')
            if sku_amt > 0:
                sku_cell.fill = PatternFill('solid', fgColor=cat_cell_bg)
            else:
                sku_cell.fill = PatternFill('solid', fgColor='FAFAFA' if r % 2 == 0 else 'FFFFFF')

        ws.row_dimensions[r].height = 20

    end_data_row = start_data_row + len(store_matrix) - 1

    # 5. Summary Row (TỔNG CỘNG)
    tot_row = end_data_row + 1
    ws.merge_cells(f'A{tot_row}:E{tot_row}')
    tot_title = ws.cell(row=tot_row, column=1, value=f"TỔNG CỘNG ({len(store_matrix)} CỬA HÀNG FAMILYMART)")
    tot_title.font = Font(bold=True, color='FFFFFF', size=11)
    tot_title.fill = PatternFill('solid', fgColor='1F4E78')
    tot_title.alignment = Alignment(horizontal='center', vertical='center')

    for c in range(2, 6):
        ws.cell(row=tot_row, column=c).fill = PatternFill('solid', fgColor='1F4E78')

    # Target SUBTOTAL
    tot_tgt_cell = ws.cell(row=tot_row, column=6, value=f"=SUBTOTAL(9, F{start_data_row}:F{end_data_row})")
    style_money(tot_tgt_cell)
    tot_tgt_cell.font = Font(bold=True, color='FFFFFF', size=11)
    tot_tgt_cell.fill = PatternFill('solid', fgColor='1F4E78')

    # Actual SUBTOTAL
    tot_act_cell = ws.cell(row=tot_row, column=7, value=f"=SUBTOTAL(9, G{start_data_row}:G{end_data_row})")
    style_money(tot_act_cell)
    tot_act_cell.font = Font(bold=True, color='FFFF00', size=11)
    tot_act_cell.fill = PatternFill('solid', fgColor='1F4E78')

    # % Dat Formula
    tot_pct_cell = ws.cell(row=tot_row, column=8, value=f"=IF(F{tot_row}>0, G{tot_row}/F{tot_row}, 0)")
    tot_pct_cell.number_format = '0.0%'
    tot_pct_cell.font = Font(bold=True, color='FFFF00', size=12)
    tot_pct_cell.alignment = Alignment(horizontal='center', vertical='center')
    tot_pct_cell.fill = PatternFill('solid', fgColor='1F4E78')

    # SKU Column SUBTOTALs
    for icode, (col_idx, _) in sku_col_map.items():
        col_let = get_column_letter(col_idx)
        sum_c = ws.cell(row=tot_row, column=col_idx, value=f"=SUBTOTAL(9, {col_let}{start_data_row}:{col_let}{end_data_row})")
        style_money(sum_c)
        sum_c.font = Font(bold=True, color='FFFFFF', size=9)
        sum_c.fill = PatternFill('solid', fgColor='1F4E78')

    ws.row_dimensions[tot_row].height = 26

    # 6. Borders, AutoFilter & Conditional Formatting
    add_borders(ws, f'A4:{last_col_letter}{tot_row}')

    # AutoFilter on row 5 across all columns (including Họ Tên PG at Col D for easy filtering)
    ws.auto_filter.ref = f"A5:{last_col_letter}{end_data_row}"

    # Conditional format % Dat (Col H)
    apply_timegone_conditional_formatting(ws, f'H{start_data_row}:H{end_data_row}', time_pct)

    # Freeze panes at column I (scroll horizontally while keeping store master info & targets visible)
    ws.freeze_panes = 'I6'


# ==========================================
# EXCEL EXPORT
# ==========================================
def create_team_report(config, employees, cvs_data, output_file):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheet_names = ["Dashboard", "Tiến độ Team", "Chi tiết CVS", "Data BHX", "Config"]
    for name in sheet_names:
        wb.create_sheet(name)

    time_pct = config.get('time_percentage', 65)

    # ============================================
    # SHEET 1: DASHBOARD
    # ============================================
    ws = wb["Dashboard"]

    ws.merge_cells('B2:H2')
    ws['B2'] = f"📊 THEO DÕI DOANH SỐ TEAM - {config['team_lead'].upper()}"
    ws['B2'].font = Font(size=16, bold=True, color='FFFFFF')
    ws['B2'].fill = PatternFill('solid', fgColor='1F4E78')
    ws['B2'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 35

    ws['B3'] = "Tháng:"
    ws['C3'] = f"{config['month']}/{config['year']}"
    ws['D3'] = "Team Lead:"
    ws['E3'] = config['team_lead']
    ws['F3'] = "% Timegone:"
    ws['G3'] = time_pct / 100
    ws['G3'].number_format = '0.0%'
    ws['H3'] = f"Cập nhật: {datetime.now().strftime('%d/%m/%Y')}"

    for cell in ['B3', 'D3', 'F3']:
        ws[cell].font = Font(bold=True)
        ws[cell].fill = PatternFill('solid', fgColor='FFF2CC')
    for cell in ['C3', 'E3', 'G3', 'H3']:
        ws[cell].fill = PatternFill('solid', fgColor='FFF9E6')
        ws[cell].font = Font(bold=True, color='C00000')
        ws[cell].alignment = Alignment(horizontal='center')

    # Totals
    total_target = sum(e['target'] for e in employees)
    total_bhx = sum(calc_bhx_allocation(config['total_actual_bhx'],
                                         config['total_stores_bhx'],
                                         e['bhx_stores']) for e in employees)
    total_cvs = sum(sum(s['actual'] for s in cvs_data.get(e['name'], []))
                    for e in employees)
    total_actual = total_bhx + total_cvs
    team_pct = (total_actual / total_target * 100) if total_target > 0 else 0

    # KPI section
    ws.merge_cells('B5:H5')
    ws.cell(row=5, column=2, value="🎯 TỔNG QUAN TEAM").font = Font(bold=True, size=13, color='1F4E78')
    ws.cell(row=5, column=2).alignment = Alignment(horizontal='center')

    kpis = [
        ("Target Team", total_target, '#,##0', '4472C4'),
        ("Thực Hiện BHX", total_bhx, '#,##0', '70AD47'),
        ("Thực Hiện CVS", total_cvs, '#,##0', 'ED7D31'),
        ("Tổng Thực Hiện", total_actual, '#,##0', '2E75B6'),
        ("Thiếu", total_target - total_actual, '#,##0;[Red](#,##0)', COLOR_RED),
        ("% Đạt Team", team_pct / 100, '0.0%', COLOR_GREEN if team_pct >= time_pct else (COLOR_ORANGE if team_pct >= time_pct * 0.75 else COLOR_RED)),
    ]

    for idx, (label, value, fmt, color) in enumerate(kpis):
        col = 2 + idx
        lbl_cell = ws.cell(row=6, column=col, value=label)
        lbl_cell.font = Font(bold=True, color='000000' if color == COLOR_GREEN else 'FFFFFF', size=10)
        lbl_cell.fill = PatternFill('solid', fgColor=color)
        lbl_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        val_cell = ws.cell(row=7, column=col, value=value)
        val_cell.number_format = fmt
        val_cell.font = Font(bold=True, size=12)
        val_cell.alignment = Alignment(horizontal='center', vertical='center')
        val_cell.fill = PatternFill('solid', fgColor='F2F2F2')

    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 28

    # Status bar
    ws.merge_cells('B9:H9')
    diff = team_pct - time_pct
    if team_pct >= time_pct:
        status_icon = "🟢"
        status_text = "VƯỢT tiến độ"
        status_bg = COLOR_GREEN
        status_fg = '000000'
    elif team_pct >= time_pct * 0.75:
        status_icon = "🟠"
        status_text = "CẬN tiến độ (CẢNH BÁO)"
        status_bg = COLOR_ORANGE
        status_fg = 'FFFFFF'
    else:
        status_icon = "🔴"
        status_text = "CHẬM tiến độ"
        status_bg = COLOR_RED
        status_fg = 'FFFFFF'
    ws['B9'] = f"{status_icon} Team đang {status_text}: % Đạt ({team_pct:.1f}%) vs % Timegone ({time_pct:.1f}%) → chênh lệch {diff:+.1f}%"
    ws['B9'].font = Font(bold=True, size=11, color=status_fg)
    ws['B9'].fill = PatternFill('solid', fgColor=status_bg)
    ws['B9'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[9].height = 25

    # Ranking table
    ws.merge_cells('B11:H11')
    ws['B11'] = "XẾP HẠNG NHÂN VIÊN"
    ws['B11'].font = Font(bold=True, size=13, color='1F4E78')
    ws['B11'].alignment = Alignment(horizontal='center')

    rank_headers = ["Hạng", "Nhân Viên", "Target", "BHX", "CVS", "Tổng TH", "% Đạt"]
    for idx, h in enumerate(rank_headers, 2):
        cell = ws.cell(row=12, column=idx, value=h)
        style_header_cell(cell, bg='4472C4', size=10)

    # Sort by percent DESC
    ranked = sorted([(emp, calc_employee_summary(emp, config, cvs_data))
                     for emp in employees],
                    key=lambda x: x[1]['percent'], reverse=True)

    for idx, (emp, calc) in enumerate(ranked, 1):
        r = 12 + idx
        rank_cell = ws.cell(row=r, column=2, value=idx)
        rank_cell.alignment = Alignment(horizontal='center')
        rank_cell.font = Font(bold=True)

        ws.cell(row=r, column=3, value=emp['name']).font = Font(bold=True)
        ws.cell(row=r, column=4, value=emp['target']).number_format = '#,##0'
        ws.cell(row=r, column=5, value=calc['bhx_actual']).number_format = '#,##0'
        ws.cell(row=r, column=6, value=calc['cvs_total']).number_format = '#,##0'
        ws.cell(row=r, column=7, value=calc['total_actual']).number_format = '#,##0'
        pct_cell = ws.cell(row=r, column=8, value=calc['percent'] / 100)
        pct_cell.number_format = '0.0%'
        pct_cell.font = Font(bold=True)
        pct_cell.alignment = Alignment(horizontal='center')

    last_rank_row = 12 + len(ranked)
    apply_timegone_conditional_formatting(ws, f'H13:H{last_rank_row}', time_pct)
    add_borders(ws, f'B12:H{last_rank_row}')

    for col_letter, width in [('A', 3), ('B', 10), ('C', 22), ('D', 18),
                               ('E', 18), ('F', 18), ('G', 18), ('H', 12)]:
        ws.column_dimensions[col_letter].width = width

    # ============================================
    # SHEET 2: TIẾN ĐỘ TEAM
    # ============================================
    ws2 = wb["Tiến độ Team"]

    ws2.merge_cells('B2:M2')
    ws2['B2'] = f"THEO DÕI DOANH SỐ - TEAM {config['team_lead'].upper()} - THÁNG {config['month']}/{config['year']}"
    style_header_cell(ws2['B2'], bg='1F4E78', size=14)
    ws2.row_dimensions[2].height = 30

    ws2['B3'] = "% Timegone:"
    ws2['C3'] = time_pct / 100
    ws2['C3'].number_format = '0.0%'
    ws2['C3'].font = Font(bold=True, size=12, color='C00000')
    ws2['C3'].alignment = Alignment(horizontal='center')
    ws2['B3'].font = Font(bold=True)
    ws2['B3'].fill = PatternFill('solid', fgColor='FFF2CC')
    ws2['C3'].fill = PatternFill('solid', fgColor='FFF9E6')

    ws2['E3'] = "Tổng CH BHX:"
    ws2['F3'] = config['total_stores_bhx']
    ws2['F3'].font = Font(bold=True, color='C00000')
    ws2['F3'].alignment = Alignment(horizontal='center')
    ws2['E3'].fill = PatternFill('solid', fgColor='FFF2CC')
    ws2['F3'].fill = PatternFill('solid', fgColor='FFF9E6')
    ws2['E3'].font = Font(bold=True)

    ws2['H3'] = "TB BHX/CH:"
    ws2['I3'] = round(config['total_actual_bhx'] / config['total_stores_bhx'])
    ws2['I3'].number_format = '#,##0'
    ws2['I3'].font = Font(bold=True, color='C00000')
    ws2['I3'].alignment = Alignment(horizontal='center')
    ws2['H3'].fill = PatternFill('solid', fgColor='FFF2CC')
    ws2['I3'].fill = PatternFill('solid', fgColor='FFF9E6')
    ws2['H3'].font = Font(bold=True)

    ws2.merge_cells('E5:K5')
    ws2['E5'] = "THỰC HIỆN (chi tiết theo kênh)"
    style_header_cell(ws2['E5'], bg='2E75B6', size=11)

    ws2.merge_cells('L5:M5')
    ws2['L5'] = "ĐÁNH GIÁ"
    style_header_cell(ws2['L5'], bg=COLOR_RED, size=11)

    headers = [
        ('B', 'STT'), ('C', 'Nhân Viên'), ('D', 'Target'),
        ('E', 'BHX\n(cố định)'), ('F', 'GS25'), ('G', '7-Eleven'),
        ('H', 'FamilyMart'), ('I', 'Circle K'), ('J', 'Hoàng Đức'),
        ('K', 'WinMart+'),
        ('L', 'Tổng TH\n(BHX+CVS)'), ('M', '% Đạt'),
    ]
    for col, header in headers:
        cell = ws2[f'{col}6']
        cell.value = header
        style_header_cell(cell, bg='4472C4', size=10)

    ws2.row_dimensions[5].height = 22
    ws2.row_dimensions[6].height = 40

    current_row = 7
    for idx, emp in enumerate(employees, 1):
        calc = calc_employee_summary(emp, config, cvs_data)

        ws2.cell(row=current_row, column=2, value=idx).alignment = Alignment(horizontal='center')
        ws2.cell(row=current_row, column=3, value=emp['name']).font = Font(bold=True)
        ws2.cell(row=current_row, column=4, value=emp['target'])
        ws2.cell(row=current_row, column=5, value=calc['bhx_actual'])
        ws2.cell(row=current_row, column=6, value=calc['chains']['GS25'])
        ws2.cell(row=current_row, column=7, value=calc['chains']['7-Eleven'])
        ws2.cell(row=current_row, column=8, value=calc['chains']['FamilyMart'])
        ws2.cell(row=current_row, column=9, value=calc['chains']['Circle K'])
        ws2.cell(row=current_row, column=10, value=calc['chains']['Hoàng Đức'])
        ws2.cell(row=current_row, column=11, value=calc['chains']['WinMart+'])
        ws2.cell(row=current_row, column=12, value=calc['total_actual']).font = Font(bold=True)

        pct_cell = ws2.cell(row=current_row, column=13, value=calc['percent'] / 100)
        pct_cell.number_format = '0.0%'
        pct_cell.font = Font(bold=True, size=11)
        pct_cell.alignment = Alignment(horizontal='center')

        for col in range(4, 13):
            style_money(ws2.cell(row=current_row, column=col))

        current_row += 1

    # TOTAL ROW
    total_row = current_row
    ws2.cell(row=total_row, column=3, value="TỔNG TEAM").font = Font(bold=True, color='FFFFFF', size=12)
    ws2.cell(row=total_row, column=3).fill = PatternFill('solid', fgColor='1F4E78')
    ws2.cell(row=total_row, column=3).alignment = Alignment(horizontal='center')
    ws2.cell(row=total_row, column=2).fill = PatternFill('solid', fgColor='1F4E78')

    ws2.cell(row=total_row, column=4, value=f"=SUM(D7:D{total_row-1})")
    ws2.cell(row=total_row, column=5, value=f"=SUM(E7:E{total_row-1})")
    ws2.cell(row=total_row, column=6, value=f"=SUM(F7:F{total_row-1})")
    ws2.cell(row=total_row, column=7, value=f"=SUM(G7:G{total_row-1})")
    ws2.cell(row=total_row, column=8, value=f"=SUM(H7:H{total_row-1})")
    ws2.cell(row=total_row, column=9, value=f"=SUM(I7:I{total_row-1})")
    ws2.cell(row=total_row, column=10, value=f"=SUM(J7:J{total_row-1})")
    ws2.cell(row=total_row, column=11, value=f"=SUM(K7:K{total_row-1})")
    ws2.cell(row=total_row, column=12, value=f"=SUM(L7:L{total_row-1})")
    ws2.cell(row=total_row, column=13, value=f"=L{total_row}/D{total_row}")
    ws2.cell(row=total_row, column=13).number_format = '0.0%'

    for col in range(4, 13):
        c = ws2.cell(row=total_row, column=col)
        c.number_format = '#,##0'
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F4E78')
        c.alignment = Alignment(horizontal='right')
    ws2.cell(row=total_row, column=13).font = Font(bold=True, color='FFFF00', size=12)
    ws2.cell(row=total_row, column=13).fill = PatternFill('solid', fgColor='1F4E78')
    ws2.cell(row=total_row, column=13).alignment = Alignment(horizontal='center')

    widths = {'A': 2, 'B': 5, 'C': 22, 'D': 18, 'E': 16, 'F': 14,
              'G': 14, 'H': 14, 'I': 14, 'J': 14, 'K': 14, 'L': 18, 'M': 10}
    for col, w in widths.items():
        ws2.column_dimensions[col].width = w

    apply_timegone_conditional_formatting(ws2, f'M7:M{total_row-1}', time_pct)

    add_borders(ws2, f'B6:M{total_row}')
    ws2.freeze_panes = 'D7'

    legend_row = total_row + 2
    warn_pct = time_pct * 0.75
    ws2.merge_cells(f'B{legend_row}:M{legend_row}')
    ws2.cell(row=legend_row, column=2,
             value=f"📌 % Đạt ≥ % Timegone ({time_pct:.1f}%) → Vượt/Đạt tiến độ 🟢 (#28E11F) | Cận Timegone ({warn_pct:.1f}% - <{time_pct:.1f}%) → Cảnh báo 🟠 (#F5991F) | Thấp hơn nhiều (<{warn_pct:.1f}%) → Cần đẩy nhanh 🔴 (#DD2323)")
    ws2.cell(row=legend_row, column=2).font = Font(italic=True, size=10, color='555555')
    ws2.cell(row=legend_row, column=2).alignment = Alignment(horizontal='center')

    formula_row = legend_row + 1
    ws2.merge_cells(f'B{formula_row}:M{formula_row}')
    ws2.cell(row=formula_row, column=2,
             value="🧮 BHX cố định = (Tổng BHX ÷ 232) × Số CH phụ trách | Thực Hiện = BHX + GS25 + 7E + FM + CK + HĐ + WMP | % Đạt = TH ÷ Target")
    ws2.cell(row=formula_row, column=2).font = Font(italic=True, size=9, color='555555')
    ws2.cell(row=formula_row, column=2).alignment = Alignment(horizontal='center')

    # ============================================
    # SHEET 3: CHI TIẾT CVS (từng cửa hàng)
    # ============================================
    ws3 = wb["Chi tiết CVS"]
    ws3.merge_cells('A1:F1')
    ws3['A1'] = "CHI TIẾT DOANH SỐ CVS - TỪNG CỬA HÀNG"
    style_header_cell(ws3['A1'], bg='1F4E78', size=14)
    ws3.row_dimensions[1].height = 30

    cvs_headers = ["STT", "Nhân Viên", "Chuỗi", "Mã CH", "Địa Chỉ", "Thực Hiện"]
    for idx, h in enumerate(cvs_headers, 1):
        cell = ws3.cell(row=3, column=idx, value=h)
        style_header_cell(cell)
    ws3.row_dimensions[3].height = 25

    row_idx = 4
    stt = 1

    for emp in employees:
        emp_stores = cvs_data.get(emp['name'], [])
        if not emp_stores:
            continue

        ws3.merge_cells(f'A{row_idx}:D{row_idx}')
        emp_total = sum(s['actual'] for s in emp_stores)
        ws3.cell(row=row_idx, column=1, value=f"👤 {emp['name']} ({len(emp_stores)} cửa hàng CVS)")
        ws3.cell(row=row_idx, column=1).font = Font(bold=True, color='FFFFFF', size=11)
        ws3.cell(row=row_idx, column=1).fill = PatternFill('solid', fgColor='4472C4')
        ws3.cell(row=row_idx, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=1)

        ws3.cell(row=row_idx, column=5, value="TỔNG CVS:").font = Font(bold=True, color='FFFFFF')
        ws3.cell(row=row_idx, column=5).fill = PatternFill('solid', fgColor='4472C4')
        ws3.cell(row=row_idx, column=5).alignment = Alignment(horizontal='right')

        ws3.cell(row=row_idx, column=6, value=emp_total).number_format = '#,##0'
        ws3.cell(row=row_idx, column=6).font = Font(bold=True, color='FFFF00', size=11)
        ws3.cell(row=row_idx, column=6).fill = PatternFill('solid', fgColor='4472C4')
        ws3.cell(row=row_idx, column=6).alignment = Alignment(horizontal='right')
        row_idx += 1

        for store in emp_stores:
            chain_color = CHAIN_COLORS.get(store['chain'], '70AD47' if 'Hoàng Đức' in store['chain'] else '888888')
            ws3.cell(row=row_idx, column=1, value=stt).alignment = Alignment(horizontal='center')
            ws3.cell(row=row_idx, column=2, value=emp['name'])
            chain_cell = ws3.cell(row=row_idx, column=3, value=store['chain'])
            chain_cell.font = Font(bold=True, color='FFFFFF', size=10)
            chain_cell.fill = PatternFill('solid', fgColor=chain_color)
            chain_cell.alignment = Alignment(horizontal='center')
            ws3.cell(row=row_idx, column=4, value=store['store_code']).font = Font(size=10)
            ws3.cell(row=row_idx, column=5, value=store['store_address']).font = Font(size=10)
            ws3.cell(row=row_idx, column=6, value=store['actual']).number_format = '#,##0'
            ws3.cell(row=row_idx, column=6).alignment = Alignment(horizontal='right')
            row_idx += 1
            stt += 1

        row_idx += 1

    # Grand total
    ws3.merge_cells(f'A{row_idx}:E{row_idx}')
    ws3.cell(row=row_idx, column=1, value="🎯 TỔNG CVS TOÀN TEAM").font = Font(bold=True, color='FFFFFF', size=12)
    ws3.cell(row=row_idx, column=1).fill = PatternFill('solid', fgColor='1F4E78')
    ws3.cell(row=row_idx, column=1).alignment = Alignment(horizontal='center', vertical='center')

    grand_total = sum(sum(s['actual'] for s in cvs_data.get(e['name'], [])) for e in employees)
    ws3.cell(row=row_idx, column=6, value=grand_total).number_format = '#,##0'
    ws3.cell(row=row_idx, column=6).font = Font(bold=True, color='FFFF00', size=12)
    ws3.cell(row=row_idx, column=6).fill = PatternFill('solid', fgColor='1F4E78')
    ws3.cell(row=row_idx, column=6).alignment = Alignment(horizontal='right')

    ws3.column_dimensions['A'].width = 6
    ws3.column_dimensions['B'].width = 20
    ws3.column_dimensions['C'].width = 13
    ws3.column_dimensions['D'].width = 15
    ws3.column_dimensions['E'].width = 40
    ws3.column_dimensions['F'].width = 18

    ws3.freeze_panes = 'A4'

    # ============================================
    # SHEET 4: DATA BHX
    # ============================================
    ws4 = wb["Data BHX"]
    ws4.merge_cells('A1:F1')
    ws4['A1'] = f"PHÂN BỔ DOANH SỐ BHX ({config['total_stores_bhx']} CỬA HÀNG)"
    style_header_cell(ws4['A1'], bg='1F4E78', size=13)
    ws4.row_dimensions[1].height = 28

    ws4['A3'] = "Tổng Thực Hiện BHX:"
    ws4['B3'] = config['total_actual_bhx']
    ws4['B3'].number_format = '#,##0'
    ws4['B3'].font = Font(bold=True, color='C00000')

    ws4['A4'] = "Tổng số cửa hàng:"
    ws4['B4'] = config['total_stores_bhx']
    ws4['B4'].font = Font(bold=True)

    ws4['A5'] = "TB doanh số/CH:"
    ws4['B5'] = "=B3/B4"
    ws4['B5'].number_format = '#,##0'
    ws4['B5'].font = Font(bold=True, color='C00000')

    for cell in ['A3', 'A4', 'A5']:
        ws4[cell].font = Font(bold=True)
        ws4[cell].fill = PatternFill('solid', fgColor='FFF2CC')

    bhx_headers = ["STT", "Nhân Viên", "Số CH BHX", "TB/CH", "Doanh Số BHX", "% BHX/Team"]
    for idx, h in enumerate(bhx_headers, 1):
        cell = ws4.cell(row=7, column=idx, value=h)
        style_header_cell(cell)
    ws4.row_dimensions[7].height = 25

    total_bhx_stores = sum(e['bhx_stores'] for e in employees)
    for idx, emp in enumerate(employees, 1):
        r = 7 + idx
        bhx_val = calc_bhx_allocation(config['total_actual_bhx'],
                                       config['total_stores_bhx'], emp['bhx_stores'])
        ws4.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal='center')
        ws4.cell(row=r, column=2, value=emp['name']).font = Font(bold=True)
        ws4.cell(row=r, column=3, value=emp['bhx_stores']).alignment = Alignment(horizontal='center')
        ws4.cell(row=r, column=4, value="=$B$5")
        ws4.cell(row=r, column=4).number_format = '#,##0'
        ws4.cell(row=r, column=4).alignment = Alignment(horizontal='right')
        ws4.cell(row=r, column=5, value=bhx_val).number_format = '#,##0'
        ws4.cell(row=r, column=5).alignment = Alignment(horizontal='right')
        ws4.cell(row=r, column=6, value=emp['bhx_stores'] / config['total_stores_bhx'])
        ws4.cell(row=r, column=6).number_format = '0.00%'
        ws4.cell(row=r, column=6).alignment = Alignment(horizontal='center')

    last_bhx_row = 7 + len(employees)

    r = last_bhx_row + 1
    ws4.cell(row=r, column=1, value="").fill = PatternFill('solid', fgColor='1F4E78')
    ws4.cell(row=r, column=2, value="TỔNG").font = Font(bold=True, color='FFFFFF')
    ws4.cell(row=r, column=2).fill = PatternFill('solid', fgColor='1F4E78')
    ws4.cell(row=r, column=3, value=total_bhx_stores).font = Font(bold=True, color='FFFFFF')
    ws4.cell(row=r, column=3).fill = PatternFill('solid', fgColor='1F4E78')
    ws4.cell(row=r, column=3).alignment = Alignment(horizontal='center')
    ws4.cell(row=r, column=5, value=f"=SUM(E8:E{last_bhx_row})")
    ws4.cell(row=r, column=5).number_format = '#,##0'
    ws4.cell(row=r, column=5).font = Font(bold=True, color='FFFF00')
    ws4.cell(row=r, column=5).fill = PatternFill('solid', fgColor='1F4E78')
    ws4.cell(row=r, column=5).alignment = Alignment(horizontal='right')
    ws4.cell(row=r, column=6, value=total_bhx_stores / config['total_stores_bhx'])
    ws4.cell(row=r, column=6).number_format = '0.00%'
    ws4.cell(row=r, column=6).font = Font(bold=True, color='FFFF00')
    ws4.cell(row=r, column=6).fill = PatternFill('solid', fgColor='1F4E78')
    ws4.cell(row=r, column=6).alignment = Alignment(horizontal='center')
    ws4.cell(row=r, column=4).fill = PatternFill('solid', fgColor='1F4E78')

    ws4.column_dimensions['A'].width = 6
    ws4.column_dimensions['B'].width = 22
    ws4.column_dimensions['C'].width = 15
    ws4.column_dimensions['D'].width = 18
    ws4.column_dimensions['E'].width = 20
    ws4.column_dimensions['F'].width = 15

    add_borders(ws4, f'A7:F{r}')

    # ============================================
    # SHEET 5: FM - NGUYỄN ĐỨC HÒA (SKU TRACKER)
    # ============================================
    build_fm_sku_sheet(wb, config)

    # ============================================
    # SHEET 6: CONFIG
    # ============================================
    ws5 = wb["Config"]
    ws5.merge_cells('A1:B1')
    ws5['A1'] = "CẤU HÌNH BÁO CÁO"
    style_header_cell(ws5['A1'], bg='1F4E78', size=13)

    ws5['A3'] = "Thông số"
    ws5['B3'] = "Giá trị"
    style_header_cell(ws5['A3'])
    style_header_cell(ws5['B3'])

    config_items = [
        ("Team Lead", config['team_lead']),
        ("Tháng báo cáo", f"{config['month']}/{config['year']}"),
        ("% Timegone", f"{time_pct:.2f}%"),
        ("Tổng Target BHX (hệ thống)", f"{config['total_target_bhx']:,} VNĐ"),
        ("Tổng Thực Hiện BHX (hệ thống)", f"{config['total_actual_bhx']:,} VNĐ"),
        ("Tổng cửa hàng BHX", config['total_stores_bhx']),
        ("Số NV trong team", len(employees)),
        ("Tổng CH CVS toàn team", sum(len(cvs_data.get(e['name'], [])) for e in employees)),
        ("Cập nhật lúc", datetime.now().strftime('%d/%m/%Y %H:%M')),
    ]
    for idx, (k, v) in enumerate(config_items, 4):
        ws5.cell(row=idx, column=1, value=k).font = Font(bold=True)
        ws5.cell(row=idx, column=2, value=v)

    ws5.column_dimensions['A'].width = 35
    ws5.column_dimensions['B'].width = 30
    add_borders(ws5, f'A3:B{len(config_items)+3}')

    wb.active = 0
    try:
        wb.save(output_file)
    except PermissionError:
        base, ext = os.path.splitext(output_file)
        fallback = f"{base}_moi{ext}"
        print(f"[WARNING] File {output_file} dang mo trong Excel! Dang luu sang: {fallback}")
        wb.save(fallback)
        output_file = fallback
    return output_file


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    def find_file(filename):
        if os.path.exists(filename):
            return filename
        candidate = os.path.join(script_dir, filename)
        if os.path.exists(candidate):
            return candidate
        return filename

    config_file = sys.argv[1] if len(sys.argv) > 1 else find_file('team_config.csv')
    employees_file = sys.argv[2] if len(sys.argv) > 2 else find_file('team_employees.csv')
    cvs_file = sys.argv[3] if len(sys.argv) > 3 else find_file('team_cvs_stores.csv')

    print("=" * 70)
    print("  TEAM LEADER REPORT V4.1 - BHX + CVS + SKU FAMILYMART TRACKER")
    print("=" * 70)

    for f in [config_file, employees_file, cvs_file]:
        if not os.path.exists(f):
            print(f"[ERROR] File not found: {f}")
            sys.exit(1)

    config = read_config(config_file)
    employees = read_team_employees(employees_file)
    cvs_data = read_cvs_stores(cvs_file)

    print(f"Team Lead:    {config['team_lead']}")
    print(f"Thang:        {config['month']}/{config['year']}")
    print(f"% Timegone:   {config.get('time_percentage')}%")
    print(f"So NV:        {len(employees)}")
    print(f"So CH CVS:    {sum(len(cvs_data.get(e['name'], [])) for e in employees)}")
    print("-" * 70)

    print(f"{'Nhan Vien':<20} {'BHX':>15} {'CVS':>15} {'Tong TH':>17} {'% Dat':>8}")
    print("-" * 70)
    total_target = 0
    total_actual = 0
    for emp in employees:
        calc = calc_employee_summary(emp, config, cvs_data)
        total_target += emp['target']
        total_actual += calc['total_actual']
        print(f"{emp['name'][:20]:<20} {calc['bhx_actual']:>15,} {calc['cvs_total']:>15,} "
              f"{calc['total_actual']:>17,} {calc['percent']:>7.1f}%")

    team_pct = total_actual / total_target * 100 if total_target > 0 else 0
    print("-" * 70)
    print(f"{'TONG TEAM':<20} {'':<15} {'':<15} {total_actual:>17,} {team_pct:>7.1f}%")
    print(f"Target Team:  {total_target:,} VND")
    diff = team_pct - config.get('time_percentage', 0)
    status = "VUOT tien do" if diff >= 0 else "CHAM tien do"
    print(f"Status:       {status} ({diff:+.1f}%)")
    print("=" * 70)

    output_filename = config.get('output_file', 'Team_CamGiang_Report.xlsx')
    output = create_team_report(config, employees, cvs_data, output_filename)
    print(f"[DONE] Output saved: {output}")


if __name__ == "__main__":
    main()
