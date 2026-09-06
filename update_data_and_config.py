import sys
import os
import csv
import glob
import pyxlsb
import openpyxl
import re
from collections import defaultdict

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

def normalize_text(t):
    if not t:
        return ''
    t = str(t).lower()
    t = re.sub(r'[,.\s\-_/()]', '', t)
    return t

print("=== BẮT ĐẦU CẬP NHẬT DỮ LIỆU TỰ ĐỘNG TỪ FILE NGUỒN ===")

script_dir = os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------
# 1. TÌM FILE NGUỒN XLSB & XLSX
# ----------------------------------------------------
search_dirs = [
    script_dir,
    os.path.join(script_dir, '..'),
    r'C:\Users\giang\Downloads',
    r'C:\Users\giang\Downloads\DOANH SỐ',
    r'C:\Users\giang\Downloads\team-cvs-bhx-tracker-skill_1'
]

xlsb_candidates = []
for d in search_dirs:
    if os.path.exists(d):
        xlsb_candidates.extend(glob.glob(os.path.join(d, '*HNTRINH*.xlsb')))

xlsb_candidates = list(set([f for f in xlsb_candidates if not os.path.basename(f).startswith('~')]))

st_xlsb = None
cvs_xlsb = None

for f in xlsb_candidates:
    bn = os.path.basename(f)
    if 'HNTRINH_ST' in bn:
        if not st_xlsb or os.path.getmtime(f) > os.path.getmtime(st_xlsb):
            st_xlsb = f
    elif 'HNTRINH_KD6' in bn or 'HNTRINH' in bn:
        if not cvs_xlsb or os.path.getmtime(f) > os.path.getmtime(cvs_xlsb):
            cvs_xlsb = f

# Fallback if only 1 xlsb found
if not st_xlsb and cvs_xlsb:
    st_xlsb = cvs_xlsb
if not cvs_xlsb and st_xlsb:
    cvs_xlsb = st_xlsb

if not cvs_xlsb:
    raise FileNotFoundError("Không tìm thấy file XLSB chứa dữ liệu HNTRINH!")

print(f"File Siêu Thị (BHX): {st_xlsb}")
print(f"File CVS (CK, GS25, 7E, FM): {cvs_xlsb}")

# ----------------------------------------------------
# 2. TÍNH DOANH SỐ BÁCH HÓA XANH (BHX) TỪ 5 HUBS
# ----------------------------------------------------
bhx_hubs_def = [
    'Ấp 4, Xã Châu Pha, Thành phố Hồ Chí Minh, Việt Nam',
    'G243 Bùi Văn Hòa, Khu Phố 7, Phường Long Bình, Tỉnh Đồng Nai, Việt Nam',
    'Hub - Ấp 4, Xã Châu Pha, Thành phố Hồ Chí Minh, Việt Nam',
    'Hub - G243 Bùi Văn Hòa, Khu Phố 7, Phường Long Bình, Tỉnh Đồng Nai, Việt Nam',
    'Hub - Đường Phan Huy Chú, tổ 6, khu phố 1, Phường Phú Bình, Thành phố Long Khánh, Tỉnh Đồng Nai'
]

bhx_hub_amounts = defaultdict(float)
total_actual_bhx = 0
wmp_amounts = {
    'WMP_DNI_KP_TRUNG_TAM_XUAN_LAP': 0,
    'WMP_DNI_285_287_CACH_MANG_THAN': 0,
    'WMP_HCM_90A_92_PHAN_CHU_TRINH': 0,
    'WMP_DNI_A_01_04_TOPAZ_TWINS': 0,
}

with pyxlsb.open_workbook(st_xlsb) as wb:
    sheet_name = 'SO' if 'SO' in wb.sheets else ('PO' if 'PO' in wb.sheets else wb.sheets[0])
    with wb.get_sheet(sheet_name) as s:
        header = None
        for i, r in enumerate(s.rows()):
            vals = [c.v for c in r]
            if i == 0:
                header = [str(h) for h in vals]
                idx_c = header.index('CUST_NO')
                idx_a = header.index('Address')
                idx_m = header.index('SumOfAMOUNT')
                continue
            cust = str(vals[idx_c] or '')
            amt = float(vals[idx_m] or 0)

            # BHX (XH4001)
            if cust == 'XH4001':
                addr = str(vals[idx_a] or '').strip()
                norm_a = normalize_text(addr)
                for h in bhx_hubs_def:
                    if addr == h or norm_a == normalize_text(h):
                        bhx_hub_amounts[h] += amt
                        break

            # WinMart+ (WMP)
            loc_val = str(vals[7] or '').strip() if len(vals) > 7 else ''
            for wk in wmp_amounts:
                if wk.lower() in loc_val.lower() or (wk == 'WMP_DNI_285_287_CACH_MANG_THAN' and '285_287' in loc_val.lower()):
                    wmp_amounts[wk] += amt

print("\n--- CHI TIẾT 5 HUBS BÁCH HÓA XANH (BHX) ---")
for h in bhx_hubs_def:
    amt = bhx_hub_amounts[h]
    total_actual_bhx += amt
    print(f"  + {h}: {amt:15,.0f} VNĐ")

total_stores_bhx = 232
bhx_per_store = total_actual_bhx / total_stores_bhx if total_stores_bhx > 0 else 0
print(f"  => TỔNG 5 HUBS BHX: {total_actual_bhx:15,.0f} VNĐ")
print(f"  => ĐƠN GIÁ BHX / CỬA HÀNG (÷ {total_stores_bhx}): {bhx_per_store:15,.2f} VNĐ (~ {round(bhx_per_store):,d} VNĐ)")

print("\n--- CHI TIẾT 4 CỬA HÀNG WINMART+ (WMP) CỦA NGUYỄN ĐỨC HÒA ---")
for wk, wamt in wmp_amounts.items():
    print(f"  + {wk}: {wamt:15,.0f} VNĐ")
print(f"  => TỔNG 4 WMP: {sum(wmp_amounts.values()):15,.0f} VNĐ")

# ----------------------------------------------------
# 3. TÍNH DOANH SỐ GS25 TỪ 2 KHO DC (* 20.46% ÷ 72)
# ----------------------------------------------------
gs25_dcs_def = [
    'Lô H.04, Đường số 1, Khu Công Nghiệp Long Hậu, Xã Long Hậu, Huyện Cần Giuộc, Tỉnh Long An',
    'Lô IIIB2, Trung tâm thương mại Bình Điền, đường Nguyễn Văn Linh, Khu phố 06, Phường 07, Quận 08, Thành Phố Hồ Chí Minh'
]
gs25_amounts = defaultdict(float)
total_gs25_dcs = 0

# ----------------------------------------------------
# 4. TÍNH DOANH SỐ 7-ELEVEN (7E) TỪ 2 KHO DC (* 3.62% ÷ 5 CỬA HÀNG)
# ----------------------------------------------------
se_dcs_def = [
    'BW Tân Phú Trung - Lô D2, Khu công nghiệp Tân Phú Trung, xã Tân Phú Trung, huyện Củ Chi, Thành phố Hồ Chí Minh',
    'Lô II-3 Nhóm CN II,Đường số 11,KCN Tân Bình,P.Tây Thạnh,Q.Tân Phú,Tp.HCM'
]
se_amounts = defaultdict(float)
total_se_dcs = 0

# ----------------------------------------------------
# 5. TÍNH DOANH SỐ CIRCLE K: KHO KHÔ / 233 + TỪNG CỬA HÀNG
# ----------------------------------------------------
ck_kho_kho_amt = 0
ck_stores_so = defaultdict(float)

# ----------------------------------------------------
# 6. ĐỌC DỮ LIỆU CVS (GS25, 7E, CK, FM) TỪ CVS XLSB
# ----------------------------------------------------
with pyxlsb.open_workbook(cvs_xlsb) as wb:
    sheet_name = 'SO' if 'SO' in wb.sheets else ('PO' if 'PO' in wb.sheets else wb.sheets[0])
    with wb.get_sheet(sheet_name) as s:
        header = None
        for i, r in enumerate(s.rows()):
            vals = [c.v for c in r]
            if i == 0:
                header = [str(h) for h in vals]
                idx_c = header.index('CUST_NO')
                idx_a = header.index('Address')
                idx_m = header.index('SumOfAMOUNT')
                idx_n = header.index('CUST_NAME') if 'CUST_NAME' in header else idx_c
                continue
            cust = str(vals[idx_c] or '')
            cname = str(vals[idx_n] or '')
            addr = str(vals[idx_a] or '').strip()
            amt = float(vals[idx_m] or 0)
            norm_a = normalize_text(addr)

            # GS25 (GS0003)
            if cust == 'GS0003':
                if 'longhau' in norm_a or 'h04' in norm_a:
                    gs25_amounts[gs25_dcs_def[0]] += amt
                elif 'binhdien' in norm_a or 'iiib2' in norm_a:
                    gs25_amounts[gs25_dcs_def[1]] += amt

            # 7-Eleven (SS4001, SS4002, SS5003 hoặc Seven System)
            if cust in ['SS4001', 'SS4002', 'SS5003'] or 'seven' in cname.lower():
                if 'tanphutrung' in norm_a or 'd2' in norm_a:
                    se_amounts[se_dcs_def[0]] += amt
                elif 'tanbinh' in norm_a or 'ii3' in norm_a:
                    se_amounts[se_dcs_def[1]] += amt

            # Circle K (VT4050, VT3013, VT3014, VT3015)
            if cust in ['VT4050', 'VT3013', 'VT3014', 'VT3015']:
                if 'namtanuyen' in norm_a or 'g19' in norm_a:
                    ck_kho_kho_amt += amt
                else:
                    ck_stores_so[addr] += amt

print("\n--- CHI TIẾT 2 KHO DC GS25 ---")
for d in gs25_dcs_def:
    amt = gs25_amounts[d]
    total_gs25_dcs += amt
    print(f"  + {d}: {amt:15,.0f} VNĐ")

# Công thức GS25: Tổng 2 DC * 20.46% / 72
gs25_per_store = (total_gs25_dcs * 0.2046) / 72 if total_gs25_dcs > 0 else 0
print(f"  => TỔNG 2 KHO DC GS25: {total_gs25_dcs:15,.0f} VNĐ")
print(f"  => ĐƠN GIÁ GS25 / CỬA HÀNG ((Tổng × 20.46%) ÷ 72): {gs25_per_store:15,.2f} VNĐ (~ {round(gs25_per_store):,d} VNĐ)")

print("\n--- CHI TIẾT 2 KHO DC 7-ELEVEN (7E) ---")
for d in se_dcs_def:
    amt = se_amounts[d]
    total_se_dcs += amt
    print(f"  + {d}: {amt:15,.0f} VNĐ")

# Công thức 7-Eleven: Tổng 2 DC * 3.62% / 5
total_stores_7e = 5
se_per_store = (total_se_dcs * 0.0362) / total_stores_7e if total_stores_7e > 0 else 0
print(f"  => TỔNG 2 KHO DC 7-ELEVEN: {total_se_dcs:15,.0f} VNĐ")
print(f"  => ĐƠN GIÁ 7-ELEVEN / CỬA HÀNG ((Tổng × 3.62%) ÷ {total_stores_7e}): {se_per_store:15,.2f} VNĐ (~ {round(se_per_store):,d} VNĐ)")

print("\n--- CHI TIẾT CIRCLE K KHO KHÔ & HỆ SỐ T ---")
print(f"  + Doanh số Kho khô Tân Uyên: {ck_kho_kho_amt:15,.0f} VNĐ")
total_stores_ck_system = 233
T_ck = ck_kho_kho_amt / total_stores_ck_system if total_stores_ck_system > 0 else 0
print(f"  => HỆ SỐ T (Kho khô ÷ {total_stores_ck_system}): {T_ck:15,.2f} VNĐ (~ {round(T_ck):,d} VNĐ)")

# ----------------------------------------------------
# 7. PARSE FAMILYMART MATRIX
# ----------------------------------------------------
from team_leader_report_v4 import get_fm_sku_matrix
all_fm_stores_list, store_matrix, categories = get_fm_sku_matrix()
fm_matrix_dict = {}
for m in store_matrix:
    st = m['store']
    key = st['addr'].split(',')[0].strip().lower()
    fm_matrix_dict[key] = m['actual']
    if st.get('code'):
        fm_matrix_dict[st['code']] = m['actual']
    if st.get('store_code'):
        fm_matrix_dict[st['store_code']] = m['actual']

# ----------------------------------------------------
# 8. CẬP NHẬT TEAM_CVS_STORES.CSV
# ----------------------------------------------------
with open('team_cvs_stores.csv', 'r', encoding='utf-8-sig') as f:
    old_stores = list(csv.DictReader(f))

# Danh sách 5 cửa hàng 7E chuẩn từ DOANH SỐ T8.xlsx
se_stores_list = [
    {
        'employee_name': 'Não Thị Anh Đào',
        'chain': '7-Eleven',
        'store_code': 'CVS_7ELEVEN.79.2106',
        'store_address': 'B1.01.02, Block B1, KCH - TMDV cao tầng (Opal Boulevard), số 10 Kha Vạn Cân, Phường An Bình, Thành phố Dĩ An, Tỉnh Bình Dương',
        'actual': round(se_per_store)
    },
    {
        'employee_name': 'Não Thị Anh Đào',
        'chain': '7-Eleven',
        'store_code': 'CVS_7ELEVEN.79.2087',
        'store_address': 'Đường Thống Nhất, Phường Đông Hoà, TP Hồ Chí Minh',
        'actual': round(se_per_store)
    },
    {
        'employee_name': 'Não Thị Anh Đào',
        'chain': '7-Eleven',
        'store_code': 'CVS_7ELEVEN.79.2105',
        'store_address': 'R01.20, tòa Ruby Charm City, 115 đường ĐT 743C',
        'actual': round(se_per_store)
    },
    {
        'employee_name': 'Nguyễn Đức Hoà',
        'chain': '7-Eleven',
        'store_code': 'CVS_7ELEVEN.79.2164',
        'store_address': 'Căn TMDV Số 4, , Khu B, Lô C17, Đại lộ Hùng Vương, Chung cư SORA garden II, Phường Bình Dương, TP Hồ Chí Minh',
        'actual': round(se_per_store)
    },
    {
        'employee_name': 'Nguyễn Đức Hoà',
        'chain': '7-Eleven',
        'store_code': 'CVS_7ELEVEN.79.2104',
        'store_address': 'Tòa A Happy One Central, 113 đường 30/4, Phường Phú Hòa, Thủ Dầu Một, Bình Dương',
        'actual': round(se_per_store)
    }
]

# Danh sách 4 cửa hàng WinMart+ (WMP) của Nguyễn Đức Hòa
wmp_stores_list = [
    {
        'employee_name': 'Nguyễn Đức Hoà',
        'chain': 'WinMart+',
        'store_code': 'WMP_DNI_KP_TRUNG_TAM_XUAN_LAP',
        'store_address': 'Thửa đất số 58, TBĐ số 54, KP. Trung Tâm, P. Xuân Lập, Thành phố Đồng Nai',
        'actual': round(wmp_amounts['WMP_DNI_KP_TRUNG_TAM_XUAN_LAP'])
    },
    {
        'employee_name': 'Nguyễn Đức Hoà',
        'chain': 'WinMart+',
        'store_code': 'WMP_DNI_285_287_CACH_MANG_THAN',
        'store_address': '285-287 Cách Mạng Tháng Tám, Phường Trấn Biên, Thành phố Đồng Nai',
        'actual': round(wmp_amounts['WMP_DNI_285_287_CACH_MANG_THAN'])
    },
    {
        'employee_name': 'Nguyễn Đức Hoà',
        'chain': 'WinMart+',
        'store_code': 'WMP_HCM_90A_92_PHAN_CHU_TRINH',
        'store_address': '90A-92 Phan Chu Trinh, Phường Vũng Tàu, TP. Hồ Chí Minh',
        'actual': round(wmp_amounts['WMP_HCM_90A_92_PHAN_CHU_TRINH'])
    },
    {
        'employee_name': 'Nguyễn Đức Hoà',
        'chain': 'WinMart+',
        'store_code': 'WMP_DNI_A_01_04_TOPAZ_TWINS',
        'store_address': 'A01-04, Chung cư Topaz Twins, đường số 7, KP. Vinh Thạnh, P. Trấn Biên, T. Đồng Nai',
        'actual': round(wmp_amounts['WMP_DNI_A_01_04_TOPAZ_TWINS'])
    }
]

new_stores = []
ck_matched_count = 0
ck_total_count = 0
for s in old_stores:
    emp = s['employee_name']
    ch = s['chain']
    addr = s['store_address']
    code = s['store_code']
    old_act = float(s['actual'])
    new_act = 0

    if ch == 'GS25':
        new_act = round(gs25_per_store)

    elif ch in ['7E', '7-Eleven']:
        new_act = round(se_per_store)

    elif ch == 'Circle K':
        ck_total_count += 1
        raw_amt = 0
        found = False
        if addr in ck_stores_so:
            raw_amt = ck_stores_so[addr]
            found = True
        else:
            norm_s = normalize_text(addr)
            for a_so, amt in ck_stores_so.items():
                if norm_s == normalize_text(a_so):
                    raw_amt = amt
                    found = True
                    break
        new_act = round(T_ck + raw_amt)
        if found and raw_amt > 0:
            ck_matched_count += 1

    elif 'Hoàng Đức' in ch or ch in ['Hoàng Đức Long Khánh', 'Hoàng Đức Gia Kiệm']:
        if '198 Hùng Vương' in addr:
            new_act = 74676300
        elif 'Bạch Lâm' in addr:
            new_act = 75553900
        else:
            new_act = old_act

    elif ch == 'FamilyMart':
        first_part = addr.split(',')[0].strip().lower()
        if code and code in fm_matrix_dict:
            new_act = round(fm_matrix_dict[code])
        elif first_part in fm_matrix_dict:
            new_act = round(fm_matrix_dict[first_part])
        else:
            matched_val = None
            for fk, fv in fm_matrix_dict.items():
                if len(fk) > 3 and (fk in first_part or first_part in fk or fk in addr.lower()):
                    matched_val = fv
                    break
            if matched_val is not None:
                new_act = round(matched_val)
            else:
                new_act = old_act

    elif ch in ['WinMart+', 'WMP', 'WIN']:
        if code in wmp_amounts:
            new_act = round(wmp_amounts[code])
        else:
            new_act = old_act
    else:
        new_act = old_act

    new_stores.append({
        'employee_name': emp,
        'chain': ch,
        'store_code': code,
        'store_address': addr,
        'actual': int(new_act)
    })

# Đảm bảo 5 cửa hàng 7E có trong danh sách
existing_7e_codes = set([s['store_code'] for s in new_stores if s['chain'] in ['7E', '7-Eleven']])
for se in se_stores_list:
    if se['store_code'] not in existing_7e_codes:
        new_stores.append(se)

# Đảm bảo 4 cửa hàng WinMart+ của Nguyễn Đức Hòa có trong danh sách
existing_wmp_codes = set([s['store_code'] for s in new_stores if s['chain'] in ['WinMart+', 'WMP', 'WIN']])
for wmp in wmp_stores_list:
    if wmp['store_code'] not in existing_wmp_codes:
        new_stores.append(wmp)

print(f"\nCircle K: Khớp đơn hàng {ck_matched_count}/{ck_total_count} CH (+ hệ số T = {round(T_ck):,d} VNĐ/CH)")
print(f"7-Eleven: Đã cập nhật {len(se_stores_list)} CH ({round(se_per_store):,d} VNĐ/CH)")
print(f"GS25: Đã áp dụng công thức mới cho {len([s for s in new_stores if s['chain'] == 'GS25'])} CH ({round(gs25_per_store):,d} VNĐ/CH)")
print(f"WinMart+: Đã cập nhật 4 CH cho Nguyễn Đức Hòa (Tổng: {sum(w['actual'] for w in wmp_stores_list):,d} VNĐ)")

# Ghi file team_cvs_stores.csv
target_dirs = [
    script_dir,
    r'C:\Users\giang\.gemini\config\skills\team-cvs-bhx-tracker'
]

for td in set(target_dirs):
    if os.path.exists(td):
        target_path = os.path.join(td, 'team_cvs_stores.csv')
        with open(target_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['employee_name', 'chain', 'store_code', 'store_address', 'actual'])
            writer.writeheader()
            writer.writerows(new_stores)
        print(f"Đã lưu: {target_path}")

# ----------------------------------------------------
# 9. CẬP NHẬT TEAM_CONFIG.CSV
# ----------------------------------------------------
from team_leader_report_v4 import get_realtime_timegone
realtime_time_pct = get_realtime_timegone(8, 2026)

config_data = [
    ('team_lead', 'Trần Thị Cẩm Giang'),
    ('month', '8'),
    ('year', '2026'),
    ('total_target_bhx', '37382000000'),
    ('total_actual_bhx', str(round(total_actual_bhx))),
    ('total_stores_bhx', str(total_stores_bhx)),
    ('time_percentage', str(realtime_time_pct)),
    ('output_file', 'Team_CamGiang_Report.xlsx')
]

for td in set(target_dirs):
    if os.path.exists(td):
        target_path = os.path.join(td, 'team_config.csv')
        with open(target_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['key', 'value'])
            writer.writerows(config_data)
        print(f"Đã lưu: {target_path} (BHX Total: {total_actual_bhx:,.0f} VNĐ, % Timegone: {realtime_time_pct}%)")

print("\n=== ĐÃ HOÀN TẤT CẬP NHẬT CONFIG & DANH SÁCH CỬA HÀNG CVS THÀNH CÔNG ===")
