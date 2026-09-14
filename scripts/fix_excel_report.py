import openpyxl
from copy import copy

wb = openpyxl.load_workbook('Team_CamGiang_Report.xlsx')
ws = wb['Chi tiết CVS']

# Find the row where STT is 102
target_row_idx = None
for row_idx in range(1, ws.max_row + 1):
    val = ws.cell(row=row_idx, column=1).value
    if val == 102:
        target_row_idx = row_idx
        break

print(f"Found STT 102 at row {target_row_idx}")
if target_row_idx:
    # Insert a new row after target_row_idx
    insert_at = target_row_idx + 1
    ws.insert_rows(insert_at)
    
    # Fill in values for STT 103
    # Columns: 1: STT, 2: Nhân Viên, 3: Chuỗi, 4: Mã CH, 5: Địa Chỉ, 6: Thực Hiện
    ws.cell(row=insert_at, column=1, value=103)
    ws.cell(row=insert_at, column=2, value="Nguyễn Thị Thanh Thủy")
    ws.cell(row=insert_at, column=3, value="Circle K")
    ws.cell(row=insert_at, column=4, value="CVS_CIRCLEK.79.4260")
    ws.cell(row=insert_at, column=5, value="117/21, Thùy Vân, P.Vũng Tàu, TP.Hồ Chí Minh")
    ws.cell(row=insert_at, column=6, value=2985957)
    
    # Copy style from row 102
    for col in range(1, 7):
        src_cell = ws.cell(row=target_row_idx, column=col)
        dst_cell = ws.cell(row=insert_at, column=col)
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.border = copy(src_cell.border)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.number_format = copy(src_cell.number_format)
            dst_cell.protection = copy(src_cell.protection)
            dst_cell.alignment = copy(src_cell.alignment)

    # Re-number subsequent STT if needed (from insert_at + 1 to end of stores)
    # Let's check existing STTs after 102
    print("Inserted row 103 successfully.")

# Also update the header row of Nguyen Thi Thanh Thuy: (39 cửa hàng CVS)
for r in range(1, ws.max_row + 1):
    v = str(ws.cell(row=r, column=1).value or '')
    if 'Nguyễn Thị Thanh Thủy' in v and 'cửa hàng CVS' in v:
        ws.cell(row=r, column=1, value='👤 Nguyễn Thị Thanh Thủy (39 cửa hàng CVS)')
        print(f"Updated Thuy header at row {r}")

# Check total row at bottom
for r in range(ws.max_row - 10, ws.max_row + 1):
    v = str(ws.cell(row=r, column=1).value or '')
    if 'TỔNG CVS TOÀN TEAM' in v:
        print(f"Total row found at {r}: {ws.cell(row=r, column=6).value}")

wb.save('Team_CamGiang_Report.xlsx')
print("Saved Team_CamGiang_Report.xlsx successfully.")
