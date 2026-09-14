/**
 * ============================================================================
 * GOOGLE APPS SCRIPT BACKEND: HỆ THỐNG BÁO CÁO DOANH SỐ TEAM CVS + BHX
 * Tác giả: Team Lead Trần Thị Cẩm Giang
 * Output: Chuẩn giao diện và cấu trúc Team_CamGiang_Report.xlsx
 * ============================================================================
 */

/**
 * Phục vụ Giao diện Web App (HTML5 + CSS3 + JS)
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'Index';
  try {
    return HtmlService.createHtmlOutputFromFile(page)
      .setTitle('Báo Cáo Doanh Số Team (CVS & BHX) — Command Center Hybrid')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (err) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Báo Cáo Doanh Số Team (CVS & BHX) — Command Center Hybrid')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
}

/**
 * Phục vụ API POST từ Vercel Webhook / Backend
 */
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    if (data.action === 'create_sheet') {
      var res = createGoogleSheetReport(data.payload);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Lấy Master Data (Danh sách NV, Target, Cửa hàng CVS, Hubs BHX) đã lưu
 */
function getMasterDataFromBackend() {
  try {
    var props = PropertiesService.getScriptProperties().getProperty('CUSTOM_MASTER_DATA');
    if (props) {
      return { success: true, data: JSON.parse(props) };
    }
  } catch (e) {
    Logger.log('Lỗi đọc custom master data: ' + e);
  }
  return { success: false, data: null };
}

/**
 * Lưu cấu hình Master Data / KPI Target tháng mới vào bộ nhớ đám mây của Google Apps Script
 */
function saveCustomMasterData(data) {
  try {
    PropertiesService.getScriptProperties().setProperty('CUSTOM_MASTER_DATA', JSON.stringify(data));
    return {
      success: true,
      message: 'Đã lưu cấu hình Target Tháng ' + (data.month || '') + '/' + (data.year || '') + ' thành công!'
    };
  } catch (e) {
    return { success: false, message: 'Lỗi lưu cấu hình: ' + e.toString() };
  }
}

/**
 * Tạo Google Spreadsheet báo cáo tự động với 5 sheets chuẩn format như Team_CamGiang_Report.xlsx
 * @param {Object} payload Dữ liệu đã được tổng hợp từ 2 file
 */
function createGoogleSheetReport(payload) {
  try {
    var month = payload.month || (new Date().getMonth() + 1);
    var year = payload.year || new Date().getFullYear();
    var teamLead = payload.team_lead || 'Trần Thị Cẩm Giang';
    var timePct = parseFloat(payload.timegone) || 0;

    var ssTitle = 'Team_' + teamLead.replace(/\s+/g, '_') + '_Report_Thang_' + month + '_' + year;
    var ss = SpreadsheetApp.create(ssTitle);

    var defaultSheet = ss.getSheets()[0];

    // 1. Tạo Sheet "Dashboard"
    createDashboardSheet(ss, payload, month, year, teamLead, timePct);

    // 2. Tạo Sheet "Tiến độ Team"
    createTeamProgressSheet(ss, payload, month, year, teamLead, timePct);

    // 3. Tạo Sheet "Chi tiết CVS"
    createCvsDetailsSheet(ss, payload);

    // 4. Tạo Sheet "Data BHX"
    createBhxDataSheet(ss, payload);

    // 5. Tạo Sheet "Config"
    createConfigSheet(ss, payload, month, year, teamLead, timePct);

    // 6. Tạo Sheet "Chi tiết SKU FamilyMart"
    createFamilyMartSkuSheet(ss, payload, month, year, timePct);

    // Xóa sheet rỗng ban đầu
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {}

    var dash = ss.getSheetByName('Dashboard');
    if (dash) ss.setActiveSheet(dash);

    // Tự động cấp quyền xem cho bất kỳ ai có liên kết để cấp trên xem trực tiếp không cần cấp quyền thủ công
    try {
      var file = DriveApp.getFileById(ss.getId());
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {
      Logger.log('Không thể set quyền public tự động: ' + eShare);
    }

    return {
      success: true,
      url: ss.getUrl(),
      id: ss.getId(),
      name: ss.getName()
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * ----------------------------------------------------------------------------
 * SHEET 1: DASHBOARD
 * Chuẩn 100% theo Team_CamGiang_Report.xlsx
 * ----------------------------------------------------------------------------
 */
function createDashboardSheet(ss, p, month, year, teamLead, timePct) {
  var ws = ss.insertSheet('Dashboard', 0);
  ws.setTabColor('#1F4E78');

  // Column Widths
  var widths = [30, 80, 200, 150, 150, 150, 150, 120];
  for (var c = 1; c <= widths.length; c++) {
    ws.setColumnWidth(c, widths[c - 1]);
  }

  // Row 2: Banner Title B2:H2
  var titleRange = ws.getRange('B2:H2');
  titleRange.merge();
  titleRange.setValue('📊 THEO DÕI DOANH SỐ TEAM - ' + teamLead.toUpperCase());
  titleRange.setFontSize(16).setFontWeight('bold').setFontColor('#FFFFFF').setFontFamily('Calibri');
  titleRange.setBackground('#1F4E78').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(2, 35);

  // Row 3: Metadata B3:H3
  ws.getRange('B3').setValue('Tháng:').setFontWeight('bold').setBackground('#FFF2CC').setFontFamily('Calibri');
  ws.getRange('C3').setValue(month + '/' + year).setFontWeight('bold').setFontColor('#C00000').setBackground('#FFF9E6').setHorizontalAlignment('center');
  
  ws.getRange('D3').setValue('Team Lead:').setFontWeight('bold').setBackground('#FFF2CC').setFontFamily('Calibri');
  ws.getRange('E3').setValue(teamLead).setFontWeight('bold').setFontColor('#1F4E78').setBackground('#FFF9E6').setHorizontalAlignment('center');

  ws.getRange('F3').setValue('% Timegone:').setFontWeight('bold').setBackground('#FFF2CC').setFontFamily('Calibri');
  ws.getRange('G3').setValue(timePct / 100).setNumberFormat('0.0%').setFontWeight('bold').setFontColor('#C00000').setBackground('#FFF9E6').setHorizontalAlignment('center');

  var nowStr = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  ws.getRange('H3').setValue('Cập nhật: ' + nowStr).setBackground('#FFF9E6').setHorizontalAlignment('center').setFontSize(10);
  ws.setRowHeight(3, 24);

  // Row 5: Section Title B5:H5
  var secTitle = ws.getRange('B5:H5');
  secTitle.merge();
  secTitle.setValue('🎯 TỔNG QUAN TEAM').setFontSize(13).setFontWeight('bold').setFontColor('#1F4E78').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(5, 26);

  // Row 6: KPI Labels & Row 7: KPI Values
  var totalTarget = p.total_target || 0;
  var totalBhx = p.total_actual_bhx || 0;
  var totalCvs = p.total_cvs || 0;
  var totalActual = p.total_actual || (totalBhx + totalCvs);
  var missing = totalTarget - totalActual;
  var teamPct = totalTarget > 0 ? (totalActual / totalTarget) : 0;

  var kpiPctColor = (teamPct * 100 >= timePct ? '#28E11F' : (teamPct * 100 >= timePct * 0.75 ? '#F5991F' : '#DD2323'));

  var kpis = [
    { label: 'Target Team', val: totalTarget, fmt: '#,##0', bg: '#4472C4', fg: '#FFFFFF' },
    { label: 'Thực Hiện BHX', val: totalBhx, fmt: '#,##0', bg: '#70AD47', fg: '#FFFFFF' },
    { label: 'Thực Hiện CVS', val: totalCvs, fmt: '#,##0', bg: '#ED7D31', fg: '#FFFFFF' },
    { label: 'Tổng Thực Hiện', val: totalActual, fmt: '#,##0', bg: '#2E75B6', fg: '#FFFFFF' },
    { label: 'Thiếu', val: missing, fmt: '#,##0;[Red](#,##0)', bg: '#DD2323', fg: '#FFFFFF' },
    { label: '% Đạt Team', val: teamPct, fmt: '0.0%', bg: kpiPctColor, fg: (teamPct * 100 >= timePct ? '#000000' : '#FFFFFF') }
  ];

  for (var i = 0; i < kpis.length; i++) {
    var col = 2 + i;
    var lbl = ws.getRange(6, col);
    lbl.setValue(kpis[i].label).setFontWeight('bold').setFontSize(10).setFontColor(kpis[i].fg).setBackground(kpis[i].bg)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

    var val = ws.getRange(7, col);
    val.setValue(kpis[i].val).setNumberFormat(kpis[i].fmt).setFontWeight('bold').setFontSize(12)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#F2F2F2');
  }
  ws.setRowHeight(6, 30);
  ws.setRowHeight(7, 28);

  // Row 9: Status Bar B9:H9
  var diff = (teamPct * 100) - timePct;
  var statusText = '';
  var statusBg = '#28E11F';
  var statusFg = '#000000';

  if (teamPct * 100 >= timePct) {
    statusText = '🟢 Team đang VƯỢT tiến độ: % Đạt (' + (teamPct * 100).toFixed(1) + '%) vs % Timegone (' + timePct.toFixed(1) + '%) → chênh lệch ' + (diff >= 0 ? '+' : '') + diff.toFixed(1) + '%';
    statusBg = '#28E11F';
    statusFg = '#000000';
  } else if (teamPct * 100 >= timePct * 0.75) {
    statusText = '🟠 Team đang CẬN tiến độ (CẢNH BÁO): % Đạt (' + (teamPct * 100).toFixed(1) + '%) vs % Timegone (' + timePct.toFixed(1) + '%) → chênh lệch ' + diff.toFixed(1) + '%';
    statusBg = '#F5991F';
    statusFg = '#FFFFFF';
  } else {
    statusText = '🔴 Team đang CHẬM tiến độ: % Đạt (' + (teamPct * 100).toFixed(1) + '%) vs % Timegone (' + timePct.toFixed(1) + '%) → chênh lệch ' + diff.toFixed(1) + '%';
    statusBg = '#DD2323';
    statusFg = '#FFFFFF';
  }

  var sBar = ws.getRange('B9:H9');
  sBar.merge();
  sBar.setValue(statusText).setFontWeight('bold').setFontSize(11).setFontColor(statusFg).setBackground(statusBg)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(9, 25);

  // Row 11: Header Bảng Xếp Hạng
  var rTitle = ws.getRange('B11:H11');
  rTitle.merge();
  rTitle.setValue('XẾP HẠNG NHÂN VIÊN').setFontSize(13).setFontWeight('bold').setFontColor('#1F4E78')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(11, 26);

  // Row 12: Headers B12:H12
  var rankHeaders = ['Hạng', 'Nhân Viên', 'Target', 'BHX', 'CVS', 'Tổng TH', '% Đạt'];
  for (var j = 0; j < rankHeaders.length; j++) {
    ws.getRange(12, 2 + j).setValue(rankHeaders[j]).setFontWeight('bold').setFontSize(10).setFontColor('#FFFFFF')
      .setBackground('#4472C4').setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  ws.setRowHeight(12, 26);

  // Sắp xếp theo % Đạt giảm dần
  var sorted = (p.employees || []).slice().sort(function(a, b) {
    return (b.percent || 0) - (a.percent || 0);
  });

  for (var k = 0; k < sorted.length; k++) {
    var emp = sorted[k];
    var r = 13 + k;

    ws.getRange(r, 2).setValue(k + 1).setHorizontalAlignment('center').setFontWeight('bold');
    ws.getRange(r, 3).setValue(emp.name).setFontWeight('bold');
    ws.getRange(r, 4).setValue(emp.target).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 5).setValue(emp.bhx_actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 6).setValue(emp.cvs_total).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 7).setValue(emp.total_actual).setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('right');

    var pct = (emp.percent || 0) / 100;
    var pctCell = ws.getRange(r, 8);
    pctCell.setValue(pct).setNumberFormat('0.0%').setFontWeight('bold').setHorizontalAlignment('center');

    if (emp.percent >= timePct) {
      pctCell.setBackground('#28E11F').setFontColor('#000000');
    } else if (emp.percent >= timePct * 0.75) {
      pctCell.setBackground('#F5991F').setFontColor('#FFFFFF');
    } else {
      pctCell.setBackground('#DD2323').setFontColor('#FFFFFF');
    }

    ws.setRowHeight(r, 22);
  }

  // Borders B12:H(12 + sorted.length)
  var endRankRow = 12 + sorted.length;
  ws.getRange('B12:H' + endRankRow).setBorder(true, true, true, true, true, true, '#B4B4B4', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * ----------------------------------------------------------------------------
 * SHEET 2: TIẾN ĐỘ TEAM
 * Chuẩn 100% theo Team_CamGiang_Report.xlsx
 * ----------------------------------------------------------------------------
 */
function createTeamProgressSheet(ss, p, month, year, teamLead, timePct) {
  var ws = ss.insertSheet('Tiến độ Team', 1);
  ws.setTabColor('#70AD47');

  // Exact Column Widths
  var colW = [15, 45, 175, 135, 125, 105, 105, 105, 105, 105, 105, 145, 95];
  for (var c = 1; c <= colW.length; c++) {
    ws.setColumnWidth(c, colW[c - 1]);
  }

  // Row 2: Banner Title B2:M2
  var banner = ws.getRange('B2:M2');
  banner.merge();
  banner.setValue('THEO DÕI DOANH SỐ - TEAM ' + teamLead.toUpperCase() + ' - THÁNG ' + month + '/' + year);
  banner.setFontSize(14).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E78')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(2, 30);

  // Row 3: Info Row
  ws.getRange('B3').setValue('% Timegone:').setFontWeight('bold').setBackground('#FFF2CC');
  ws.getRange('C3').setValue(timePct / 100).setNumberFormat('0.0%').setFontWeight('bold').setFontColor('#C00000')
    .setBackground('#FFF9E6').setHorizontalAlignment('center').setFontSize(12);

  ws.getRange('E3').setValue('Tổng CH BHX:').setFontWeight('bold').setBackground('#FFF2CC');
  ws.getRange('F3').setValue(p.total_stores_bhx || 232).setFontWeight('bold').setFontColor('#C00000')
    .setBackground('#FFF9E6').setHorizontalAlignment('center');

  ws.getRange('H3').setValue('TB BHX/CH:').setFontWeight('bold').setBackground('#FFF2CC');
  ws.getRange('I3').setValue(p.bhx_per_store || 0).setNumberFormat('#,##0').setFontWeight('bold').setFontColor('#C00000')
    .setBackground('#FFF9E6').setHorizontalAlignment('center');
  ws.setRowHeight(3, 24);

  // Row 5: Group Headers
  var grp1 = ws.getRange('E5:K5');
  grp1.merge();
  grp1.setValue('THỰC HIỆN (chi tiết theo kênh)').setFontWeight('bold').setFontSize(11).setFontColor('#FFFFFF')
    .setBackground('#2E75B6').setHorizontalAlignment('center').setVerticalAlignment('middle');

  var grp2 = ws.getRange('L5:M5');
  grp2.merge();
  grp2.setValue('ĐÁNH GIÁ').setFontWeight('bold').setFontSize(11).setFontColor('#FFFFFF')
    .setBackground('#DD2323').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(5, 22);

  // Row 6: Column Headers B6:M6
  var headers = [
    'STT', 'Nhân Viên', 'Target', 'BHX\n(cố định)', 'GS25', '7-Eleven', 'FamilyMart', 'Circle K', 'Hoàng Đức', 'WinMart+', 'Tổng TH\n(BHX+CVS)', '% Đạt'
  ];
  for (var i = 0; i < headers.length; i++) {
    var cell = ws.getRange(6, 2 + i);
    cell.setValue(headers[i]).setFontWeight('bold').setFontSize(10).setFontColor('#FFFFFF')
      .setBackground('#4472C4').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  }
  ws.setRowHeight(6, 40);

  // Rows 7+: 10 Employees
  var emps = p.employees || [];
  var startRow = 7;
  for (var k = 0; k < emps.length; k++) {
    var emp = emps[k];
    var r = startRow + k;

    ws.getRange(r, 2).setValue(k + 1).setHorizontalAlignment('center');
    ws.getRange(r, 3).setValue(emp.name).setFontWeight('bold');
    ws.getRange(r, 4).setValue(emp.target).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 5).setValue(emp.bhx_actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 6).setValue(emp.gs25_actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 7).setValue(emp.se_actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 8).setValue(emp.fm_actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 9).setValue(emp.ck_actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 10).setValue(emp.hd_actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 11).setValue(emp.wmp_actual).setNumberFormat('#,##0').setHorizontalAlignment('right');

    // Tổng TH = SUM(E..:K..)
    ws.getRange(r, 12).setFormula('=SUM(E' + r + ':K' + r + ')').setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('right');

    // % Đạt = L.. / D..
    var pct = ws.getRange(r, 13);
    pct.setFormula('=IF(D' + r + '>0, L' + r + '/D' + r + ', 0)').setNumberFormat('0.0%').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

    if (emp.percent >= timePct) {
      pct.setBackground('#28E11F').setFontColor('#000000');
    } else if (emp.percent >= timePct * 0.75) {
      pct.setBackground('#F5991F').setFontColor('#FFFFFF');
    } else {
      pct.setBackground('#DD2323').setFontColor('#FFFFFF');
    }

    ws.setRowHeight(r, 22);
  }

  // Row Total: TỔNG TEAM
  var totalRow = startRow + emps.length;
  var totLabel = ws.getRange(totalRow, 2, 1, 2);
  totLabel.merge();
  totLabel.setValue('TỔNG TEAM').setFontWeight('bold').setFontSize(12).setFontColor('#FFFFFF')
    .setBackground('#1F4E78').setHorizontalAlignment('center').setVerticalAlignment('middle');

  for (var col = 4; col <= 12; col++) {
    var colL = String.fromCharCode(64 + col);
    ws.getRange(totalRow, col).setFormula('=SUM(' + colL + startRow + ':' + colL + (totalRow - 1) + ')')
      .setNumberFormat('#,##0').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E78').setHorizontalAlignment('right');
  }

  // % Đạt Total
  ws.getRange(totalRow, 13).setFormula('=L' + totalRow + '/D' + totalRow).setNumberFormat('0.0%')
    .setFontWeight('bold').setFontSize(12).setFontColor('#FFFF00').setBackground('#1F4E78').setHorizontalAlignment('center');
  ws.setRowHeight(totalRow, 28);

  // Full Borders B6:M(totalRow)
  ws.getRange('B6:M' + totalRow).setBorder(true, true, true, true, true, true, '#B4B4B4', SpreadsheetApp.BorderStyle.SOLID);

  // Freeze Panes at D7 (STT, Name, Target visible)
  ws.setFrozenRows(6);
  ws.setFrozenColumns(3);

  // Legend rows below
  var legendRow = totalRow + 2;
  var legRange = ws.getRange(legendRow, 2, 1, 12);
  legRange.merge();
  legRange.setValue('📌 % Đạt ≥ % Timegone (' + timePct.toFixed(1) + '%) → Vượt tiến độ 🟢 (#28E11F) | Cận Timegone → Cảnh báo 🟠 (#F5991F) | Thấp hơn nhiều → Cần đẩy nhanh 🔴 (#DD2323)')
    .setFontStyle('italic').setFontSize(10).setFontColor('#555555').setHorizontalAlignment('center');

  var fmlaRow = legendRow + 1;
  var fmlaRange = ws.getRange(fmlaRow, 2, 1, 12);
  fmlaRange.merge();
  fmlaRange.setValue('🧮 BHX cố định = (Tổng BHX ÷ 232) × Số CH phụ trách | Thực Hiện = BHX + GS25 + 7E + FM + CK + HĐ + WMP | % Đạt = TH ÷ Target')
    .setFontStyle('italic').setFontSize(9).setFontColor('#555555').setHorizontalAlignment('center');
}

/**
 * ----------------------------------------------------------------------------
 * SHEET 3: CHI TIẾT CVS (Từng cửa hàng gom theo nhân viên)
 * Chuẩn 100% theo Team_CamGiang_Report.xlsx
 * ----------------------------------------------------------------------------
 */
function createCvsDetailsSheet(ss, p) {
  var ws = ss.insertSheet('Chi tiết CVS', 2);
  ws.setTabColor('#ED7D31');

  // Column Widths
  var widths = [50, 160, 110, 140, 380, 140];
  for (var c = 1; c <= widths.length; c++) {
    ws.setColumnWidth(c, widths[c - 1]);
  }

  // Row 1: Title A1:F1
  var title = ws.getRange('A1:F1');
  title.merge();
  title.setValue('CHI TIẾT DOANH SỐ CVS - TỪNG CỬA HÀNG').setFontSize(14).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#1F4E78').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 30);

  // Row 3: Headers A3:F3
  var headers = ['STT', 'Nhân Viên', 'Chuỗi', 'Mã CH', 'Địa Chỉ', 'Thực Hiện'];
  for (var i = 0; i < headers.length; i++) {
    ws.getRange(3, 1 + i).setValue(headers[i]).setFontWeight('bold').setFontSize(10).setFontColor('#FFFFFF')
      .setBackground('#2E75B6').setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  ws.setRowHeight(3, 25);

  var stores = p.stores || [];
  var emps = p.employees || [];
  var rowIdx = 4;
  var stt = 1;

  var chainColors = {
    'GS25': '#4472C4',
    '7-Eleven': '#00B050',
    '7E': '#00B050',
    'FamilyMart': '#ED7D31',
    'Circle K': '#C00000',
    'Hoàng Đức': '#70AD47',
    'Hoàng Đức Long Khánh': '#70AD47',
    'Hoàng Đức Gia Kiệm': '#008080',
    'WinMart+': '#D9534F',
    'WMP': '#D9534F'
  };

  for (var e = 0; e < emps.length; e++) {
    var empName = emps[e].name;
    var empStores = stores.filter(function(s) { return s.employee_name === empName; });
    if (empStores.length === 0) continue;

    // Subtotal Header Row for Employee
    var subRange = ws.getRange(rowIdx, 1, 1, 4);
    subRange.merge();
    subRange.setValue('👤 ' + empName + ' (' + empStores.length + ' cửa hàng CVS)').setFontWeight('bold').setFontSize(11)
      .setFontColor('#FFFFFF').setBackground('#4472C4').setVerticalAlignment('middle');

    ws.getRange(rowIdx, 5).setValue('TỔNG CVS:').setFontWeight('bold').setFontColor('#FFFFFF')
      .setBackground('#4472C4').setHorizontalAlignment('right').setVerticalAlignment('middle');

    var empTotal = empStores.reduce(function(acc, s) { return acc + (s.actual || 0); }, 0);
    ws.getRange(rowIdx, 6).setValue(empTotal).setNumberFormat('#,##0').setFontWeight('bold').setFontSize(11)
      .setFontColor('#FFFF00').setBackground('#4472C4').setHorizontalAlignment('right').setVerticalAlignment('middle');
    ws.setRowHeight(rowIdx, 24);
    rowIdx++;

    for (var s = 0; s < empStores.length; s++) {
      var st = empStores[s];
      ws.getRange(rowIdx, 1).setValue(stt).setHorizontalAlignment('center');
      ws.getRange(rowIdx, 2).setValue(empName);

      var chColor = chainColors[st.chain] || '#70AD47';
      ws.getRange(rowIdx, 3).setValue(st.chain).setFontWeight('bold').setFontSize(10).setFontColor('#FFFFFF')
        .setBackground(chColor).setHorizontalAlignment('center');

      ws.getRange(rowIdx, 4).setValue(st.store_code || '').setFontSize(10);
      ws.getRange(rowIdx, 5).setValue(st.store_address || '').setFontSize(10);
      ws.getRange(rowIdx, 6).setValue(st.actual || 0).setNumberFormat('#,##0').setHorizontalAlignment('right');

      ws.setRowHeight(rowIdx, 20);
      rowIdx++;
      stt++;
    }
  }

  // Grand Total Row
  var grandRange = ws.getRange(rowIdx, 1, 1, 5);
  grandRange.merge();
  grandRange.setValue('🎯 TỔNG CVS TOÀN TEAM').setFontWeight('bold').setFontSize(12).setFontColor('#FFFFFF')
    .setBackground('#1F4E78').setHorizontalAlignment('center').setVerticalAlignment('middle');

  var grandTotal = stores.reduce(function(acc, s) { return acc + (s.actual || 0); }, 0);
  ws.getRange(rowIdx, 6).setValue(grandTotal).setNumberFormat('#,##0').setFontWeight('bold').setFontSize(12)
    .setFontColor('#FFFF00').setBackground('#1F4E78').setHorizontalAlignment('right').setVerticalAlignment('middle');
  ws.setRowHeight(rowIdx, 28);

  // Borders & Freeze Panes
  ws.getRange('A3:F' + rowIdx).setBorder(true, true, true, true, true, true, '#B4B4B4', SpreadsheetApp.BorderStyle.SOLID);
  ws.setFrozenRows(3);
}

/**
 * ----------------------------------------------------------------------------
 * SHEET 4: DATA BHX
 * Chuẩn 100% theo Team_CamGiang_Report.xlsx
 * ----------------------------------------------------------------------------
 */
function createBhxDataSheet(ss, p) {
  var ws = ss.insertSheet('Data BHX', 3);
  ws.setTabColor('#4472C4');

  // Column Widths
  var widths = [45, 180, 110, 130, 150, 110];
  for (var c = 1; c <= widths.length; c++) {
    ws.setColumnWidth(c, widths[c - 1]);
  }

  // Row 1: Title A1:F1
  var title = ws.getRange('A1:F1');
  title.merge();
  title.setValue('PHÂN BỔ DOANH SỐ BHX (' + (p.total_stores_bhx || 232) + ' CỬA HÀNG)').setFontSize(13).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#1F4E78').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 28);

  // Top Summary
  ws.getRange('A3').setValue('Tổng DS 5 Hubs Hệ Thống:').setFontWeight('bold').setBackground('#FFF2CC');
  ws.getRange('B3').setValue(p.total_actual_bhx_hubs || p.total_actual_bhx || 0).setNumberFormat('#,##0').setFontWeight('bold').setFontColor('#C00000').setBackground('#FFF9E6');

  ws.getRange('A4').setValue('Tổng số CH hệ thống:').setFontWeight('bold').setBackground('#FFF2CC');
  ws.getRange('B4').setValue(p.total_stores_bhx || 232).setFontWeight('bold').setBackground('#FFF9E6');

  ws.getRange('A5').setValue('TB doanh số/CH:').setFontWeight('bold').setBackground('#FFF2CC');
  ws.getRange('B5').setFormula('=B3/B4').setNumberFormat('#,##0').setFontWeight('bold').setFontColor('#C00000').setBackground('#FFF9E6');

  // Row 7: Headers A7:F7
  var headers = ['STT', 'Nhân Viên', 'Số CH BHX', 'TB/CH', 'Doanh Số BHX', '% BHX/Team'];
  for (var i = 0; i < headers.length; i++) {
    ws.getRange(7, 1 + i).setValue(headers[i]).setFontWeight('bold').setFontSize(10).setFontColor('#FFFFFF')
      .setBackground('#2E75B6').setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  ws.setRowHeight(7, 25);

  var emps = p.employees || [];
  var startRow = 8;
  var totalRow = startRow + emps.length;
  for (var k = 0; k < emps.length; k++) {
    var r = startRow + k;
    ws.getRange(r, 1).setValue(k + 1).setHorizontalAlignment('center');
    ws.getRange(r, 2).setValue(emps[k].name).setFontWeight('bold');
    ws.getRange(r, 3).setValue(emps[k].bhx_stores).setHorizontalAlignment('center');
    ws.getRange(r, 4).setFormula('=$B$5').setNumberFormat('#,##0').setHorizontalAlignment('right');
    ws.getRange(r, 5).setFormula('=C' + r + '*D' + r).setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('right');
    ws.getRange(r, 6).setFormula('=IF($E$' + totalRow + '>0, E' + r + '/$E$' + totalRow + ', 0)').setNumberFormat('0.0%').setHorizontalAlignment('center');
    ws.setRowHeight(r, 20);
  }

  // Row Total
  var totLabel = ws.getRange(totalRow, 1, 1, 2);
  totLabel.merge();
  totLabel.setValue('TỔNG CỘNG TEAM').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('center');

  ws.getRange(totalRow, 3).setFormula('=SUM(C' + startRow + ':C' + (totalRow - 1) + ')').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('center');
  ws.getRange(totalRow, 4).setFormula('=$B$5').setNumberFormat('#,##0').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('right');
  ws.getRange(totalRow, 5).setFormula('=SUM(E' + startRow + ':E' + (totalRow - 1) + ')').setNumberFormat('#,##0').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('right');
  ws.getRange(totalRow, 6).setFormula('=SUM(F' + startRow + ':F' + (totalRow - 1) + ')').setNumberFormat('0.0%').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('center');
  ws.setRowHeight(totalRow, 24);

  ws.getRange('A7:F' + totalRow).setBorder(true, true, true, true, true, true, '#B4B4B4', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * ----------------------------------------------------------------------------
 * SHEET 5: CONFIG
 * ----------------------------------------------------------------------------
 */
function createConfigSheet(ss, p, month, year, teamLead, timePct) {
  var ws = ss.insertSheet('Config', 4);
  ws.setTabColor('#595959');

  ws.setColumnWidth(1, 240);
  ws.setColumnWidth(2, 240);

  var title = ws.getRange('A1:B1');
  title.merge();
  title.setValue('THÔNG SỐ CẤU HÌNH HỆ THỐNG').setFontWeight('bold').setFontSize(12)
    .setFontColor('#FFFFFF').setBackground('#1F4E78').setHorizontalAlignment('center');
  ws.setRowHeight(1, 28);

  var configItems = [
    ['Team Lead', teamLead],
    ['Tháng Báo Cáo', month],
    ['Năm Báo Cáo', year],
    ['% Timegone Thực Tế', timePct + '%'],
    ['Tổng Cửa Hàng BHX Hệ Thống', p.total_stores_bhx || 232],
    ['Tổng Cửa Hàng GS25 Hệ Thống', 72],
    ['Hệ Số Phân Bổ GS25 (2 DC)', '20.46%'],
    ['Tổng Cửa Hàng 7-Eleven Phụ Trách', 5],
    ['Hệ Số Phân Bổ 7-Eleven (2 DC)', '3.62%'],
    ['Thời Gian Xuất Báo Cáo', Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss')]
  ];

  for (var i = 0; i < configItems.length; i++) {
    var r = 3 + i;
    ws.getRange(r, 1).setValue(configItems[i][0]).setFontWeight('bold').setBackground('#F2F2F2');
    ws.getRange(r, 2).setValue(configItems[i][1]).setHorizontalAlignment('center');
    ws.setRowHeight(r, 22);
  }

  ws.getRange('A3:B' + (3 + configItems.length - 1)).setBorder(true, true, true, true, true, true, '#B4B4B4', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * ----------------------------------------------------------------------------
 * SHEET 6: CHI TIẾT SKU FAMILYMART
 * ----------------------------------------------------------------------------
 */
function createFamilyMartSkuSheet(ss, p, month, year, timePct) {
  var ws = ss.insertSheet('Chi tiết SKU FamilyMart', 5);
  ws.setTabColor('#ED7D31');

  var fm = p.fm_sku_matrix;
  var stores = (fm && fm.stores && fm.stores.length > 0) ? fm.stores : [];
  var cats = (fm && fm.categories && fm.categories.length > 0) ? fm.categories : [];

  if (stores.length === 0 || cats.length === 0) {
    return;
  }

  // Flatten all SKUs with category info
  var allSkus = [];
  for (var c = 0; c < cats.length; c++) {
    var cat = cats[c];
    for (var s = 0; s < cat.skus.length; s++) {
      allSkus.push({
        code: cat.skus[s].code,
        name: cat.skus[s].name,
        catName: cat.name,
        headerBg: '#' + (cat.headerBg ? cat.headerBg.replace(/^FF/, '') : 'E2EFDA'),
        cellBg: '#' + (cat.cellBg ? cat.cellBg.replace(/^FF/, '') : 'FFFFFF')
      });
    }
  }

  var totalCols = 8 + allSkus.length; // 8 base cols + 46 skus = 54 cols

  // Column widths
  ws.setColumnWidth(1, 280); // Dia chi
  ws.setColumnWidth(2, 110); // Location
  ws.setColumnWidth(3, 90);  // Ma PG
  ws.setColumnWidth(4, 150); // Ho ten PG
  ws.setColumnWidth(5, 60);  // VTCV
  ws.setColumnWidth(6, 130); // Target
  ws.setColumnWidth(7, 140); // Actual
  ws.setColumnWidth(8, 80);  // % Dat
  for (var k = 0; k < allSkus.length; k++) {
    ws.setColumnWidth(9 + k, 110);
  }

  // Row 1: Title Banner A1:L1
  var tRange = ws.getRange('A1:L1');
  tRange.merge();
  tRange.setValue('🛒 BÁO CÁO CHI TIẾT NHẬP HÀNG THEO CỬA HÀNG & SKU - CHUỖI FAMILYMART')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E78')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 32);

  // Row 2 & 3: 7 KPI Cards
  var totTgt = fm.total_target || 0;
  var totAct = fm.total_actual || 0;
  var pctTotal = totTgt > 0 ? (totAct / totTgt * 100) : 0;
  var passedCount = fm.passed_count || 0;

  var kpiBg = '#DD2323';
  var evalText = '🔴 CHẬM (' + (pctTotal - timePct).toFixed(1) + '%)';
  if (pctTotal >= timePct) {
    kpiBg = '#28E11F';
    evalText = '🟢 VƯỢT (+' + (pctTotal - timePct).toFixed(1) + '%)';
  } else if (pctTotal >= timePct * 0.75) {
    kpiBg = '#F5991F';
    evalText = '🟠 CẬN (' + (pctTotal - timePct).toFixed(1) + '%)';
  }

  var kpis = [
    ['Số Cửa Hàng', stores.length + ' CH', '#1F4E78', '#FFFFFF'],
    ['Tổng Chỉ Tiêu', Utilities.formatString('%,d đ', Math.round(totTgt)), '#2E75B6', '#FFFFFF'],
    ['Tổng Thực Hiện', Utilities.formatString('%,d đ', Math.round(totAct)), '#70AD47', '#FFFFFF'],
    ['% Đạt', pctTotal.toFixed(1) + '%', kpiBg, pctTotal >= timePct ? '#000000' : '#FFFFFF'],
    ['% Timegone', timePct.toFixed(1) + '%', '#ED7D31', '#FFFFFF'],
    ['Đánh Giá', evalText, kpiBg, pctTotal >= timePct ? '#000000' : '#FFFFFF'],
    ['CH Đạt Tiến Độ', passedCount + '/' + stores.length + ' CH (' + Math.round(passedCount/stores.length*100) + '%)', '#4472C4', '#FFFFFF']
  ];

  for (var i = 0; i < kpis.length; i++) {
    var c = i + 1;
    ws.getRange(2, c).setValue(kpis[i][0]).setFontSize(9).setFontWeight('bold').setFontColor('#555555')
      .setBackground('#F2F2F2').setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(3, c).setValue(kpis[i][1]).setFontSize(11).setFontWeight('bold').setFontColor(kpis[i][3])
      .setBackground(kpis[i][2]).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  ws.setRowHeight(2, 18);
  ws.setRowHeight(3, 24);

  // Row 4: Group Headers
  var g1 = ws.getRange('A4:E4');
  g1.merge();
  g1.setValue('THÔNG TIN ĐIỂM BÁN (STORE MASTER)').setFontSize(11).setFontWeight('bold').setFontColor('#1F4E78')
    .setBackground('#FFF2CC').setHorizontalAlignment('center').setVerticalAlignment('middle');

  var g2 = ws.getRange('F4:H4');
  g2.merge();
  g2.setValue('CHỈ TIÊU & TỔNG THỰC HIỆN').setFontSize(11).setFontWeight('bold').setFontColor('#C00000')
    .setBackground('#FCE4D6').setHorizontalAlignment('center').setVerticalAlignment('middle');

  var curCol = 9;
  for (var cIdx = 0; cIdx < cats.length; cIdx++) {
    var catItem = cats[cIdx];
    var skuCount = catItem.skus.length;
    var catRange = ws.getRange(4, curCol, 1, skuCount);
    if (skuCount > 1) catRange.merge();
    catRange.setValue(catItem.name).setFontSize(10).setFontWeight('bold').setFontColor('#000000')
      .setBackground('#' + (catItem.headerBg ? catItem.headerBg.replace(/^FF/, '') : 'C6E0B4'))
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    curCol += skuCount;
  }
  ws.setRowHeight(4, 24);

  // Row 5: Subheaders
  var baseHeaders = [
    'Địa chỉ Store (ship to)',
    'Tên Store (Location)',
    'Mã PG',
    'Họ Tên PG',
    'VTCV',
    'Chỉ tiêu Doanh Số Nhập (VND)',
    'Tổng Doanh Số Nhập Thực Hiện (VND)',
    '% Đạt'
  ];
  for (var b = 0; b < baseHeaders.length; b++) {
    var bCell = ws.getRange(5, b + 1);
    bCell.setValue(baseHeaders[b]).setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    if (b < 5) {
      bCell.setBackground('#FFE699').setFontColor('#000000');
    } else if (b === 5) {
      bCell.setBackground('#F8CBAD').setFontColor('#000000');
    } else if (b === 6) {
      bCell.setBackground('#F4B084').setFontColor('#000000');
    } else {
      bCell.setBackground('#C00000').setFontColor('#FFFFFF');
    }
  }

  for (var sIdx = 0; sIdx < allSkus.length; sIdx++) {
    var sku = allSkus[sIdx];
    var sCol = 9 + sIdx;
    var sCell = ws.getRange(5, sCol);
    sCell.setValue(sku.name + '\n(' + sku.code + ')').setFontSize(8).setFontWeight('bold').setFontColor('#1F4E78')
      .setBackground(sku.headerBg).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  }
  ws.setRowHeight(5, 45);

  // Data Rows (Row 6 onwards)
  var startRow = 6;
  var endRow = startRow + stores.length - 1;

  for (var rIdx = 0; rIdx < stores.length; rIdx++) {
    var rowNum = startRow + rIdx;
    var st = stores[rIdx];
    var rowBg = (rIdx % 2 === 0) ? '#FFFFFF' : '#FAFAFA';

    ws.getRange(rowNum, 1).setValue(st.addr).setHorizontalAlignment('left').setBackground(rowBg);
    ws.getRange(rowNum, 2).setValue(st.location || 'ST00_CVS_FM').setHorizontalAlignment('center').setBackground(rowBg);
    ws.getRange(rowNum, 3).setValue(st.ma_pg || '').setHorizontalAlignment('center').setBackground(rowBg);
    ws.getRange(rowNum, 4).setValue(st.ten_pg || '').setHorizontalAlignment('left').setBackground(rowBg);
    ws.getRange(rowNum, 5).setValue(st.vtcv || 'SR').setHorizontalAlignment('center').setBackground(rowBg);

    // Target Col F
    ws.getRange(rowNum, 6).setValue(st.target || 0).setNumberFormat('#,##0').setFontWeight('bold')
      .setBackground('#FFF9E6').setHorizontalAlignment('right');

    // Actual Formula Col G
    var lastColLetter = getA1NotationColumn(totalCols);
    ws.getRange(rowNum, 7).setFormula('=SUM(I' + rowNum + ':' + lastColLetter + rowNum + ')')
      .setNumberFormat('#,##0').setFontWeight('bold').setBackground('#F2F9F2').setHorizontalAlignment('right');

    // % Dat Formula Col H
    var pctCell = ws.getRange(rowNum, 8);
    pctCell.setFormula('=IF(F' + rowNum + '>0, G' + rowNum + '/F' + rowNum + ', 0)')
      .setNumberFormat('0.0%').setFontWeight('bold').setHorizontalAlignment('center');

    var actualVal = st.actual || 0;
    var tgtVal = st.target || 0;
    var stPct = tgtVal > 0 ? (actualVal / tgtVal * 100) : 0;
    if (stPct >= timePct) {
      pctCell.setBackground('#28E11F').setFontColor('#000000');
    } else if (stPct >= timePct * 0.75) {
      pctCell.setBackground('#F5991F').setFontColor('#FFFFFF');
    } else {
      pctCell.setBackground('#DD2323').setFontColor('#FFFFFF');
    }

    // SKU order values (Col I onwards)
    for (var sk = 0; sk < allSkus.length; sk++) {
      var icode = allSkus[sk].code;
      var amt = (st.orders && st.orders[icode]) ? st.orders[icode] : 0;
      var cell = ws.getRange(rowNum, 9 + sk);
      cell.setValue(amt).setNumberFormat('#,##0').setHorizontalAlignment('right');
      if (amt > 0) {
        cell.setBackground(allSkus[sk].cellBg).setFontColor('#000000').setFontWeight('normal');
      } else {
        cell.setBackground(rowBg).setFontColor('#A0A0A0');
      }
    }
    ws.setRowHeight(rowNum, 20);
  }

  // Row Total (TỔNG CỘNG)
  var totalRow = endRow + 1;
  var totRange = ws.getRange(totalRow, 1, 1, 5);
  totRange.merge();
  totRange.setValue('TỔNG CỘNG (' + stores.length + ' CỬA HÀNG FAMILYMART)')
    .setFontSize(11).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E78')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Target SUBTOTAL
  ws.getRange(totalRow, 6).setFormula('=SUBTOTAL(9, F' + startRow + ':F' + endRow + ')')
    .setNumberFormat('#,##0').setFontSize(11).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E78')
    .setHorizontalAlignment('right');

  // Actual SUBTOTAL
  ws.getRange(totalRow, 7).setFormula('=SUBTOTAL(9, G' + startRow + ':G' + endRow + ')')
    .setNumberFormat('#,##0').setFontSize(11).setFontWeight('bold').setFontColor('#FFFF00').setBackground('#1F4E78')
    .setHorizontalAlignment('right');

  // % Dat Formula
  ws.getRange(totalRow, 8).setFormula('=IF(F' + totalRow + '>0, G' + totalRow + '/F' + totalRow + ', 0)')
    .setNumberFormat('0.0%').setFontSize(12).setFontWeight('bold').setFontColor('#FFFF00').setBackground('#1F4E78')
    .setHorizontalAlignment('center');

  // SKU SUBTOTALs
  for (var skIdx = 0; skIdx < allSkus.length; skIdx++) {
    var cLet = getA1NotationColumn(9 + skIdx);
    ws.getRange(totalRow, 9 + skIdx).setFormula('=SUBTOTAL(9, ' + cLet + startRow + ':' + cLet + endRow + ')')
      .setNumberFormat('#,##0').setFontSize(9).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E78')
      .setHorizontalAlignment('right');
  }
  ws.setRowHeight(totalRow, 26);

  // Borders
  ws.getRange(4, 1, (totalRow - 4 + 1), totalCols).setBorder(true, true, true, true, true, true, '#B4B4B4', SpreadsheetApp.BorderStyle.SOLID);

  // Freeze panes at column I (freeze 8 columns) and 5 rows
  ws.setFrozenRows(5);
  ws.setFrozenColumns(8);
}

// Helper to convert column number to letter
function getA1NotationColumn(columnNumber) {
  var dividend = columnNumber;
  var columnName = '';
  var modulo;
  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return columnName;
}

