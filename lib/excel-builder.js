const ExcelJS = require("exceljs");
const { FM_CATEGORIES, DEFAULT_FM_STORES } = require("./fm-data.js");

function openpyxlColLetter(colIdx) {
      let temp = colIdx;
      let letter = '';
      while (temp > 0) {
        let mod = (temp - 1) % 26;
        letter = String.fromCharCode(65 + mod) + letter;
        temp = Math.floor((temp - mod) / 26);
      }
      return letter;
    }

    async function buildExcelWorkbook(p) {
      const wb = new ExcelJS.Workbook();
        wb.creator = 'Team Lead ' + p.team_lead;
        wb.created = new Date();

        const thinBorder = {
          top: { style: 'thin', color: { argb: 'FFB4B4B4' } },
          left: { style: 'thin', color: { argb: 'FFB4B4B4' } },
          bottom: { style: 'thin', color: { argb: 'FFB4B4B4' } },
          right: { style: 'thin', color: { argb: 'FFB4B4B4' } }
        };

        // -------------------------------------------------------------
        // SHEET 1: DASHBOARD
        // -------------------------------------------------------------
        const ws1 = wb.addWorksheet('Dashboard', { properties: { tabColor: { argb: 'FF1F4E78' } } });
        ws1.columns = [
          { width: 4 }, { width: 10 }, { width: 25 }, { width: 18 },
          { width: 18 }, { width: 18 }, { width: 18 }, { width: 14 }
        ];

        // Banner Title B2:H2
        ws1.mergeCells('B2:H2');
        const b2 = ws1.getCell('B2');
        b2.value = '📊 THEO DÕI DOANH SỐ TEAM - ' + p.team_lead.toUpperCase();
        b2.font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
        b2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        b2.alignment = { horizontal: 'center', vertical: 'middle' };
        ws1.getRow(2).height = 35;

        // Metadata B3:H3
        const setMeta = (c, val, isLbl, isRed) => {
          const cell = ws1.getCell(c);
          cell.value = val;
          cell.font = { name: 'Calibri', bold: true, color: { argb: isRed ? 'FFC00000' : 'FF000000' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isLbl ? 'FFFFF2CC' : 'FFFFF9E6' } };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = thinBorder;
        };
        setMeta('B3', 'Tháng:', true, false);
        setMeta('C3', p.month + '/' + p.year, false, true);
        setMeta('D3', 'Team Lead:', true, false);
        setMeta('E3', p.team_lead, false, false);
        setMeta('F3', '% Timegone:', true, false);
        setMeta('G3', p.timegone / 100, false, true);
        ws1.getCell('G3').numFmt = '0.0%';
        const nowD = new Date();
        setMeta('H3', 'Cập nhật: ' + String(nowD.getDate()).padStart(2,'0') + '/' + String(nowD.getMonth()+1).padStart(2,'0') + '/' + nowD.getFullYear(), false, false);
        ws1.getRow(3).height = 24;

        // Section Title B5:G5
        ws1.mergeCells('B5:G5');
        const b5 = ws1.getCell('B5');
        b5.value = '🎯 TỔNG QUAN DOANH SỐ TEAM';
        b5.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FF1F4E78' } };
        b5.alignment = { horizontal: 'center', vertical: 'middle' };
        ws1.getRow(5).height = 26;

        // Row 6 & 7: Thẻ KPI Doanh Số Thực Tế (B to G)
        const kpiDef = [
          { col: 'B', lbl: 'Thực Hiện BHX', val: p.total_actual_bhx, fmt: '#,##0', bg: 'FF70AD47', fg: 'FFFFFFFF' },
          { col: 'C', lbl: 'Thực Hiện CVS', val: p.total_cvs, fmt: '#,##0', bg: 'FFED7D31', fg: 'FFFFFFFF' },
          { col: 'D', lbl: 'Tổng Thực Hiện', val: p.total_actual, fmt: '#,##0', bg: 'FF2E75B6', fg: 'FFFFFFFF' },
          { col: 'E', lbl: 'Tổng CH BHX Team', val: (p.total_stores_bhx_team || 178) + ' CH', fmt: null, bg: 'FF4472C4', fg: 'FFFFFFFF' },
          { col: 'F', lbl: 'TB Doanh Số/CH BHX', val: p.bhx_per_store, fmt: '#,##0', bg: 'FF70AD47', fg: 'FFFFFFFF' },
          { col: 'G', lbl: 'Số Nhân Viên', val: (p.employees || []).length + ' NV', fmt: null, bg: 'FF595959', fg: 'FFFFFFFF' }
        ];

        kpiDef.forEach(k => {
          const lCell = ws1.getCell(k.col + '6');
          lCell.value = k.lbl;
          lCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: k.fg } };
          lCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: k.bg } };
          lCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          lCell.border = thinBorder;

          const vCell = ws1.getCell(k.col + '7');
          vCell.value = k.val;
          if (k.fmt) vCell.numFmt = k.fmt;
          vCell.font = { name: 'Calibri', size: 12, bold: true };
          vCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
          vCell.alignment = { horizontal: 'center', vertical: 'middle' };
          vCell.border = thinBorder;
        });
        ws1.getRow(6).height = 30;
        ws1.getRow(7).height = 28;

        // Row 9: Status Bar B9:G9
        ws1.mergeCells('B9:G9');
        const b9 = ws1.getCell('B9');
        b9.value = '📊 BẢNG THEO DÕI DOANH SỐ CHI TIẾT - TEAM ' + p.team_lead.toUpperCase() + ' - THÁNG ' + p.month + '/' + p.year;
        b9.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        b9.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        b9.alignment = { horizontal: 'center', vertical: 'middle' };
        ws1.getRow(9).height = 25;

        // Row 11: Header Bảng Xếp Hạng
        ws1.mergeCells('B11:G11');
        const b11 = ws1.getCell('B11');
        b11.value = 'XẾP HẠNG DOANH SỐ NHÂN VIÊN';
        b11.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FF1F4E78' } };
        b11.alignment = { horizontal: 'center', vertical: 'middle' };
        ws1.getRow(11).height = 26;

        // Row 12: Headers B12:G12 (Bỏ Target và % Đạt)
        const rHeaders = ['Hạng', 'Nhân Viên', 'Doanh Số BHX', 'Doanh Số CVS', 'Tổng Thực Hiện', 'Tỷ Trọng DS'];
        const rCols = ['B', 'C', 'D', 'E', 'F', 'G'];
        rHeaders.forEach((h, i) => {
          const c = ws1.getCell(rCols[i] + '12');
          c.value = h;
          c.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.border = thinBorder;
        });
        ws1.getRow(12).height = 26;

        const sortedEmps = (p.employees || []).slice().sort((a, b) => (b.total_actual || 0) - (a.total_actual || 0));
        const rankTotRow = 13 + sortedEmps.length;
        sortedEmps.forEach((e, idx) => {
          const rowNum = 13 + idx;
          const setCell = (col, val, fmt, isBold, align) => {
            const c = ws1.getCell(col + rowNum);
            c.value = val;
            if (fmt) c.numFmt = fmt;
            c.font = { name: 'Calibri', bold: !!isBold };
            c.alignment = { horizontal: align || 'left', vertical: 'middle' };
            c.border = thinBorder;
            return c;
          };
          setCell('B', idx + 1, null, true, 'center');
          setCell('C', e.name, null, true, 'left');
          setCell('D', e.bhx_actual || 0, '#,##0', false, 'right');
          setCell('E', e.cvs_total || 0, '#,##0', false, 'right');
          setCell('F', { formula: `D${rowNum}+E${rowNum}`, result: e.total_actual || 0 }, '#,##0', true, 'right');
          const pctShare = p.total_actual > 0 ? ((e.total_actual || 0) / p.total_actual) : 0;
          setCell('G', { formula: `F${rowNum}/$F$${rankTotRow}`, result: pctShare }, '0.0%', false, 'center');
          ws1.getRow(rowNum).height = 22;
        });

        // Row Total for Rank table
        ws1.mergeCells(`B${rankTotRow}:C${rankTotRow}`);
        const rTotLbl = ws1.getCell(`B${rankTotRow}`);
        rTotLbl.value = 'TỔNG CỘNG TEAM';
        rTotLbl.font = { name: 'Calibri', bold: true, color: { argb: 'FFFFFFFF' } };
        rTotLbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        rTotLbl.alignment = { horizontal: 'center', vertical: 'middle' };
        rTotLbl.border = thinBorder;
        ws1.getCell(`C${rankTotRow}`).border = thinBorder;

        const setTotCell = (col, formula, result, fmt) => {
          const c = ws1.getCell(`${col}${rankTotRow}`);
          c.value = { formula, result };
          c.numFmt = fmt;
          c.font = { name: 'Calibri', bold: true, color: { argb: 'FFFFFFFF' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
          c.alignment = { horizontal: 'right', vertical: 'middle' };
          c.border = thinBorder;
        };
        setTotCell('D', `SUM(D13:D${rankTotRow-1})`, p.total_actual_bhx, '#,##0');
        setTotCell('E', `SUM(E13:E${rankTotRow-1})`, p.total_cvs, '#,##0');
        setTotCell('F', `SUM(F13:F${rankTotRow-1})`, p.total_actual, '#,##0');
        const rTotG = ws1.getCell(`G${rankTotRow}`);
        rTotG.value = { formula: `SUM(G13:G${rankTotRow-1})`, result: 1 };
        rTotG.numFmt = '0.0%';
        rTotG.font = { name: 'Calibri', bold: true, color: { argb: 'FFFFFF00' } };
        rTotG.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        rTotG.alignment = { horizontal: 'center', vertical: 'middle' };
        rTotG.border = thinBorder;
        ws1.getRow(rankTotRow).height = 25;

        // -------------------------------------------------------------
        // SHEET 2: TIẾN ĐỘ TEAM
        // -------------------------------------------------------------
        const ws2 = wb.addWorksheet('Tiến độ Team', { properties: { tabColor: { argb: 'FF70AD47' } } });
        ws2.columns = [
          { width: 3 }, { width: 6 }, { width: 24 }, { width: 16 },
          { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 },
          { width: 14 }, { width: 14 }, { width: 20 }
        ];

        // Banner B2:K2
        ws2.mergeCells('B2:K2');
        const wb2 = ws2.getCell('B2');
        wb2.value = 'THEO DÕI DOANH SỐ - TEAM ' + p.team_lead.toUpperCase() + ' - THÁNG ' + p.month + '/' + p.year;
        wb2.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        wb2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        wb2.alignment = { horizontal: 'center', vertical: 'middle' };
        ws2.getRow(2).height = 30;

        // Info Row B3:K3
        const setInfo = (col, val, isLbl, isRed) => {
          const c = ws2.getCell(col + '3');
          c.value = val;
          c.font = { name: 'Calibri', bold: true, color: { argb: isRed ? 'FFC00000' : 'FF000000' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isLbl ? 'FFFFF2CC' : 'FFFFF9E6' } };
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.border = thinBorder;
        };
        setInfo('B', 'Tháng:', true, false);
        setInfo('C', p.month + '/' + p.year, false, true);
        setInfo('D', 'Tổng CH BHX:', true, false);
        setInfo('E', p.total_stores_bhx_team || p.total_stores_bhx, false, true);
        setInfo('F', 'TB BHX/CH:', true, false);
        setInfo('G', p.bhx_per_store, false, true);
        ws2.getCell('G3').numFmt = '#,##0';
        setInfo('H', 'Cập nhật:', true, false);
        setInfo('I', String(nowD.getDate()).padStart(2,'0') + '/' + String(nowD.getMonth()+1).padStart(2,'0') + '/' + nowD.getFullYear(), false, false);
        setInfo('J', '% Timegone:', true, false);
        setInfo('K', p.timegone / 100, false, true);
        ws2.getCell('K3').numFmt = '0.0%';
        ws2.getRow(3).height = 24;

        // Group headers
        ws2.mergeCells('D5:J5');
        const g1 = ws2.getCell('D5');
        g1.value = 'THỰC HIỆN (chi tiết theo từng kênh)';
        g1.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        g1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
        g1.alignment = { horizontal: 'center', vertical: 'middle' };

        const g2 = ws2.getCell('K5');
        g2.value = 'TỔNG CỘNG';
        g2.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        g2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        g2.alignment = { horizontal: 'center', vertical: 'middle' };
        ws2.getRow(5).height = 22;

        // Subheaders Row 6 (Bỏ Target và % Đạt)
        const teamHeaders = [
          'STT', 'Nhân Viên', 'BHX\n(cố định)', 'GS25', '7-Eleven',
          'FamilyMart', 'Circle K', 'Hoàng Đức', 'WinMart+', 'Tổng TH\n(BHX+CVS)'
        ];
        const teamCols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];
        teamHeaders.forEach((th, idx) => {
          const c = ws2.getCell(teamCols[idx] + '6');
          c.value = th;
          c.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
          c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          c.border = thinBorder;
        });
        ws2.getRow(6).height = 40;

        // 10 Employee rows
        p.employees.forEach((emp, k) => {
          const r = 7 + k;
          const setC = (col, val, fmt, isBold, align) => {
            const cell = ws2.getCell(col + r);
            cell.value = val;
            if (fmt) cell.numFmt = fmt;
            cell.font = { name: 'Calibri', bold: !!isBold };
            cell.alignment = { horizontal: align || 'left', vertical: 'middle' };
            cell.border = thinBorder;
            return cell;
          };

          setC('B', k + 1, null, false, 'center');
          setC('C', emp.name, null, true, 'left');
          setC('D', emp.bhx_actual, '#,##0', false, 'right');
          setC('E', emp.gs25_actual, '#,##0', false, 'right');
          setC('F', emp.se_actual, '#,##0', false, 'right');
          setC('G', emp.fm_actual, '#,##0', false, 'right');
          setC('H', emp.ck_actual, '#,##0', false, 'right');
          setC('I', emp.hd_actual, '#,##0', false, 'right');
          setC('J', emp.wmp_actual, '#,##0', false, 'right');

          // Tổng TH formula =SUM(D..:J..)
          setC('K', { formula: `SUM(D${r}:J${r})`, result: emp.total_actual }, '#,##0', true, 'right');
          ws2.getRow(r).height = 22;
        });

        // Row Total TỔNG TEAM
        const totalR = 7 + p.employees.length;
        ws2.mergeCells(`B${totalR}:C${totalR}`);
        const totLbl = ws2.getCell(`B${totalR}`);
        totLbl.value = 'TỔNG TEAM';
        totLbl.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        totLbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        totLbl.alignment = { horizontal: 'center', vertical: 'middle' };
        totLbl.border = thinBorder;
        ws2.getCell(`C${totalR}`).border = thinBorder;

        const sumCols = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];
        const sumResults = [
          p.total_actual_bhx,
          p.employees.reduce((a,x)=>a+x.gs25_actual,0),
          p.employees.reduce((a,x)=>a+x.se_actual,0),
          p.employees.reduce((a,x)=>a+x.fm_actual,0),
          p.employees.reduce((a,x)=>a+x.ck_actual,0),
          p.employees.reduce((a,x)=>a+x.hd_actual,0),
          p.employees.reduce((a,x)=>a+x.wmp_actual,0),
          p.total_actual
        ];
        sumCols.forEach((col, idx) => {
          const sc = ws2.getCell(`${col}${totalR}`);
          sc.value = { formula: `SUM(${col}7:${col}${totalR-1})`, result: sumResults[idx] };
          sc.numFmt = '#,##0';
          sc.font = { name: 'Calibri', bold: true, color: { argb: 'FFFFFFFF' } };
          sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
          sc.alignment = { horizontal: 'right', vertical: 'middle' };
          sc.border = thinBorder;
        });
        ws2.getRow(totalR).height = 28;

        // Freeze Panes at D7
        ws2.views = [{ state: 'frozen', xSplit: 3, ySplit: 6 }];

        // -------------------------------------------------------------
        // SHEET 3: CHI TIẾT CVS
        // -------------------------------------------------------------
        const ws3 = wb.addWorksheet('Chi tiết CVS', { properties: { tabColor: { argb: 'FFED7D31' } } });
        ws3.columns = [
          { width: 6 }, { width: 22 }, { width: 14 }, { width: 16 }, { width: 45 }, { width: 18 }
        ];

        ws3.mergeCells('A1:F1');
        const t3 = ws3.getCell('A1');
        t3.value = 'CHI TIẾT DOANH SỐ CVS - TỪNG CỬA HÀNG';
        t3.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        t3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        t3.alignment = { horizontal: 'center', vertical: 'middle' };
        ws3.getRow(1).height = 30;

        const cvsHeaders = ['STT', 'Nhân Viên', 'Chuỗi', 'Mã CH', 'Địa Chỉ', 'Thực Hiện'];
        cvsHeaders.forEach((ch, idx) => {
          const colLetter = String.fromCharCode(65 + idx);
          const c = ws3.getCell(colLetter + '3');
          c.value = ch;
          c.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.border = thinBorder;
        });
        ws3.getRow(3).height = 25;

        let curR3 = 4;
        let sttCounter = 1;
        const chainColors = {
          'GS25': 'FF4472C4', '7-Eleven': 'FF00B050', '7E': 'FF00B050', 'FamilyMart': 'FFED7D31',
          'Circle K': 'FFC00000', 'Hoàng Đức': 'FF70AD47', 'Hoàng Đức Long Khánh': 'FF70AD47',
          'Hoàng Đức Gia Kiệm': 'FF008080', 'WinMart+': 'FFD9534F', 'WMP': 'FFD9534F'
        };

        p.employees.forEach(emp => {
          const empStores = p.stores.filter(s => s.employee_name === emp.name);
          if (empStores.length === 0) return;

          // Subheader row
          ws3.mergeCells(`A${curR3}:D${curR3}`);
          const sHead = ws3.getCell(`A${curR3}`);
          sHead.value = '👤 ' + emp.name + ' (' + empStores.length + ' cửa hàng CVS)';
          sHead.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
          sHead.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
          sHead.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
          sHead.border = thinBorder;

          const sLbl = ws3.getCell(`E${curR3}`);
          sLbl.value = 'TỔNG CVS:';
          sLbl.font = { name: 'Calibri', bold: true, color: { argb: 'FFFFFFFF' } };
          sLbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
          sLbl.alignment = { horizontal: 'right', vertical: 'middle' };
          sLbl.border = thinBorder;

          const sTot = ws3.getCell(`F${curR3}`);
          const empTotVal = empStores.reduce((a,x)=>a+(x.actual||0), 0);
          sTot.value = empTotVal;
          sTot.numFmt = '#,##0';
          sTot.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFF00' } };
          sTot.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
          sTot.alignment = { horizontal: 'right', vertical: 'middle' };
          sTot.border = thinBorder;
          ws3.getRow(curR3).height = 24;
          curR3++;

          empStores.forEach(st => {
            const setS = (col, val, fmt, isBold, align) => {
              const c = ws3.getCell(col + curR3);
              c.value = val;
              if (fmt) c.numFmt = fmt;
              c.font = { name: 'Calibri', size: 10, bold: !!isBold };
              c.alignment = { horizontal: align || 'left', vertical: 'middle' };
              c.border = thinBorder;
              return c;
            };

            setS('A', sttCounter, null, false, 'center');
            setS('B', emp.name, null, false, 'left');

            const chCell = setS('C', st.chain, null, true, 'center');
            chCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: chainColors[st.chain] || 'FF70AD47' } };
            chCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };

            setS('D', st.store_code || '', null, false, 'left');
            setS('E', st.store_address || '', null, false, 'left');
            setS('F', st.actual || 0, '#,##0', false, 'right');

            ws3.getRow(curR3).height = 20;
            curR3++;
            sttCounter++;
          });
        });

        // Grand Total Row
        ws3.mergeCells(`A${curR3}:E${curR3}`);
        const gtLbl = ws3.getCell(`A${curR3}`);
        gtLbl.value = '🎯 TỔNG CVS TOÀN TEAM';
        gtLbl.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        gtLbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        gtLbl.alignment = { horizontal: 'center', vertical: 'middle' };
        gtLbl.border = thinBorder;

        const gtVal = ws3.getCell(`F${curR3}`);
        gtVal.value = p.total_cvs;
        gtVal.numFmt = '#,##0';
        gtVal.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFF00' } };
        gtVal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        gtVal.alignment = { horizontal: 'right', vertical: 'middle' };
        gtVal.border = thinBorder;
        ws3.getRow(curR3).height = 28;

        ws3.views = [{ state: 'frozen', ySplit: 3 }];

        // -------------------------------------------------------------
        // SHEET 4: DATA BHX
        // -------------------------------------------------------------
        const ws4 = wb.addWorksheet('Data BHX', { properties: { tabColor: { argb: 'FF4472C4' } } });
        ws4.columns = [
          { width: 6 }, { width: 24 }, { width: 14 }, { width: 16 }, { width: 18 }, { width: 14 }
        ];

        ws4.mergeCells('A1:F1');
        const t4 = ws4.getCell('A1');
        t4.value = 'PHÂN BỔ DOANH SỐ BHX (' + p.total_stores_bhx + ' CỬA HÀNG)';
        t4.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
        t4.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        t4.alignment = { horizontal: 'center', vertical: 'middle' };
        ws4.getRow(1).height = 28;

        const setBhxTop = (row, lbl, val, fmt, isRed) => {
          const a = ws4.getCell('A' + row);
          a.value = lbl;
          a.font = { name: 'Calibri', bold: true };
          a.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
          a.border = thinBorder;

          const b = ws4.getCell('B' + row);
          b.value = val;
          if (fmt) b.numFmt = fmt;
          b.font = { name: 'Calibri', bold: true, color: { argb: isRed ? 'FFC00000' : 'FF000000' } };
          b.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9E6' } };
          b.border = thinBorder;
        };
        setBhxTop(3, 'Tổng DS 5 Hubs Hệ Thống:', p.total_actual_bhx_hubs || p.total_actual_bhx, '#,##0', true);
        setBhxTop(4, 'Tổng số CH hệ thống:', p.total_stores_bhx, null, false);
        setBhxTop(5, 'TB doanh số/CH:', { formula: 'B3/B4', result: p.bhx_per_store }, '#,##0', true);

        const bhxHeaders = ['STT', 'Nhân Viên', 'Số CH BHX', 'TB/CH', 'Doanh Số BHX', '% BHX/Team'];
        bhxHeaders.forEach((bh, idx) => {
          const colLetter = String.fromCharCode(65 + idx);
          const c = ws4.getCell(colLetter + '7');
          c.value = bh;
          c.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.border = thinBorder;
        });
        ws4.getRow(7).height = 25;

        const bhxTotRow = 8 + p.employees.length;

        p.employees.forEach((emp, idx) => {
          const r = 8 + idx;
          const setB = (col, val, fmt, isBold, align) => {
            const cell = ws4.getCell(col + r);
            cell.value = val;
            if (fmt) cell.numFmt = fmt;
            cell.font = { name: 'Calibri', bold: !!isBold };
            cell.alignment = { horizontal: align || 'left', vertical: 'middle' };
            cell.border = thinBorder;
            return cell;
          };

          setB('A', idx + 1, null, false, 'center');
          setB('B', emp.name, null, true, 'left');
          setB('C', emp.bhx_stores, null, false, 'center');
          setB('D', { formula: '$B$5', result: p.bhx_per_store }, '#,##0', false, 'right');
          setB('E', { formula: `C${r}*D${r}`, result: emp.bhx_actual }, '#,##0', true, 'right');
          setB('F', { formula: `E${r}/$E$${bhxTotRow}`, result: p.total_actual_bhx > 0 ? (emp.bhx_actual / p.total_actual_bhx) : 0 }, '0.0%', false, 'center');
          ws4.getRow(r).height = 20;
        });

        ws4.mergeCells(`A${bhxTotRow}:B${bhxTotRow}`);
        const bTotLbl = ws4.getCell(`A${bhxTotRow}`);
        bTotLbl.value = 'TỔNG CỘNG TEAM';
        bTotLbl.font = { name: 'Calibri', bold: true };
        bTotLbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
        bTotLbl.alignment = { horizontal: 'center', vertical: 'middle' };
        bTotLbl.border = thinBorder;

        const bTotC = ws4.getCell(`C${bhxTotRow}`);
        bTotC.value = { formula: `SUM(C8:C${bhxTotRow-1})`, result: p.employees.reduce((a,x)=>a+x.bhx_stores,0) };
        bTotC.font = { name: 'Calibri', bold: true };
        bTotC.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
        bTotC.alignment = { horizontal: 'center', vertical: 'middle' };
        bTotC.border = thinBorder;

        const bTotD = ws4.getCell(`D${bhxTotRow}`);
        bTotD.value = { formula: '$B$5', result: p.bhx_per_store };
        bTotD.numFmt = '#,##0';
        bTotD.font = { name: 'Calibri', bold: true };
        bTotD.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
        bTotD.alignment = { horizontal: 'right', vertical: 'middle' };
        bTotD.border = thinBorder;

        const bTotE = ws4.getCell(`E${bhxTotRow}`);
        bTotE.value = { formula: `SUM(E8:E${bhxTotRow-1})`, result: p.total_actual_bhx };
        bTotE.numFmt = '#,##0';
        bTotE.font = { name: 'Calibri', bold: true };
        bTotE.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
        bTotE.alignment = { horizontal: 'right', vertical: 'middle' };
        bTotE.border = thinBorder;

        const bTotF = ws4.getCell(`F${bhxTotRow}`);
        bTotF.value = { formula: `SUM(F8:F${bhxTotRow-1})`, result: 1 };
        bTotF.numFmt = '0.0%';
        bTotF.font = { name: 'Calibri', bold: true };
        bTotF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
        bTotF.alignment = { horizontal: 'center', vertical: 'middle' };
        bTotF.border = thinBorder;
        ws4.getRow(bhxTotRow).height = 24;


        // -------------------------------------------------------------
        // SHEET 5: CONFIG
        // -------------------------------------------------------------
        const ws5 = wb.addWorksheet('Config', { properties: { tabColor: { argb: 'FF595959' } } });
        ws5.columns = [{ width: 35 }, { width: 35 }];

        ws5.mergeCells('A1:B1');
        const t5 = ws5.getCell('A1');
        t5.value = 'THÔNG SỐ CẤU HÌNH HỆ THỐNG';
        t5.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        t5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        t5.alignment = { horizontal: 'center', vertical: 'middle' };
        ws5.getRow(1).height = 28;

        const nowStr = new Date().toLocaleDateString('vi-VN') + ' ' + new Date().toLocaleTimeString('vi-VN');
        const cfgItems = [
          ['Team Lead', p.team_lead],
          ['Tháng Báo Cáo', p.month],
          ['Năm Báo Cáo', p.year],
          ['% Timegone Thực Tế', p.timegone + '%'],
          ['Tổng Cửa Hàng BHX Hệ Thống', p.total_stores_bhx || 232],
          ['Tổng Cửa Hàng GS25 Hệ Thống', 72],
          ['Hệ Số Phân Bổ GS25 (2 DC)', '20.46%'],
          ['Tổng Cửa Hàng 7-Eleven Phụ Trách', 5],
          ['Hệ Số Phân Bổ 7-Eleven (2 DC)', '3.62%'],
          ['Thời Gian Xuất Báo Cáo', nowStr]
        ];

        cfgItems.forEach((item, idx) => {
          const r = 3 + idx;
          const ca = ws5.getCell('A' + r);
          ca.value = item[0];
          ca.font = { name: 'Calibri', bold: true };
          ca.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
          ca.border = thinBorder;

          const cb = ws5.getCell('B' + r);
          cb.value = item[1];
          cb.alignment = { horizontal: 'center', vertical: 'middle' };
          cb.border = thinBorder;
          ws5.getRow(r).height = 22;
        });

        // -------------------------------------------------------------
        // SHEET 6: CHI TIẾT SKU FAMILYMART
        // -------------------------------------------------------------
        const ws6 = wb.addWorksheet('Chi tiết SKU FamilyMart', { properties: { tabColor: { argb: 'FFED7D31' } } });

        const fmData = p.fm_sku_matrix;
        const fmStores = (fmData && fmData.stores && fmData.stores.length > 0) ? fmData.stores : DEFAULT_FM_STORES;
        const fmCats = (fmData && fmData.categories && fmData.categories.length > 0) ? fmData.categories : FM_CATEGORIES;

        // Flatten all SKUs
        const allFmSkus = [];
        fmCats.forEach(cat => {
          cat.skus.forEach(s => {
            allFmSkus.push({
              code: s.code,
              name: s.name,
              catName: cat.name,
              headerBg: cat.headerBg || 'FFC6E0B4',
              cellBg: cat.cellBg || 'FFE2EFDA'
            });
          });
        });

        // Set column widths
        const colWidths = [
          { width: 42 }, // A: Dia chi
          { width: 16 }, // B: Location
          { width: 13 }, // C: Ma PG
          { width: 22 }, // D: Ho ten PG
          { width: 8 },  // E: VTCV
          { width: 22 }, // F: Actual
        ];
        allFmSkus.forEach(() => colWidths.push({ width: 16 }));
        ws6.columns = colWidths;

        // Row 1: Title Banner A1:L1
        ws6.mergeCells('A1:L1');
        const t6 = ws6.getCell('A1');
        t6.value = '🛒 BÁO CÁO CHI TIẾT NHẬP HÀNG THEO CỬA HÀNG & SKU - CHUỖI FAMILYMART';
        t6.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
        t6.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        t6.alignment = { horizontal: 'center', vertical: 'middle' };
        ws6.getRow(1).height = 32;

        // Rows 2 & 3: 5 KPI Cards
        const totFmAct = fmData ? (fmData.total_actual || 0) : fmStores.reduce((a, b) => a + (b.actual || 0), 0);
        const storesWithOrder = fmStores.filter(s => (s.actual || 0) > 0).length;
        const pctStoresWithOrder = fmStores.length > 0 ? Math.round(storesWithOrder / fmStores.length * 100) : 0;
        const avgPerFmStore = fmStores.length > 0 ? Math.round(totFmAct / fmStores.length) : 0;

        const fmKpiCols = [
          { col: 'A', lbl: 'Số Cửa Hàng', val: fmStores.length + ' CH', bg: 'FF1F4E78', fg: 'FFFFFFFF' },
          { col: 'B', lbl: 'Tổng Thực Hiện Nhập', val: totFmAct, fmt: '#,##0 "đ"', bg: 'FF70AD47', fg: 'FFFFFFFF' },
          { col: 'C', lbl: 'TB Doanh Số/CH', val: avgPerFmStore, fmt: '#,##0 "đ"', bg: 'FF2E75B6', fg: 'FFFFFFFF' },
          { col: 'D', lbl: 'CH Có Phát Sinh Đơn', val: storesWithOrder + '/' + fmStores.length + ' CH', bg: 'FFED7D31', fg: 'FFFFFFFF' },
          { col: 'E', lbl: 'Tỷ Lệ CH Có Đơn', val: pctStoresWithOrder / 100, fmt: '0.0%', bg: 'FF4472C4', fg: 'FFFFFFFF' }
        ];

        fmKpiCols.forEach(k => {
          const lCell = ws6.getCell(k.col + '2');
          lCell.value = k.lbl;
          lCell.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF555555' } };
          lCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
          lCell.alignment = { horizontal: 'center', vertical: 'middle' };
          lCell.border = thinBorder;

          const vCell = ws6.getCell(k.col + '3');
          vCell.value = k.val;
          if (k.fmt) vCell.numFmt = k.fmt;
          vCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: k.fg } };
          vCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: k.bg } };
          vCell.alignment = { horizontal: 'center', vertical: 'middle' };
          vCell.border = thinBorder;
        });
        ws6.getRow(2).height = 18;
        ws6.getRow(3).height = 24;

        // Row 4: Group Headers
        ws6.mergeCells('A4:E4');
        const g4_1 = ws6.getCell('A4');
        g4_1.value = 'THÔNG TIN ĐIỂM BÁN (STORE MASTER)';
        g4_1.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF1F4E78' } };
        g4_1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        g4_1.alignment = { horizontal: 'center', vertical: 'middle' };

        const g4_2 = ws6.getCell('F4');
        g4_2.value = 'TỔNG THỰC HIỆN';
        g4_2.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        g4_2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
        g4_2.alignment = { horizontal: 'center', vertical: 'middle' };

        let curColIdx = 7;
        fmCats.forEach(cat => {
          const startCol = curColIdx;
          const endCol = curColIdx + cat.skus.length - 1;
          const sLet = openpyxlColLetter(startCol);
          const eLet = openpyxlColLetter(endCol);

          if (cat.skus.length > 1) {
            ws6.mergeCells(sLet + '4:' + eLet + '4');
          }
          const catCell = ws6.getCell(sLet + '4');
          catCell.value = cat.name;
          catCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF000000' } };
          catCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cat.headerBg } };
          catCell.alignment = { horizontal: 'center', vertical: 'middle' };

          curColIdx = endCol + 1;
        });
        ws6.getRow(4).height = 24;

        // Row 5: Subheaders
        const baseSubheaders = [
          { name: 'Địa chỉ Store (ship to)', bg: 'FFFFE699', fg: 'FF000000' },
          { name: 'Tên Store (Location)', bg: 'FFFFE699', fg: 'FF000000' },
          { name: 'Mã PG', bg: 'FFFFE699', fg: 'FF000000' },
          { name: 'Họ Tên PG', bg: 'FFFFE699', fg: 'FF000000' },
          { name: 'VTCV', bg: 'FFFFE699', fg: 'FF000000' },
          { name: 'Tổng Doanh Số Nhập Thực Hiện (VND)', bg: 'FFF4B084', fg: 'FF000000' }
        ];

        baseSubheaders.forEach((sh, idx) => {
          const colLet = openpyxlColLetter(idx + 1);
          const c = ws6.getCell(colLet + '5');
          c.value = sh.name;
          c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: sh.fg } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: sh.bg } };
          c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          c.border = thinBorder;
        });

        allFmSkus.forEach((sku, idx) => {
          const colLet = openpyxlColLetter(7 + idx);
          const c = ws6.getCell(colLet + '5');
          c.value = sku.name + '\n(' + sku.code + ')';
          c.font = { name: 'Calibri', size: 8, bold: true, color: { argb: 'FF1F4E78' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: sku.headerBg } };
          c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          c.border = thinBorder;
        });
        ws6.getRow(5).height = 45;

        // Data Rows (Row 6 to 6 + fmStores.length - 1)
        const startR6 = 6;
        const endR6 = startR6 + fmStores.length - 1;
        const lastColLetter = openpyxlColLetter(6 + allFmSkus.length);

        fmStores.forEach((st, idx) => {
          const r = startR6 + idx;
          const rowBg = (idx % 2 === 0) ? 'FFFFFFFF' : 'FFFAFAFA';

          const setFm = (col, val, fmt, isBold, align, fillBg, fontColor) => {
            const cell = ws6.getCell(col + r);
            cell.value = val;
            if (fmt) cell.numFmt = fmt;
            cell.font = { name: 'Calibri', size: 9, bold: !!isBold, color: { argb: fontColor || 'FF000000' } };
            cell.alignment = { horizontal: align || 'left', vertical: 'middle' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillBg || rowBg } };
            cell.border = thinBorder;
            return cell;
          };

          setFm('A', st.addr, null, false, 'left');
          setFm('B', st.location || 'ST00_CVS_FM', null, false, 'center');
          setFm('C', st.ma_pg || '', null, false, 'center');
          setFm('D', st.ten_pg || '', null, false, 'left');
          setFm('E', st.vtcv || 'SR', null, false, 'center');

          // Actual Col F = SUM(G{r}:lastCol{r})
          setFm('F', { formula: 'SUM(G' + r + ':' + lastColLetter + r + ')', result: st.actual || 0 }, '#,##0', true, 'right', 'FFF2F9F2');

          // 46 SKU columns
          allFmSkus.forEach((sku, sIdx) => {
            const cLet = openpyxlColLetter(7 + sIdx);
            const amt = (st.orders && st.orders[sku.code]) ? st.orders[sku.code] : 0;
            const cBg = amt > 0 ? sku.cellBg : rowBg;
            const cFg = amt > 0 ? 'FF000000' : 'FFA0A0A0';
            setFm(cLet, amt, '#,##0', false, 'right', cBg, cFg);
          });

          ws6.getRow(r).height = 20;
        });

        // Row Total (TỔNG CỘNG)
        const totalR6 = endR6 + 1;
        ws6.mergeCells('A' + totalR6 + ':E' + totalR6);
        const totFmLbl = ws6.getCell('A' + totalR6);
        totFmLbl.value = 'TỔNG CỘNG (' + fmStores.length + ' CỬA HÀNG FAMILYMART)';
        totFmLbl.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        totFmLbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        totFmLbl.alignment = { horizontal: 'center', vertical: 'middle' };
        for (let colIdx = 1; colIdx <= 5; colIdx++) {
          ws6.getCell(openpyxlColLetter(colIdx) + totalR6).border = thinBorder;
        }

        // Col F Actual SUBTOTAL
        const totFmActCell = ws6.getCell('F' + totalR6);
        totFmActCell.value = { formula: 'SUBTOTAL(9, F' + startR6 + ':F' + endR6 + ')', result: totFmAct };
        totFmActCell.numFmt = '#,##0';
        totFmActCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFF00' } };
        totFmActCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        totFmActCell.alignment = { horizontal: 'right', vertical: 'middle' };
        totFmActCell.border = thinBorder;

        // SKU SUBTOTALs
        allFmSkus.forEach((sku, sIdx) => {
          const cLet = openpyxlColLetter(7 + sIdx);
          const sc = ws6.getCell(cLet + totalR6);
          const skuSum = fmStores.reduce((acc, st) => acc + ((st.orders && st.orders[sku.code]) ? st.orders[sku.code] : 0), 0);
          sc.value = { formula: 'SUBTOTAL(9, ' + cLet + startR6 + ':' + cLet + endR6 + ')', result: skuSum };
          sc.numFmt = '#,##0';
          sc.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
          sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
          sc.alignment = { horizontal: 'right', vertical: 'middle' };
          sc.border = thinBorder;
        });
        ws6.getRow(totalR6).height = 26;

        // Freeze Panes at column G (freeze 6 columns, 5 rows)
        ws6.views = [{ state: 'frozen', xSplit: 6, ySplit: 5 }];

        return wb;
    }

module.exports = { buildExcelWorkbook, openpyxlColLetter };
