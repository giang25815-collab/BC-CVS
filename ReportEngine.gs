/**
 * ============================================================================
 * REPORT ENGINE: BỘ MÁY XỬ LÝ & BÓC TÁCH DỮ LIỆU BÁO CÁO DOANH SỐ CVS & BHX
 * Chạy trên nền Google Apps Script Backend (Serverless)
 * ============================================================================
 */

/**
 * Chuẩn hóa chuỗi văn bản (bỏ dấu cách, dấu phẩy, chuyển chữ thường) để so khớp địa chỉ
 */
function normalizeText(text) {
  if (!text) return '';
  return String(text).toLowerCase().replace(/[,.\s\-_/()]/g, '');
}

/**
 * Lấy thời gian hiện tại chuẩn múi giờ Việt Nam (Asia/Ho_Chi_Minh - GMT+7)
 */
function getVietnamCurrentDate() {
  var now = new Date();
  var utc = now.getTime() + (now.getTimezoneOffset() * 60000);
  return new Date(utc + (7 * 3600000));
}

/**
 * Tính tỷ lệ thời gian đã qua trong tháng (% Timegone) theo đúng thời gian mỗi lần gửi file báo cáo
 */
function calculateRealtimeTimegone(month, year, refDate) {
  var vnTime = refDate ? new Date(refDate) : getVietnamCurrentDate();
  var m = month || (vnTime.getMonth() + 1);
  var y = year || vnTime.getFullYear();
  var daysInMonth = new Date(y, m, 0).getDate();
  var day = vnTime.getDate();

  if (y === vnTime.getFullYear() && m === (vnTime.getMonth() + 1)) {
    day = Math.min(vnTime.getDate(), daysInMonth);
  } else if (y < vnTime.getFullYear() || (y === vnTime.getFullYear() && m < (vnTime.getMonth() + 1))) {
    day = daysInMonth;
  } else {
    day = 0;
  }
  return parseFloat(((day / daysInMonth) * 100).toFixed(1));
}


/**
 * Đọc Workbook từ Blob hoặc Base64 bằng thư viện SheetJS (XLSX)
 * Đã tối ưu cellDates, cellStyles, cellFormula để đọc cực nhanh file lớn (.xlsb, .xlsx)
 */
function parseWorkbook(input) {
  if (!input) throw new Error("Dữ liệu file trống.");
  if (typeof XLSX === 'undefined') {
    throw new Error("Thư viện XLSX chưa được khởi tạo trong dự án Google Apps Script.");
  }

  var options = {
    cellDates: false,
    cellStyles: false,
    cellHTML: false,
    cellFormula: false
  };

  // Nếu là Google Apps Script Blob
  if (input.getBytes && typeof input.getBytes === 'function') {
    var b64 = Utilities.base64Encode(input.getBytes());
    options.type = 'base64';
    return XLSX.read(b64, options);
  }

  // Nếu là chuỗi Base64
  if (typeof input === 'string') {
    options.type = 'base64';
    return XLSX.read(input, options);
  }

  // Nếu là byte array hoặc Uint8Array
  options.type = 'array';
  return XLSX.read(input, options);
}

/**
 * Nhận diện loại file: 'ST' (Siêu Thị / BHX + WinMart) hoặc 'CVS' (CVS & NPP)
 */
function detectFileType(fileName, wb) {
  var fn = (fileName || '').toUpperCase();

  // Quy tắc 1: Kiểm tra theo tên file
  if (fn.includes('_ST_') || fn.includes('ST_KD6') || fn.includes('SIEU THI') || fn.includes('SIEUTHI') || fn.includes('BHX')) {
    return 'ST';
  }
  if (fn.includes('KD6') || fn.includes('KD06') || fn.includes('CVS') || fn.includes('NPP') || fn.includes('FAMILY')) {
    return 'CVS';
  }

  // Quy tắc 2: Kiểm tra cấu trúc sheet và nội dung bên trong
  if (wb && wb.SheetNames) {
    if (wb.SheetNames.indexOf('NPP') >= 0) return 'CVS';

    var sName = wb.SheetNames.indexOf('SO') >= 0 ? 'SO' : (wb.SheetNames.indexOf('PO') >= 0 ? 'PO' : wb.SheetNames[0]);
    var ws = wb.Sheets[sName];
    if (ws) {
      var rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      var hasST = false, hasCVS = false;
      var limit = Math.min(50, rows.length);
      for (var i = 0; i < limit; i++) {
        var rStr = JSON.stringify(rows[i] || '');
        if (rStr.includes('XH4001') || rStr.includes('Châu Pha') || rStr.includes('WMP_')) hasST = true;
        if (rStr.includes('GS0003') || rStr.includes('SS4001') || rStr.includes('VT4050') || rStr.includes('HD3024')) hasCVS = true;
      }
      if (hasST && !hasCVS) return 'ST';
      if (hasCVS && !hasST) return 'CVS';
    }
  }

  return 'UNKNOWN';
}

/**
 * Trích xuất và tổng hợp toàn bộ báo cáo từ 2 Workbook: wbST (Siêu Thị) và wbCVS (CVS & NPP)
 * @param {Object} wbST Workbook Siêu Thị (BHX & WinMart+)
 * @param {Object} wbCVS Workbook CVS & NPP (GS25, 7E, Circle K, FamilyMart, Hoàng Đức)
 * @param {Object} customMaster Cấu hình master data (nếu có, mặc định lấy DEFAULT_MASTER)
 * @returns {Object} Payload dữ liệu hoàn chỉnh để gọi createGoogleSheetReport()
 */
function processTwoWorkbooks(wbST, wbCVS, customMaster) {
  var master = customMaster || (typeof DEFAULT_MASTER !== 'undefined' ? DEFAULT_MASTER : MASTER_DATA);
  if (!master) throw new Error("Không tìm thấy Master Data.");

  var vnNow = getVietnamCurrentDate();
  var month = master.month || (vnNow.getMonth() + 1);
  var year = master.year || vnNow.getFullYear();
  var timePct = calculateRealtimeTimegone(month, year, vnNow);

  // 1. TÍNH TOÁN BHX & WINMART+ (wbST)
  var bhxHubAmounts = {};
  master.bhx_hubs.forEach(function(h) { bhxHubAmounts[h] = 0; });
  var totalActualBhx = 0;

  var wmpAmounts = {
    'WMP_DNI_KP_TRUNG_TAM_XUAN_LAP': 0,
    'WMP_DNI_285_287_CACH_MANG_THAN': 0,
    'WMP_HCM_90A_92_PHAN_CHU_TRINH': 0,
    'WMP_DNI_A_01_04_TOPAZ_TWINS': 0
  };

  var sheetST = wbST.SheetNames.indexOf('SO') >= 0 ? 'SO' : (wbST.SheetNames.indexOf('PO') >= 0 ? 'PO' : wbST.SheetNames[0]);
  var rowsST = XLSX.utils.sheet_to_json(wbST.Sheets[sheetST], { header: 1 });

  if (rowsST.length > 0) {
    var headerST = rowsST[0].map(function(h) { return String(h || '').trim(); });
    var idxCustST = headerST.indexOf('CUST_NO');
    var idxAddrST = headerST.indexOf('Address');
    var idxAmtST = headerST.indexOf('SumOfAMOUNT');

    for (var i = 1; i < rowsST.length; i++) {
      var rST = rowsST[i];
      if (!rST) continue;
      var custST = String(rST[idxCustST] || '');
      var addrST = String(rST[idxAddrST] || '').trim();
      var amtST = parseFloat(rST[idxAmtST]) || 0;

      // Hub BHX: Mã XH4001
      if (custST === 'XH4001') {
        var normAST = normalizeText(addrST);
        for (var hIdx = 0; hIdx < master.bhx_hubs.length; hIdx++) {
          var h = master.bhx_hubs[hIdx];
          if (addrST === h || normAST === normalizeText(h)) {
            bhxHubAmounts[h] += amtST;
            totalActualBhx += amtST;
            break;
          }
        }
      }

      // WinMart+: Kiểm tra ở cột 7 (vị trí hoặc địa chỉ)
      var locVal = String(rST[7] || '').trim();
      for (var wk in wmpAmounts) {
        if (locVal.toLowerCase().indexOf(wk.toLowerCase()) >= 0 || (wk === 'WMP_DNI_285_287_CACH_MANG_THAN' && locVal.indexOf('285_287') >= 0)) {
          wmpAmounts[wk] += amtST;
        }
      }
    }
  }

  var totalStoresBhx = master.total_stores_bhx || 232;
  var bhxPerStore = totalStoresBhx > 0 ? (totalActualBhx / totalStoresBhx) : 0;

  // 2. TÍNH TOÁN CVS & NPP (wbCVS)
  var gs25Amounts = {};
  master.gs25_dcs.forEach(function(d) { gs25Amounts[d] = 0; });
  var totalGs25Dcs = 0;

  var seAmounts = {};
  master.se_dcs.forEach(function(d) { seAmounts[d] = 0; });
  var totalSeDcs = 0;

  var ckKhoKhoAmt = 0;
  var ckStoresSo = {};
  var hoangDucStores = { '198 Hùng Vương': 0, 'Bạch Lâm': 0 };

  var sheetCVS = wbCVS.SheetNames.indexOf('SO') >= 0 ? 'SO' : (wbCVS.SheetNames.indexOf('PO') >= 0 ? 'PO' : wbCVS.SheetNames[0]);
  var rowsCVS = XLSX.utils.sheet_to_json(wbCVS.Sheets[sheetCVS], { header: 1 });

  if (rowsCVS.length > 0) {
    var headerCVS = rowsCVS[0].map(function(h) { return String(h || '').trim(); });
    var idxCustCVS = headerCVS.indexOf('CUST_NO');
    var idxAddrCVS = headerCVS.indexOf('Address');
    var idxAmtCVS = headerCVS.indexOf('SumOfAMOUNT');
    var idxNameCVS = headerCVS.indexOf('CUST_NAME');

    for (var j = 1; j < rowsCVS.length; j++) {
      var rCVS = rowsCVS[j];
      if (!rCVS) continue;
      var custCVS = String(rCVS[idxCustCVS] || '');
      var cnameCVS = idxNameCVS >= 0 ? String(rCVS[idxNameCVS] || '') : '';
      var addrCVS = String(rCVS[idxAddrCVS] || '').trim();
      var amtCVS = parseFloat(rCVS[idxAmtCVS]) || 0;
      var normACVS = normalizeText(addrCVS);

      // GS25: Mã GS0003
      if (custCVS === 'GS0003') {
        if (normACVS.indexOf('longhau') >= 0 || normACVS.indexOf('h04') >= 0) {
          gs25Amounts[master.gs25_dcs[0]] += amtCVS;
          totalGs25Dcs += amtCVS;
        } else if (normACVS.indexOf('binhdien') >= 0 || normACVS.indexOf('iiib2') >= 0) {
          gs25Amounts[master.gs25_dcs[1]] += amtCVS;
          totalGs25Dcs += amtCVS;
        }
      }

      // 7-Eleven: SS4001, SS4002, SS5003 hoặc tên có Seven
      if (['SS4001', 'SS4002', 'SS5003'].indexOf(custCVS) >= 0 || cnameCVS.toLowerCase().indexOf('seven') >= 0) {
        if (normACVS.indexOf('tanphutrung') >= 0 || normACVS.indexOf('d2') >= 0) {
          seAmounts[master.se_dcs[0]] += amtCVS;
          totalSeDcs += amtCVS;
        } else if (normACVS.indexOf('tanbinh') >= 0 || normACVS.indexOf('ii3') >= 0) {
          seAmounts[master.se_dcs[1]] += amtCVS;
          totalSeDcs += amtCVS;
        }
      }

      // Circle K: VT4050, VT3013, VT3014, VT3015
      if (['VT4050', 'VT3013', 'VT3014', 'VT3015'].indexOf(custCVS) >= 0) {
        if (normACVS.indexOf('namtanuyen') >= 0 || normACVS.indexOf('g19') >= 0) {
          ckKhoKhoAmt += amtCVS;
        } else {
          ckStoresSo[addrCVS] = (ckStoresSo[addrCVS] || 0) + amtCVS;
        }
      }

      // NPP Hoàng Đức: HD3024
      if (custCVS === 'HD3024' || cnameCVS.toLowerCase().indexOf('hoàng đức') >= 0 || cnameCVS.toLowerCase().indexOf('hoang duc') >= 0) {
        if (addrCVS.indexOf('198') >= 0 || addrCVS.toLowerCase().indexOf('hùng vương') >= 0) {
          hoangDucStores['198 Hùng Vương'] += amtCVS;
        } else if (addrCVS.toLowerCase().indexOf('bạch lâm') >= 0 || addrCVS.indexOf('166') >= 0) {
          hoangDucStores['Bạch Lâm'] += amtCVS;
        }
      }
    }
  }

  var gs25PerStore = Math.round((totalGs25Dcs * 0.2046) / 72);
  var sePerStore = Math.round((totalSeDcs * 0.0362) / 5);
  var totalCkStoresSystem = Object.keys(ckStoresSo).length;
  var tCk = totalCkStoresSystem > 0 ? Math.round(ckKhoKhoAmt / totalCkStoresSystem) : 0;

  // 3. BÓC TÁCH ĐƠN HÀNG SKU FAMILYMART TỪ SHEET NPP
  var nppData = {};
  var sheetNPP = null;
  for (var sIdx = 0; sIdx < wbCVS.SheetNames.length; sIdx++) {
    var sNameNPP = wbCVS.SheetNames[sIdx];
    var upper = sNameNPP.toUpperCase();
    if (upper === 'NPP' || upper.indexOf('NPP') >= 0 || upper.indexOf('FAMILY') >= 0) {
      sheetNPP = sNameNPP;
      break;
    }
  }

  if (sheetNPP) {
    var rowsNPP = XLSX.utils.sheet_to_json(wbCVS.Sheets[sheetNPP], { header: 1 });
    if (rowsNPP.length > 1) {
      var idxOtype = 0, idxOcode = 4, idxOname = 5, idxOaddr = 7, idxIcode = 8, idxIname = 9, idxQty = 11, idxAmt = 12;
      var hRow = rowsNPP[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
      if (hRow.indexOf('CUST_NO') >= 0 || hRow.indexOf('MA_KH') >= 0 || hRow.indexOf('OCODE') >= 0) {
        idxOcode = hRow.findIndex(function(h) { return h.indexOf('CUST_NO') >= 0 || h.indexOf('MA_KH') >= 0 || h.indexOf('OCODE') >= 0; });
        idxOaddr = hRow.findIndex(function(h) { return h.indexOf('ADDR') >= 0 || h.indexOf('DIA_CHI') >= 0; });
        idxIcode = hRow.findIndex(function(h) { return h.indexOf('ITEM') >= 0 || h.indexOf('MA_SP') >= 0 || h.indexOf('ICODE') >= 0; });
        idxIname = hRow.findIndex(function(h) { return h.indexOf('NAME') >= 0 || h.indexOf('TEN_SP') >= 0 || h.indexOf('INAME') >= 0; });
        idxAmt = hRow.findIndex(function(h) { return h.indexOf('AMOUNT') >= 0 || h.indexOf('THANH_TIEN') >= 0 || h.indexOf('AMT') >= 0; });
        idxOtype = hRow.findIndex(function(h) { return h.indexOf('TYPE') >= 0 || h.indexOf('LOAI') >= 0; });
        idxOname = hRow.findIndex(function(h) { return h.indexOf('CUST_NAME') >= 0 || h.indexOf('TEN_KH') >= 0; });
      }

      for (var rNppIdx = 1; rNppIdx < rowsNPP.length; rNppIdx++) {
        var rN = rowsNPP[rNppIdx];
        if (!rN || rN.length === 0) continue;
        var otype = String(rN[idxOtype] || '');
        var oname = String(rN[idxOname] || '');
        if (otype.toUpperCase().indexOf('FM') >= 0 || oname.toUpperCase().indexOf('FAMILY') >= 0) {
          var ocode = String(rN[idxOcode] || '').trim();
          var oaddr = String(rN[idxOaddr] || '').trim();
          var icode = String(rN[idxIcode] || '').trim();
          var iname = String(rN[idxIname] || '').trim();
          var qty = parseFloat(rN[idxQty]) || 0;
          var amt = parseFloat(rN[idxAmt]) || 0;

          if (!nppData[ocode]) nppData[ocode] = { addr: oaddr, items: {}, total_amt: 0 };
          if (!nppData[ocode].items[icode]) nppData[ocode].items[icode] = { name: iname, qty: 0, amt: 0 };
          nppData[ocode].items[icode].qty += qty;
          nppData[ocode].items[icode].amt += amt;
          nppData[ocode].total_amt += amt;
        }
      }
    }
  }

  // Khớp nối dữ liệu 35 cửa hàng FamilyMart
  var fmStoreMaster = (master.fm_stores && master.fm_stores.length > 0) ? master.fm_stores : (typeof DEFAULT_FM_STORES !== 'undefined' ? DEFAULT_FM_STORES : []);
  var matchedFmMatrix = [];
  var hasNppOrders = Object.keys(nppData).length > 0;

  for (var fmIdx = 0; fmIdx < fmStoreMaster.length; fmIdx++) {
    var sFm = fmStoreMaster[fmIdx];
    var codeFm = sFm.code || '';
    var matchedNpp = null;

    if (hasNppOrders) {
      if (codeFm && nppData[codeFm]) {
        matchedNpp = nppData[codeFm];
      } else {
        var suffix = codeFm.length >= 4 ? codeFm.slice(-4) : codeFm;
        var candidates = Object.keys(nppData).filter(function(k) { return k.endsWith(suffix); });
        if (candidates.length === 1) {
          matchedNpp = nppData[candidates[0]];
        } else {
          var stStreet = (sFm.addr || '').split(',')[0].trim().toLowerCase();
          if (stStreet.length > 3) {
            for (var nCode in nppData) {
              var nAddr = (nppData[nCode].addr || '').toLowerCase();
              if (nAddr.indexOf(stStreet) >= 0 || stStreet.indexOf(nAddr) >= 0) {
                matchedNpp = nppData[nCode];
                break;
              }
            }
          }
        }
      }
    }

    var orders = {};
    if (matchedNpp && matchedNpp.items) {
      for (var ic in matchedNpp.items) {
        orders[ic] = Math.round(matchedNpp.items[ic].amt || 0);
      }
    } else if (!hasNppOrders && sFm.orders) {
      Object.assign(orders, sFm.orders);
    }

    var actualFm = Object.values(orders).reduce(function(a, b) { return a + b; }, 0);

    matchedFmMatrix.push({
      addr: sFm.addr,
      location: sFm.location || 'ST00_CVS_FM',
      ma_pg: sFm.ma_pg || '',
      ten_pg: sFm.ten_pg || '',
      vtcv: sFm.vtcv || 'SR',
      target: sFm.target || 0,
      actual: actualFm,
      orders: orders
    });
  }

  matchedFmMatrix.sort(function(a, b) {
    var pgCmp = (a.ten_pg || '').localeCompare(b.ten_pg || '');
    if (pgCmp !== 0) return pgCmp;
    return (a.addr || '').localeCompare(b.addr || '');
  });

  // 4. CẬP NHẬT DOANH SỐ CHO TỪNG ĐIỂM BÁN (STORES)
  var storesRef = (typeof getInitialStores === 'function' ? getInitialStores() : (master.stores || []));
  var updatedStores = [];

  for (var stIdx = 0; stIdx < storesRef.length; stIdx++) {
    var s = storesRef[stIdx];
    var actualStore = 0;
    var ch = s.chain;
    var addr = s.store_address;
    var code = s.store_code;

    if (ch === 'GS25') {
      actualStore = gs25PerStore;
    } else if (ch === '7-Eleven' || ch === '7E') {
      actualStore = sePerStore;
    } else if (ch === 'Circle K') {
      var foundAmt = ckStoresSo[addr] || 0;
      if (!foundAmt) {
        var normS = normalizeText(addr);
        for (var aSo in ckStoresSo) {
          if (normS === normalizeText(aSo)) {
            foundAmt = ckStoresSo[aSo];
            break;
          }
        }
      }
      actualStore = foundAmt > 0 ? (tCk + Math.round(foundAmt)) : 0;
    } else if (ch.indexOf('Hoàng Đức') >= 0) {
      if (addr.indexOf('198') >= 0 || addr.toLowerCase().indexOf('hùng vương') >= 0) {
        actualStore = Math.round(hoangDucStores['198 Hùng Vương'] || 0);
      } else {
        actualStore = Math.round(hoangDucStores['Bạch Lâm'] || 0);
      }
    } else if (ch === 'FamilyMart') {
      var fmMatch = matchedFmMatrix.find(function(m) {
        return m.addr === addr || normalizeText(m.addr) === normalizeText(addr);
      });
      if (fmMatch && fmMatch.actual > 0) {
        actualStore = fmMatch.actual;
      } else {
        actualStore = s.actual || 0;
      }
    } else if (['WinMart+', 'WMP'].indexOf(ch) >= 0) {
      actualStore = Math.round(wmpAmounts[code] || 0);
    } else {
      actualStore = s.actual || 0;
    }

    updatedStores.push({
      employee_name: s.employee_name,
      chain: s.chain,
      store_code: s.store_code,
      store_address: s.store_address,
      actual: actualStore
    });
  }

  // 5. TỔNG HỢP THEO TỪNG NHÂN VIÊN
  var totalTargetTeam = 0;
  var totalCvsTeam = 0;
  var employeesSummary = [];

  for (var eIdx = 0; eIdx < master.employees.length; eIdx++) {
    var emp = master.employees[eIdx];
    totalTargetTeam += emp.target;
    var bhxAct = Math.round(bhxPerStore * emp.bhx_stores);

    var empStores = updatedStores.filter(function(st) { return st.employee_name === emp.name; });
    var gs25Act = 0, seAct = 0, fmAct = 0, ckAct = 0, wmpAct = 0, hdAct = 0;

    for (var esIdx = 0; esIdx < empStores.length; esIdx++) {
      var valSt = empStores[esIdx].actual || 0;
      var cName = empStores[esIdx].chain;
      if (cName === 'GS25') gs25Act += valSt;
      else if (cName === '7-Eleven') seAct += valSt;
      else if (cName === 'FamilyMart') fmAct += valSt;
      else if (cName === 'Circle K') ckAct += valSt;
      else if (cName.indexOf('WinMart') >= 0) wmpAct += valSt;
      else if (cName.indexOf('Hoàng Đức') >= 0) hdAct += valSt;
    }

    var cvsTotal = gs25Act + seAct + fmAct + ckAct + wmpAct + hdAct;
    totalCvsTeam += cvsTotal;
    var totalAct = bhxAct + cvsTotal;
    var missing = emp.target - totalAct;
    var pct = emp.target > 0 ? parseFloat(((totalAct / emp.target) * 100).toFixed(2)) : 0;
    var status = pct >= timePct ? 'VƯỢT' : (pct >= timePct * 0.75 ? 'CẬN' : 'CHẬM');

    employeesSummary.push({
      name: emp.name,
      target: emp.target,
      bhx_stores: emp.bhx_stores,
      bhx_actual: bhxAct,
      gs25_actual: gs25Act,
      se_actual: seAct,
      fm_actual: fmAct,
      ck_actual: ckAct,
      wmp_actual: wmpAct,
      hd_actual: hdAct,
      cvs_total: cvsTotal,
      total_actual: totalAct,
      missing: missing,
      percent: pct,
      status: status
    });
  }

  var totalBhxTeam = employeesSummary.reduce(function(acc, e) { return acc + (e.bhx_actual || 0); }, 0);
  var totalStoresBhxTeam = employeesSummary.reduce(function(acc, e) { return acc + (e.bhx_stores || 0); }, 0);
  var totalActualTeam = totalBhxTeam + totalCvsTeam;
  var teamPct = totalTargetTeam > 0 ? parseFloat(((totalActualTeam / totalTargetTeam) * 100).toFixed(2)) : 0;

  var bhxHubsFormatted = master.bhx_hubs.map(function(h) {
    return {
      name: h,
      amount: Math.round(bhxHubAmounts[h] || 0)
    };
  });

  var totFmTgt = matchedFmMatrix.reduce(function(a, b) { return a + (b.target || 0); }, 0);
  var totFmAct = matchedFmMatrix.reduce(function(a, b) { return a + (b.actual || 0); }, 0);
  var pctFmTotal = totFmTgt > 0 ? parseFloat((totFmAct / totFmTgt * 100).toFixed(1)) : 0;
  var passedFmCount = matchedFmMatrix.filter(function(m) {
    var p = m.target > 0 ? (m.actual / m.target * 100) : 0;
    return p >= timePct;
  }).length;

  var fmCategories = master.fm_categories || (typeof FM_CATEGORIES !== 'undefined' ? FM_CATEGORIES : []);
  var fmSkuMatrixData = {
    stores: matchedFmMatrix,
    categories: fmCategories,
    total_target: totFmTgt,
    total_actual: totFmAct,
    pct: pctFmTotal,
    passed_count: passedFmCount
  };

  return {
    team_lead: master.team_lead || 'Trần Thị Cẩm Giang',
    month: month,
    year: year,
    timegone: timePct,
    total_target: totalTargetTeam,
    total_actual_bhx: totalBhxTeam,
    total_actual_bhx_hubs: Math.round(totalActualBhx),
    total_stores_bhx_team: totalStoresBhxTeam,
    total_stores_bhx: totalStoresBhx,
    bhx_per_store: Math.round(bhxPerStore),
    total_cvs: totalCvsTeam,
    total_actual: totalActualTeam,
    percent_achieved: teamPct,
    bhx_hubs: bhxHubsFormatted,
    employees: employeesSummary,
    stores: updatedStores,
    fm_sku_matrix: fmSkuMatrixData
  };
}

/**
 * Hàm tổng xử lý khi nhận 2 file nhị phân Blob từ Telegram
 * @param {Blob} blob1 Blob File 1
 * @param {Blob} blob2 Blob File 2
 * @returns {Object} Payload dữ liệu báo cáo
 */
function processUploadedBlobs(blob1, blob2) {
  var wb1 = parseWorkbook(blob1);
  var wb2 = parseWorkbook(blob2);

  var type1 = detectFileType(blob1.getName(), wb1);
  var type2 = detectFileType(blob2.getName(), wb2);

  var wbST = null;
  var wbCVS = null;

  if (type1 === 'ST' || type2 === 'CVS') {
    wbST = wb1;
    wbCVS = wb2;
  } else if (type1 === 'CVS' || type2 === 'ST') {
    wbST = wb2;
    wbCVS = wb1;
  } else {
    // Mặc định blob1 là ST, blob2 là CVS
    wbST = wb1;
    wbCVS = wb2;
  }

  return processTwoWorkbooks(wbST, wbCVS);
}

if (typeof module !== "undefined") {
  module.exports = {
    normalizeText: normalizeText,
    calculateRealtimeTimegone: calculateRealtimeTimegone,
    detectFileType: detectFileType,
    processTwoWorkbooks: processTwoWorkbooks
  };
}
