const XLSX = require('xlsx');
global.XLSX = XLSX;
const zlib = require('zlib');

const masterData = require('../master_data.js');
const { detectFileType, processTwoWorkbooks } = require('../ReportEngine.gs');
const { buildExcelWorkbook } = require('../lib/excel-builder.js');
const { generateReportUrl, generateCleanReportUrl } = require('../lib/report-share.js');

// Bộ nhớ đệm phiên làm việc trên Vercel Serverless
global.__TG_SESSIONS = global.__TG_SESSIONS || new Map();
const sessions = global.__TG_SESSIONS;

// Xóa session quá hạn (> 2 giờ)
function cleanExpiredSessions() {
  const now = Date.now();
  for (const [chatId, sess] of sessions.entries()) {
    if (now - sess.timestamp > 2 * 3600 * 1000) {
      sessions.delete(chatId);
    }
  }
}

// Định dạng tiền VNĐ
function formatMoney(num) {
  return (num || 0).toLocaleString('vi-VN');
}

// =========================================================================
// CÁC HÀM GỌI TELEGRAM BOT API NATIVE
// =========================================================================

async function sendMessage(token, chatId, text, replyMarkup) {
  const body = {
    chat_id: String(chatId),
    text: text,
    parse_mode: 'HTML'
  };
  if (replyMarkup) body.reply_markup = replyMarkup;
  const res = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  return res.json();
}

async function sendChatAction(token, chatId, action) {
  try {
    await fetch(`https://api.telegram.org/bot${token}/sendChatAction`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: String(chatId),
        action: action
      })
    });
  } catch (e) {}
}

async function sendPhoto(token, chatId, imageSource, fileName, caption) {
  // 1. Nếu imageSource là URL string
  if (typeof imageSource === 'string' && imageSource.startsWith('http')) {
    const res = await fetch(`https://api.telegram.org/bot${token}/sendPhoto`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: String(chatId),
        photo: imageSource,
        caption: caption || '',
        parse_mode: 'HTML'
      })
    });
    return res.json();
  }

  // 2. Nếu imageSource là Buffer
  const formData = new FormData();
  formData.append('chat_id', String(chatId));
  formData.append('photo', new Blob([imageSource], { type: 'image/png' }), fileName || 'Bao_Cao.png');
  if (caption) {
    formData.append('caption', caption);
    formData.append('parse_mode', 'HTML');
  }

  const res = await fetch(`https://api.telegram.org/bot${token}/sendPhoto`, {
    method: 'POST',
    body: formData
  });
  return res.json();
}

async function sendDocument(token, chatId, docBuffer, fileName, caption) {
  const formData = new FormData();
  formData.append('chat_id', String(chatId));
  formData.append('document', new Blob([docBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName || 'Team_CamGiang_Report.xlsx');
  if (caption) {
    formData.append('caption', caption);
    formData.append('parse_mode', 'HTML');
  }

  const res = await fetch(`https://api.telegram.org/bot${token}/sendDocument`, {
    method: 'POST',
    body: formData
  });
  return res.json();
}

async function downloadTelegramFileBuffer(token, fileId) {
  const getFileRes = await fetch(`https://api.telegram.org/bot${token}/getFile?file_id=${fileId}`);
  const fileData = await getFileRes.json();
  if (!fileData.ok || !fileData.result || !fileData.result.file_path) {
    throw new Error('Không thể lấy đường dẫn file từ Telegram: ' + (fileData.description || ''));
  }

  const downloadUrl = `https://api.telegram.org/file/bot${token}/${fileData.result.file_path}`;
  const fileRes = await fetch(downloadUrl);
  const arrayBuffer = await fileRes.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

// Nén dữ liệu và tạo đường dẫn render 4 bảng báo cáo HTML
function getCompressedTableUrl(type, payload) {
  const common = {
    m: payload.month || 9,
    y: payload.year || 2026,
    tl: payload.team_lead || 'Trần Thị Cẩm Giang',
    tg: payload.timegone || 0,
    pct: payload.percent_achieved || 0,
    dt: new Date().toLocaleDateString('vi-VN') + ' | ' + new Date().toLocaleTimeString('vi-VN')
  };

  let dataObj = null;

  if (type === 1) {
    // 1. Bảng Tiến Độ 10 Nhân Viên Toàn Team (BHX + CVS)
    dataObj = {
      ...common,
      emps: (payload.employees || []).map(e => [
        e.name,
        e.bhx_actual || 0,
        e.gs25_actual || 0,
        e.se_actual || 0,
        e.fm_actual || 0,
        e.ck_actual || 0,
        e.hd_actual || 0,
        e.wmp_actual || 0,
        e.total_actual || 0
      ])
    };
  } else if (type === 2) {
    // 2. Bảng Xếp Hạng Doanh Số Toàn Team (Xếp theo Tổng Thực Hiện)
    const sorted = (payload.employees || []).slice().sort((a, b) => (b.total_actual || 0) - (a.total_actual || 0));
    dataObj = {
      ...common,
      ranks: sorted.map(e => [
        e.name,
        e.bhx_actual || 0,
        e.cvs_total || 0,
        e.total_actual || 0
      ])
    };
  } else if (type === 3) {
    // 3. Bảng Phân Bổ Doanh Số 5 Hubs Bách Hóa Xanh (10 NV)
    const totalBhxTeam = (payload.employees || []).reduce((a, e) => a + (e.bhx_actual || 0), 0) || payload.total_actual_bhx || 0;
    dataObj = {
      ...common,
      pricePerStore: payload.bhx_per_store || 0,
      bhxDist: (payload.employees || []).map(e => [
        e.name,
        e.bhx_stores || 0,
        payload.bhx_per_store || 0,
        e.bhx_actual || 0,
        totalBhxTeam > 0 ? ((e.bhx_actual / totalBhxTeam) * 100).toFixed(1) : '0.0'
      ])
    };
  } else if (type === 4) {
    // 4. Bảng Tổng Hợp Doanh Số Nhập FamilyMart (Bỏ cột Target và % Đạt)
    const fmStores = (payload.fm_sku_matrix && payload.fm_sku_matrix.stores) ? payload.fm_sku_matrix.stores : [];
    dataObj = {
      ...common,
      stores: fmStores.map(s => [
        s.addr || '',
        s.location || 'ST00_CVS_FM',
        s.ma_pg || '',
        s.ten_pg || '',
        s.vtcv || 'SR',
        s.actual || 0
      ])
    };
  }

  const jsonStr = JSON.stringify(dataObj);
  const compressed = zlib.deflateSync(jsonStr).toString('base64url');
  return `https://bcgiang.vercel.app/render.html?t=${type}&d=${compressed}`;
}

// Chụp ảnh bảng HTML chuẩn Retina từ Microlink hoặc Fallback Thum.io
async function fetchScreenshotImage(targetUrl) {
  // 1. Microlink Screenshot API
  try {
    const microApi = `https://api.microlink.io?url=${encodeURIComponent(targetUrl)}&screenshot=true&screenshot.element=%23renderCard&viewport.width=1920&viewport.deviceScaleFactor=2&meta=false`;
    const res = await fetch(microApi, { signal: AbortSignal.timeout(14000) });
    const json = await res.json();
    if (json && json.status === 'success' && json.data && json.data.screenshot && json.data.screenshot.url) {
      try {
        const imgRes = await fetch(json.data.screenshot.url, { signal: AbortSignal.timeout(9000) });
        if (imgRes.ok) {
          const arr = await imgRes.arrayBuffer();
          return { buffer: Buffer.from(arr), url: json.data.screenshot.url };
        }
      } catch (eBuf) {
        return { buffer: null, url: json.data.screenshot.url };
      }
    }
  } catch (e1) {
    console.warn('Microlink error, falling back:', e1.message);
  }

  // 2. Thum.io fallback
  try {
    const thumUrl = `https://image.thum.io/get/width/1400/crop/1200/noanimate/${encodeURIComponent(targetUrl)}`;
    const res = await fetch(thumUrl, { signal: AbortSignal.timeout(14000) });
    if (res.ok) {
      const arr = await res.arrayBuffer();
      return { buffer: Buffer.from(arr), url: thumUrl };
    }
  } catch (e2) {
    console.warn('thum.io fallback error:', e2.message);
  }

  return null;
}

// =========================================================================
// ĐIỂM ĐÓN WEBHOOK CHÍNH TRÊN VERCEL
// =========================================================================

module.exports = async function handler(req, res) {
  // 1. Healthcheck khi mở bằng trình duyệt GET
  if (req.method === 'GET') {
    return res.status(200).json({
      status: 'ok',
      message: '🚀 Telegram Bot Webhook is ACTIVE on Vercel!',
      time: new Date().toISOString()
    });
  }

  // 2. Chỉ nhận POST từ Telegram
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  const token = req.query.token || process.env.TELEGRAM_BOT_TOKEN;
  if (!token) {
    console.error('❌ Thiếu TELEGRAM_BOT_TOKEN trong URL hoặc Environment Variables');
    return res.status(400).json({ error: 'Missing Telegram Bot Token' });
  }

  cleanExpiredSessions();

  try {
    const update = req.body;
    if (!update || !update.message) {
      return res.status(200).json({ ok: true, note: 'No message in update' });
    }

    const msg = update.message;
    const chatId = msg.chat.id;
    const text = (msg.text || '').trim();

    // -------------------------------------------------------------
    // 3. XỬ LÝ LỆNH DẠNG TEXT
    // -------------------------------------------------------------
    if (text.startsWith('/start') || text.startsWith('/help')) {
      const welcome = 
`📊 <b>HỆ THỐNG BÁO CÁO DOANH SỐ TEAM (CVS & BHX)</b>
<i>Chạy trực tiếp trên nền tảng Vercel Cloud 🚀</i>

🤖 <b>Chào bạn! Tôi là Bot Báo Cáo Doanh Số tự động.</b>

👉 <b>CÁCH SỬ DỤNG:</b>
Gửi cho tôi <b>2 file Excel (.xlsb hoặc .xlsx)</b>:
1️⃣ <b>File Siêu Thị:</b> Chứa dữ liệu BHX 5 Hubs + WinMart+
2️⃣ <b>File CVS & NPP:</b> Chứa dữ liệu GS25, 7-Eleven, FamilyMart, Circle K, Hoàng Đức

⚡ <b>Bot sẽ tự động trả về:</b>
• 📊 <b>Ảnh biểu đồ & Xếp hạng tiến độ</b> trực quan
• 📁 <b>File Excel chuẩn 5 sheets</b> <code>Team_CamGiang_Report.xlsx</code>
• 💬 <b>Bản tin tóm tắt số liệu</b> chi tiết toàn team!

📌 <b>Lệnh nhanh:</b>
/link - Lấy lại link xem báo cáo trực tuyến mới nhất
/status - Kiểm tra trạng thái 2 file đã gửi
/reset - Xóa phiên để gửi lại từ đầu`;

      await sendMessage(token, chatId, welcome);
      return res.status(200).json({ ok: true });
    }

    if (text.startsWith('/link') || text.startsWith('/xem') || text.startsWith('/online')) {
      if (global.__LAST_REPORT_URL) {
        const ik = {
          inline_keyboard: [
            [{ text: '🌐 MỞ XEM BÁO CÁO TRỰC TUYẾN', url: global.__LAST_REPORT_URL }]
          ]
        };
        await sendMessage(token, chatId, `🌐 <b>BÁO CÁO DOANH SỐ TRỰC TUYẾN MỚI NHẤT:</b>\n👉 <a href="${global.__LAST_REPORT_URL}">Bấm vào đây để mở Báo Cáo Trên Web</a>\n\n<i>Cấp trên có thể bấm nút bên dưới để xem trực tiếp đầy đủ các bảng biểu và tải file Excel trên điện thoại/máy tính.</i>`, ik);
      } else {
        await sendMessage(token, chatId, '💡 <i>Chưa có báo cáo nào vừa được tạo trong phiên này. Hãy gửi 2 file doanh số để Bot tạo báo cáo và trả link trực tuyến ngay!</i>');
      }
      return res.status(200).json({ ok: true });
    }

    if (text.startsWith('/reset')) {
      sessions.delete(chatId);
      await sendMessage(token, chatId, '🔄 <b>Đã làm mới phiên!</b> Bạn hãy gửi 2 file Excel doanh số (.xlsb hoặc .xlsx) để bắt đầu.');
      return res.status(200).json({ ok: true });
    }

    if (text.startsWith('/status')) {
      const sess = sessions.get(chatId);
      if (!sess) {
        await sendMessage(token, chatId, 'ℹ️ Hiện tại bạn chưa gửi file nào. Hãy gửi File Siêu Thị (BHX) và File CVS & NPP để bắt đầu.');
      } else {
        let stText = '📋 <b>TRẠNG THÁI HIỆN TẠI:</b>\n';
        stText += sess.fileST ? `• 🏬 File Siêu Thị: ✅ <i>${sess.fileST.fileName}</i>\n` : '• 🏬 File Siêu Thị: ❌ Chưa nhận\n';
        stText += sess.fileCVS ? `• 🏪 File CVS & NPP: ✅ <i>${sess.fileCVS.fileName}</i>\n` : '• 🏪 File CVS & NPP: ❌ Chưa nhận\n';
        if (!sess.fileST || !sess.fileCVS) {
          stText += '\n👉 Vui lòng gửi tiếp file còn thiếu!';
        }
        await sendMessage(token, chatId, stText);
      }
      return res.status(200).json({ ok: true });
    }

    // -------------------------------------------------------------
    // 4. XỬ LÝ KHI NGƯỜI DÙNG GỬI FILE (DOCUMENT)
    // -------------------------------------------------------------
    if (msg.document) {
      const doc = msg.document;
      const fileName = doc.file_name || 'unknown.xlsx';
      const fileId = doc.file_id;
      const fileSizeMB = (doc.file_size || 0) / (1024 * 1024);

      if (fileSizeMB > 25) {
        await sendMessage(token, chatId, `⚠️ File <b>${fileName}</b> vượt quá giới hạn 25MB của Telegram.`);
        return res.status(200).json({ ok: true });
      }

      await sendChatAction(token, chatId, 'typing');

      // Tải buffer và parse workbook
      const buffer = await downloadTelegramFileBuffer(token, fileId);
      const wb = XLSX.read(buffer, {
        type: 'buffer',
        cellDates: false,
        cellStyles: false,
        cellHTML: false,
        cellFormula: false
      });

      const detectedType = detectFileType(fileName, wb);
      if (detectedType === 'UNKNOWN') {
        await sendMessage(token, chatId, `❓ <b>Chưa nhận diện được file:</b> ${fileName}\nVui lòng gửi File Siêu Thị (BHX) hoặc File CVS & NPP.`);
        return res.status(200).json({ ok: true });
      }

      // Quản lý phiên làm việc
      let sess = sessions.get(chatId) || { timestamp: Date.now() };
      sess.timestamp = Date.now();

      if (detectedType === 'ST') {
        sess.fileST = { fileId, fileName, wb };
      } else if (detectedType === 'CVS') {
        sess.fileCVS = { fileId, fileName, wb };
      }
      sessions.set(chatId, sess);

      // Kiểm tra xem đã đủ 2 file chưa
      if (sess.fileST && sess.fileCVS) {
        await sendMessage(token, chatId, 
`✨ <b>ĐÃ NHẬN ĐỦ CẢ 2 FILE!</b>
1️⃣ 🏬 <b>File Siêu Thị:</b> ${sess.fileST.fileName}
2️⃣ 🏪 <b>File CVS & NPP:</b> ${sess.fileCVS.fileName}

⏳ <i>Hệ thống Vercel đang bóc tách số liệu, vẽ biểu đồ và tạo file Excel... Vui lòng đợi trong giây lát!</i>`);

        await sendChatAction(token, chatId, 'upload_document');

        // Bóc tách số liệu qua Report Engine
        const payload = processTwoWorkbooks(sess.fileST.wb, sess.fileCVS.wb, masterData);

        // Tạo link xem báo cáo trực tuyến trên Vercel qua Cloud Storage (Giải pháp 2C siêu ngắn, chuyên nghiệp)
        const reportWebUrl = await generateCleanReportUrl(req, payload);
        global.__LAST_REPORT_URL = reportWebUrl;

        // Tùy chọn gọi Google Apps Script tạo Google Sheet nếu có cấu hình GOOGLE_WEBAPP_URL
        let googleSheetUrl = null;
        const gasUrl = process.env.GOOGLE_WEBAPP_URL || process.env.GAS_URL;
        if (gasUrl) {
          try {
            const gasRes = await fetch(gasUrl, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ action: 'create_sheet', payload })
            });
            const gasData = await gasRes.json();
            if (gasData && gasData.url) {
              googleSheetUrl = gasData.url;
            }
          } catch (eGas) {
            console.warn('Lỗi gọi Google Apps Script tạo sheet:', eGas.message);
          }
        }

        // 1. Chụp và gửi lần lượt 4 ảnh bảng biểu chuẩn theo đúng mẫu giao diện
        const fmCount = (payload.fm_sku_matrix && payload.fm_sku_matrix.stores) ? payload.fm_sku_matrix.stores.length : 35;
        const fmActual = payload.fm_sku_matrix ? (payload.fm_sku_matrix.total_actual || 0) : 0;

        const tableConfigs = [
          {
            type: 1,
            fileName: `Bang_Tien_Do_10_NV_${payload.month}_${payload.year}.png`,
            caption: `📊 <b>1/4. Bảng Tiến Độ 10 Nhân Viên Toàn Team (BHX + CVS)</b>\n👤 Team Lead: ${payload.team_lead} | 📅 Tháng ${payload.month}/${payload.year}\n💰 Tổng Thực Hiện: ${formatMoney(payload.total_actual)} đ (BHX: ${formatMoney(payload.total_actual_bhx)} đ | CVS: ${formatMoney(payload.total_cvs)} đ)`
          },
          {
            type: 2,
            fileName: `Bang_Xep_Hang_Team_${payload.month}_${payload.year}.png`,
            caption: `🏆 <b>2/4. Bảng Xếp Hạng Doanh Số Toàn Team</b>\n👥 10 Nhân viên kinh doanh | 💰 Tổng TH: ${formatMoney(payload.total_actual)} đ\n⏳ % Timegone (tiến độ tháng): ${payload.timegone}%`
          },
          {
            type: 3,
            fileName: `Phan_Bo_5_Hubs_BHX_${payload.month}_${payload.year}.png`,
            caption: `🚚 <b>3/4. Bảng Phân Bổ Doanh Số 5 Hubs Bách Hóa Xanh (10 NV)</b>\n🛒 Tổng BHX: ${formatMoney(payload.total_actual_bhx)} đ (${payload.total_stores_bhx_team || 178} Cửa hàng)\n🏷 Đơn giá phân bổ: ${formatMoney(payload.bhx_per_store)} đ/CH`
          },
          {
            type: 4,
            fileName: `Chi_Tiet_SKU_FamilyMart_${payload.month}_${payload.year}.png`,
            caption: `🛒 <b>4/4. Bảng Chi Tiết SKU Cửa Hàng FamilyMart</b>\n🏪 Kênh CVS FamilyMart (${fmCount} Cửa hàng)\n💰 Tổng thực hiện nhập: ${formatMoney(fmActual)} đ`
          }
        ];

        for (const cfg of tableConfigs) {
          try {
            await sendChatAction(token, chatId, 'upload_photo');
            const targetUrl = getCompressedTableUrl(cfg.type, payload);
            const shot = await fetchScreenshotImage(targetUrl);
            if (shot) {
              const photoData = shot.buffer || shot.url;
              await sendPhoto(token, chatId, photoData, cfg.fileName, cfg.caption);
            }
          } catch (eShot) {
            console.error(`Lỗi chụp/gửi ảnh bảng ${cfg.type}:`, eShot.message);
          }
        }

        // 2. Tạo file Excel chuẩn 5 sheets qua ExcelJS & Gửi đính kèm
        try {
          const excelWb = await buildExcelWorkbook(payload);
          const excelBuffer = await excelWb.xlsx.writeBuffer();
          const docCaption = 
`📊 <b>BÁO CÁO DOANH SỐ TEAM ${payload.team_lead.toUpperCase()}</b>
Tháng ${payload.month}/${payload.year} (% Timegone: ${payload.timegone}%)
• Tổng Thực Hiện: ${formatMoney(payload.total_actual)} ₫
• BHX: ${formatMoney(payload.total_actual_bhx)} ₫
• CVS: ${formatMoney(payload.total_cvs)} ₫
🌐 Link xem online: ${reportWebUrl}`;

          const excelFileName = `Team_${payload.team_lead.replace(/\s+/g, '_')}_Report_Thang_${payload.month}_${payload.year}.xlsx`;
          await sendDocument(token, chatId, excelBuffer, excelFileName, docCaption);
        } catch (eDoc) {
          console.error('Lỗi tạo file Excel:', eDoc);
        }

        // 3. Gửi tin nhắn tóm tắt số liệu chi tiết
        const topEmployees = (payload.employees || []).slice().sort((a, b) => (b.total_actual || 0) - (a.total_actual || 0));
        let summaryMsg = 
`🎯 <b>TỔNG HỢP DOANH SỐ TOÀN TEAM:</b>
━━━━━━━━━━━━━━━━━━━━
• <b>Thực Hiện BHX:</b> <code>${formatMoney(payload.total_actual_bhx)} ₫</code> (TB ${formatMoney(payload.bhx_per_store)} ₫/CH)
• <b>Thực Hiện CVS:</b> <code>${formatMoney(payload.total_cvs)} ₫</code>
• <b>Tổng Thực Hiện:</b> <b>${formatMoney(payload.total_actual)} ₫</b>
• <b>Tiến Độ Thời Gian Tháng:</b> <b>${payload.timegone}%</b>
━━━━━━━━━━━━━━━━━━━━
🏆 <b>BẢNG XẾP HẠNG DOANH SỐ NHÂN VIÊN:</b>\n`;

        for (let i = 0; i < Math.min(10, topEmployees.length); i++) {
          const emp = topEmployees[i];
          const medal = i === 0 ? '🥇' : (i === 1 ? '🥈' : (i === 2 ? '🥉' : `${i + 1}.`));
          const pctShare = payload.total_actual > 0 ? (((emp.total_actual || 0) / payload.total_actual) * 100).toFixed(1) : '0.0';
          summaryMsg += `${medal} <b>${emp.name}</b>: <b>${formatMoney(emp.total_actual)} ₫</b> (${pctShare}% team)\n`;
        }

        summaryMsg += 
`━━━━━━━━━━━━━━━━━━━━
🌐 <b>LINK XEM BÁO CÁO TRỰC TUYẾN:</b>
👉 <b><a href="${reportWebUrl}">${reportWebUrl}</a></b>
<i>(Cấp trên có thể bấm link xem trực tiếp đầy đủ 5 bảng biểu trên điện thoại & máy tính, tải file Excel ngay trên web)</i>\n`;

        if (googleSheetUrl) {
          summaryMsg += 
`━━━━━━━━━━━━━━━━━━━━
📊 <b>LINK GOOGLE SHEET:</b>
👉 <a href="${googleSheetUrl}">Bấm vào đây để mở Google Sheet</a>\n`;
        }

        summaryMsg += 
`━━━━━━━━━━━━━━━━━━━━
<i>Đã đính kèm 4 ảnh chụp bảng biểu và file Excel (.xlsx) ở phía trên!</i>`;

        const inlineKeyboard = {
          inline_keyboard: [
            [
              { text: '🌐 MỞ XEM BÁO CÁO TRỰC TUYẾN', url: reportWebUrl }
            ]
          ]
        };
        if (googleSheetUrl) {
          inlineKeyboard.inline_keyboard.push([
            { text: '📊 MỞ GOOGLE SPREADSHEET', url: googleSheetUrl }
          ]);
        }

        await sendMessage(token, chatId, summaryMsg, inlineKeyboard);

        // Xóa session hoàn tất
        sessions.delete(chatId);

      } else {
        const receivedTitle = detectedType === 'ST' ? '🏬 File Siêu Thị (BHX & WinMart+)' : '🏪 File CVS & NPP (GS25, 7E, CK, FM, Hoàng Đức)';
        const missingTitle = detectedType === 'ST' ? '🏪 File CVS & NPP' : '🏬 File Siêu Thị (BHX)';

        await sendMessage(token, chatId, 
`✅ <b>Đã nhận diện:</b> ${receivedTitle}
📄 <i>${fileName}</i>

⏳ <b>Còn thiếu:</b> ${missingTitle}
👉 <i>Vui lòng gửi tiếp file thứ 2 để Bot tạo báo cáo hoàn chỉnh!</i>`);
      }

      return res.status(200).json({ ok: true });
    }

    // Tin nhắn văn bản không khớp lệnh
    await sendMessage(token, chatId, '💡 <i>Vui lòng gửi 2 file Excel doanh số (.xlsb hoặc .xlsx) để tôi tạo báo cáo, hoặc gõ /help để xem hướng dẫn.</i>');
    return res.status(200).json({ ok: true });

  } catch (err) {
    console.error('Lỗi xử lý webhook:', err);
    return res.status(200).json({ ok: false, error: err.message });
  }
};
