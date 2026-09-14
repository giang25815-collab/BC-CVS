// ============================================================================
// ⚙️ CẤU HÌNH NHANH (BẠN CHỈ CẦN DÁN 2 THÔNG TIN VÀO ĐÂY RỒI CHẠY "ONE_CLICK_SETUP"):
// ============================================================================
var CONFIG_BOT_TOKEN = 'DÁN_MÃ_TOKEN_TỪ_BOTFATHER_VÀO_ĐÂY';
var CONFIG_WEBAPP_URL = 'DÁN_LINK_EXEC_TỪ_DEPLOY_VÀO_ĐÂY';

/**
 * 🌟 HÀM CÀI ĐẶT 1-CLICK TIỆN LỢI NHẤT
 * Chỉ cần chọn hàm này và bấm nút Run (Chạy) ▶️, hệ thống sẽ tự làm mọi thứ!
 */
function ONE_CLICK_SETUP() {
  Logger.log('🚀 Bắt đầu cài đặt Bot Telegram...');

  // 1. Kiểm tra Token
  var token = (CONFIG_BOT_TOKEN || '').trim();
  if (!token || token.indexOf('DÁN_MÃ_TOKEN') >= 0) {
    throw new Error('❌ Bạn chưa điền CONFIG_BOT_TOKEN ở dòng 5! Hãy mở @BotFather copy Token dán vào.');
  }

  // 2. Kiểm tra Token với máy chủ Telegram
  var meUrl = 'https://api.telegram.org/bot' + token + '/getMe';
  var meRes = UrlFetchApp.fetch(meUrl, { muteHttpExceptions: true });
  var meData = JSON.parse(meRes.getContentText());

  if (!meData.ok) {
    Logger.log('❌ LỖI TỪ TELEGRAM: ' + JSON.stringify(meData));
    throw new Error('❌ Mã Token không chính xác (' + meData.description + ')!\nHãy vào @BotFather trên Telegram gõ lệnh /mybots > chọn Bot của bạn > bấm API Token để copy lại mã chuẩn 100%.');
  }

  Logger.log('✅ Token chuẩn xác! Đã nhận diện Bot: @' + meData.result.username + ' (' + meData.result.first_name + ')');

  // Lưu Token vào Properties
  PropertiesService.getScriptProperties().setProperty('TELEGRAM_BOT_TOKEN', token);
  Logger.log('✅ Đã lưu Token vào hệ thống.');

  // 3. Kiểm tra Web App URL
  var url = (CONFIG_WEBAPP_URL || '').trim();
  if (!url || url.indexOf('DÁN_LINK_EXEC') >= 0) {
    url = ScriptApp.getService().getUrl() || '';
  }

  if (!url) {
    throw new Error('❌ Chưa có URL Web App! Hãy bấm Deploy > New deployment > Web app > Anyone để copy link /exec dán vào dòng 6.');
  }

  if (url.endsWith('/dev')) {
    Logger.log('⚠️ URL có đuôi /dev, tự động chuyển sang /exec...');
    url = url.replace(/\/dev$/, '/exec');
  }

  // 4. Đăng ký Webhook
  var hookUrl = 'https://api.telegram.org/bot' + token + '/setWebhook?url=' + encodeURIComponent(url);
  var hookRes = UrlFetchApp.fetch(hookUrl, { muteHttpExceptions: true });
  var hookData = JSON.parse(hookRes.getContentText());

  if (!hookData.ok) {
    Logger.log('❌ Lỗi đăng ký Webhook: ' + JSON.stringify(hookData));
    throw new Error('❌ Không thể đăng ký Webhook: ' + hookData.description);
  }

  Logger.log('====================================================');
  Logger.log('🎉 THÀNH CÔNG 100%! BOT ĐÃ ĐƯỢC KẾT NỐI HOÀN TOÀN!');
  Logger.log('🤖 Tên Bot: @' + meData.result.username);
  Logger.log('🔗 Webhook URL: ' + url);
  Logger.log('👉 Bây giờ hãy mở Telegram và gửi /start vào bot để dùng ngay!');
  Logger.log('====================================================');
}

/**
 * Tự động đăng ký Webhook Telegram trỏ về Web App của Apps Script
 * @param {string} [customUrl] Tuỳ chọn URL Web App cụ thể kết thúc bằng /exec
 */
function setupTelegramWebhook(customUrl) {
  var token = getTelegramBotToken();
  var webAppUrl = customUrl || ScriptApp.getService().getUrl();

  if (!webAppUrl) {
    throw new Error('Chưa lấy được Web App URL. Hãy chắc chắn bạn đã Deploy Web App (Deploy > Manage deployments).');
  }

  // CỰC KỲ QUAN TRỌNG: Telegram chỉ hoạt động với link /exec, link /dev sẽ bị Google chặn đăng nhập
  if (webAppUrl.endsWith('/dev')) {
    Logger.log('⚠️ Phát hiện URL đang có đuôi /dev, tự động chuyển sang đuôi /exec...');
    webAppUrl = webAppUrl.replace(/\/dev$/, '/exec');
  }

  var apiUrl = 'https://api.telegram.org/bot' + token + '/setWebhook?url=' + encodeURIComponent(webAppUrl);
  var response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
  var result = JSON.parse(response.getContentText());

  Logger.log('Đường link Web App được đăng ký: ' + webAppUrl);
  Logger.log('Kết quả từ Telegram: ' + JSON.stringify(result));

  if (result.ok) {
    return '✅ Đăng ký Webhook thành công!\nURL: ' + webAppUrl;
  } else {
    return '❌ Đăng ký Webhook thất bại: ' + (result.description || '');
  }
}

/**
 * Hàm kiểm tra nhanh trạng thái kết nối của Bot Telegram
 * Chọn hàm này và bấm Run để xem Telegram có gặp lỗi gì không
 */
function checkBotStatus() {
  var token = getTelegramBotToken();
  var apiUrl = 'https://api.telegram.org/bot' + token + '/getWebhookInfo';
  var response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
  var data = JSON.parse(response.getContentText());

  Logger.log('================ THÔNG TIN BOT ================');
  Logger.log('URL Webhook hiện tại: ' + (data.result ? data.result.url : 'Chưa có'));
  Logger.log('Tin nhắn đang chờ xử lý: ' + (data.result ? data.result.pending_update_count : 0));
  if (data.result && data.result.last_error_message) {
    Logger.log('⚠️ LỖI GẦN NHẤT TỪ TELEGRAM: ' + data.result.last_error_message);
    Logger.log('Thời gian lỗi: ' + new Date(data.result.last_error_date * 1000).toLocaleString('vi-VN'));
  } else {
    Logger.log('✅ Không có lỗi nào từ Telegram!');
  }
  Logger.log('================================================');
}

/**
 * Hàm cài đặt Webhook thủ công bằng cách dán URL /exec (nếu hàm tự động chưa lấy đúng)
 */
function setWebhookWithExecUrl() {
  var execUrl = 'DÁN_WEB_APP_URL_EXEC_VÀO_ĐÂY'; // <-- Ví dụ: https://script.google.com/macros/s/AKfycb.../exec
  Logger.log(setupTelegramWebhook(execUrl));
}

/**
 * Hủy đăng ký Webhook Telegram khi cần bảo trì
 */
function removeTelegramWebhook() {
  var token = getTelegramBotToken();
  var apiUrl = 'https://api.telegram.org/bot' + token + '/deleteWebhook';
  var response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
  return response.getContentText();
}

/**
 * Kiểm tra trạng thái Webhook hiện tại
 */
function getTelegramWebhookInfo() {
  var token = getTelegramBotToken();
  var apiUrl = 'https://api.telegram.org/bot' + token + '/getWebhookInfo';
  var response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
  return JSON.parse(response.getContentText());
}

/**
 * ============================================================================
 * ĐIỂM ĐÓN WEBHOOK CHÍNH CỦA GOOGLE APPS SCRIPT: doPost(e)
 * ============================================================================
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput('OK - NO DATA');
    }

    var update = JSON.parse(e.postData.contents);
    handleTelegramUpdate(update);

    return ContentService.createTextOutput('OK');
  } catch (err) {
    Logger.log('Lỗi trong doPost: ' + err.toString());
    return ContentService.createTextOutput('ERROR: ' + err.toString());
  }
}

/**
 * Xử lý nội dung tin nhắn hoặc sự kiện từ Telegram Update
 */
function handleTelegramUpdate(update) {
  if (!update || !update.message) return;

  var msg = update.message;
  var chatId = msg.chat.id;
  var text = (msg.text || '').trim();

  // 1. Xử lý các lệnh dạng Text
  if (text.startsWith('/start') || text.startsWith('/help') || text.startsWith('/huongdan')) {
    sendWelcomeMessage(chatId);
    return;
  }

  if (text.startsWith('/reset') || text.startsWith('/clear')) {
    clearUserSession(chatId);
    sendTelegramMessage(chatId, '🔄 <b>Đã làm mới phiên làm việc!</b>\nBạn hãy gửi 2 file báo cáo Excel (File Siêu Thị BHX và File CVS & NPP) để bắt đầu tạo báo cáo mới.');
    return;
  }

  if (text.startsWith('/status')) {
    checkUserSessionStatus(chatId);
    return;
  }

  // 2. Xử lý khi người dùng gửi File (Document)
  if (msg.document) {
    handleIncomingDocument(chatId, msg.document);
    return;
  }

  // Phản hồi mặc định nếu gửi tin nhắn thường
  sendTelegramMessage(chatId, '💡 <i>Vui lòng gửi 2 file Excel dữ liệu doanh số (.xlsx, .xlsb, .csv) để tôi tạo báo cáo cho bạn, hoặc gõ /help để xem hướng dẫn.</i>');
}

/**
 * Gửi tin nhắn hướng dẫn ban đầu
 */
function sendWelcomeMessage(chatId) {
  var welcomeText = 
    '📊 <b>HỆ THỐNG BÁO CÁO DOANH SỐ TỰ ĐỘNG (CVS & BHX)</b>\n' +
    '<i>Tác giả: Team Lead Trần Thị Cẩm Giang</i>\n\n' +
    '🤖 <b>Chào bạn! Tôi là Bot tự động tổng hợp báo cáo.</b>\n\n' +
    '👉 <b>CÁCH SỬ DỤNG:</b>\n' +
    'Gửi cho tôi <b>2 file Excel doanh số</b> (có thể gửi lần lượt hoặc gửi cùng lúc):\n' +
    '1️⃣ <b>File Siêu Thị:</b> Chứa dữ liệu BHX 5 Hubs + WinMart+\n' +
    '2️⃣ <b>File CVS & NPP:</b> Chứa dữ liệu GS25, 7-Eleven, Circle K, FamilyMart, Hoàng Đức\n\n' +
    '⚡ Bot sẽ tự động nhận diện, bóc tách số liệu, tạo Google Sheet 6 sheets chuẩn format và gửi ngược lại cho bạn <b>File Excel (.xlsx)</b> hoàn chỉnh kèm Link xem trực tiếp!\n\n' +
    '📌 <b>Lệnh nhanh:</b>\n' +
    '/status - Xem trạng thái file đã gửi\n' +
    '/reset - Xóa phiên cũ để nạp lại từ đầu';

  sendTelegramMessage(chatId, welcomeText);
}

/**
 * Xử lý khi nhận được File Document từ Telegram
 */
function handleIncomingDocument(chatId, doc) {
  var fileName = doc.file_name || 'unknown.xlsx';
  var fileId = doc.file_id;
  var fileSizeMB = (doc.file_size || 0) / (1024 * 1024);

  if (fileSizeMB > 25) {
    sendTelegramMessage(chatId, '⚠️ File <b>' + fileName + '</b> vượt quá giới hạn 25MB của Telegram Bot.');
    return;
  }

  // Báo trạng thái đang nhận diện
  sendTelegramChatAction(chatId, 'typing');

  try {
    // 1. Nhận diện sơ bộ theo tên file trước (cực nhanh <10ms, không tốn thời gian tải file nặng)
    var detectedType = detectFileType(fileName, null);

    // Nếu tên file không rõ ràng, mới phải tải file về để xem cấu trúc sheet
    if (detectedType === 'UNKNOWN') {
      var blob = downloadTelegramFile(fileId);
      blob.setName(fileName);
      var wb = parseWorkbook(blob);
      detectedType = detectFileType(fileName, wb);
    }

    if (detectedType === 'UNKNOWN') {
      sendTelegramMessage(chatId, 
        '❓ <b>Chưa xác định được loại file:</b> ' + fileName + '\n' +
        'File cần là <b>File Siêu Thị (BHX)</b> hoặc <b>File CVS & NPP</b> (có sheet SO, PO hoặc NPP).\n' +
        'Vui lòng kiểm tra lại file của bạn.');
      return;
    }

    // 2. Quản lý phiên trong CacheService
    var cache = CacheService.getScriptCache();
    var sessionKey = 'tg_session_' + chatId;
    var sessionData = {};

    var cachedStr = cache.get(sessionKey);
    if (cachedStr) {
      try { sessionData = JSON.parse(cachedStr); } catch (e) {}
    }

    // Lưu thông tin file vừa nhận
    if (detectedType === 'ST') {
      sessionData.fileST = {
        file_id: fileId,
        file_name: fileName
      };
    } else if (detectedType === 'CVS') {
      sessionData.fileCVS = {
        file_id: fileId,
        file_name: fileName
      };
    }

    // Kiểm tra xem đã đủ cả 2 file chưa
    if (sessionData.fileST && sessionData.fileCVS) {
      // Đã đủ 2 file -> Gửi thông báo ngay để người dùng yên tâm
      sendTelegramMessage(chatId, 
        '✨ <b>ĐÃ NHẬN ĐỦ 2 FILE DỮ LIỆU!</b>\n' +
        '1️⃣ 🏬 <b>File Siêu Thị:</b> ' + sessionData.fileST.file_name + '\n' +
        '2️⃣ 🏪 <b>File CVS &amp; NPP:</b> ' + sessionData.fileCVS.file_name + '\n\n' +
        '⏳ <i>Hệ thống đang tiến hành bóc tách số liệu và tạo báo cáo chuẩn Team Cẩm Giang... Vui lòng đợi trong khoảng 15-25 giây!</i>');

      // Kích hoạt tạo báo cáo
      generateAndSendReport(chatId, sessionData.fileST, sessionData.fileCVS);

      // Xóa session sau khi đã hoàn thành
      cache.remove(sessionKey);

    } else {
      // Mới nhận được 1 trong 2 file -> Lưu session và nhắc người dùng gửi file còn lại
      cache.put(sessionKey, JSON.stringify(sessionData), 7200); // Lưu 2 tiếng

      var receivedTitle = detectedType === 'ST' ? '🏬 File Siêu Thị (BHX &amp; WinMart)' : '🏪 File CVS &amp; NPP (GS25, 7E, CK, FM)';
      var missingTitle = detectedType === 'ST' ? '🏪 File CVS &amp; NPP' : '🏬 File Siêu Thị (BHX)';

      sendTelegramMessage(chatId, 
        '✅ <b>Đã nhận diện:</b> ' + receivedTitle + '\n' +
        '📄 <i>' + fileName + '</i>\n\n' +
        '⏳ <b>Còn thiếu:</b> ' + missingTitle + '\n' +
        '👉 <i>Vui lòng gửi tiếp file thứ 2 để Bot hoàn thiện báo cáo!</i>');
    }

  } catch (err) {
    Logger.log('Lỗi xử lý file: ' + err.toString());
    var safeErr = String(err.message || err).replace(/[<>&]/g, '');
    sendTelegramMessage(chatId, '❌ <b>Lỗi khi đọc file:</b> ' + safeErr);
  }
}

/**
 * Kiểm tra trạng thái hiện tại của phiên làm việc
 */
function checkUserSessionStatus(chatId) {
  var cache = CacheService.getScriptCache();
  var cachedStr = cache.get('tg_session_' + chatId);
  if (!cachedStr) {
    sendTelegramMessage(chatId, 'ℹ️ Hiện tại bạn chưa gửi file nào. Hãy gửi File Siêu Thị (BHX) và File CVS & NPP để bắt đầu.');
    return;
  }

  var s = JSON.parse(cachedStr);
  var text = '📋 <b>TRẠNG THÁI HIỆN TẠI:</b>\n';
  text += s.fileST ? ('• 🏬 File Siêu Thị: ✅ <i>' + s.fileST.file_name + '</i>\n') : '• 🏬 File Siêu Thị: ❌ Chưa nhận\n';
  text += s.fileCVS ? ('• 🏪 File CVS & NPP: ✅ <i>' + s.fileCVS.file_name + '</i>\n') : '• 🏪 File CVS & NPP: ❌ Chưa nhận\n';

  if (!s.fileST || !s.fileCVS) {
    text += '\n👉 Vui lòng gửi file còn thiếu để tiến hành tổng hợp!';
  }
  sendTelegramMessage(chatId, text);
}

/**
 * Xóa phiên làm việc của người dùng
 */
function clearUserSession(chatId) {
  CacheService.getScriptCache().remove('tg_session_' + chatId);
}

/**
 * ============================================================================
 * QUY TRÌNH XỬ LÝ VÀ XUẤT BÁO CÁO GỬI VỀ TELEGRAM
 * ============================================================================
 */
function generateAndSendReport(chatId, fileSTInfo, fileCVSInfo) {
  try {
    sendTelegramChatAction(chatId, 'upload_document');

    // 1. Tải 2 file nhị phân từ Telegram
    var blobST = downloadTelegramFile(fileSTInfo.file_id);
    blobST.setName(fileSTInfo.file_name);

    var blobCVS = downloadTelegramFile(fileCVSInfo.file_id);
    blobCVS.setName(fileCVSInfo.file_name);

    // 2. Parse Workbook
    var wbST = parseWorkbook(blobST);
    var wbCVS = parseWorkbook(blobCVS);

    // 3. Bóc tách và tính toán các chỉ số
    var customMasterRes = getMasterDataFromBackend();
    var customMaster = (customMasterRes && customMasterRes.success) ? customMasterRes.data : null;
    var payload = processTwoWorkbooks(wbST, wbCVS, customMaster);

    // 4. Tạo Google Spreadsheet 6 sheet trên Google Drive
    var sheetResult = createGoogleSheetReport(payload);
    if (!sheetResult || !sheetResult.success) {
      throw new Error('Lỗi khi tạo Google Sheet: ' + (sheetResult ? sheetResult.error : 'Không rõ nguyên nhân'));
    }

    var ssId = sheetResult.id;
    var ssUrl = sheetResult.url;
    var ssName = sheetResult.name;

    // 5. Xuất Google Sheet thành file Excel (.xlsx) nhị phân
    var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
    var exportRes = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    var xlsxBlob = exportRes.getBlob().setName(ssName + '.xlsx');

    // 6. Chuẩn bị Caption và Gửi File Excel qua Telegram
    var caption = 
      '📊 <b>BÁO CÁO DOANH SỐ TEAM ' + payload.team_lead.toUpperCase() + '</b>\n' +
      'Tháng ' + payload.month + '/' + payload.year + ' (% Timegone: ' + payload.timegone + '%)\n' +
      '• Target: ' + (payload.total_target || 0).toLocaleString('vi-VN') + ' ₫\n' +
      '• Tổng Thực Hiện: ' + (payload.total_actual || 0).toLocaleString('vi-VN') + ' ₫ (' + payload.percent_achieved + '%)\n' +
      '• Tình trạng: ' + (payload.percent_achieved >= payload.timegone ? '🟢 VƯỢT' : (payload.percent_achieved >= payload.timegone * 0.75 ? '🟠 CẬN' : '🔴 CHẬM')) + ' TIẾN ĐỘ';

    // 6.1 Tự động tạo và gửi Ảnh biểu đồ tiến độ & bảng xếp hạng trực quan
    try {
      var chartBlob = generateRankingChartImage(payload);
      if (chartBlob) {
        sendTelegramPhoto(chatId, chartBlob, '📊 <b>ẢNH BIỂU ĐỒ TIẾN ĐỘ & XẾP HẠNG DOANH SỐ TEAM ' + payload.team_lead.toUpperCase() + '</b>\nTháng ' + payload.month + '/' + payload.year + ' | ⏳ % Timegone: ' + payload.timegone + '% | 🎯 % Đạt Team: ' + payload.percent_achieved + '%');
      }
    } catch (eImg) {
      Logger.log('Bỏ qua lỗi ảnh biểu đồ: ' + eImg);
    }

    sendTelegramDocument(chatId, xlsxBlob, caption);

    // 7. Gửi Tin Nhắn Tóm Tắt Chi Tiết & Link Google Sheets
    var topEmployees = (payload.employees || []).slice().sort(function(a, b) {
      return (b.percent || 0) - (a.percent || 0);
    });

    var summaryMsg = 
      '🎯 <b>TỔNG HỢP DOANH SỐ CHI TIẾT:</b>\n' +
      '━━━━━━━━━━━━━━━━━━━━\n' +
      '• <b>Target Team:</b> ' + (payload.total_target || 0).toLocaleString('vi-VN') + ' ₫\n' +
      '• <b>Thực Hiện BHX:</b> ' + (payload.total_actual_bhx || 0).toLocaleString('vi-VN') + ' ₫ (TB ' + (payload.bhx_per_store || 0).toLocaleString('vi-VN') + ' ₫/CH - ' + (payload.total_stores_bhx_team || 178) + '/' + (payload.total_stores_bhx || 232) + ' CH)\n' +
      '• <b>Thực Hiện CVS:</b> ' + (payload.total_cvs || 0).toLocaleString('vi-VN') + ' ₫\n' +
      '• <b>Tổng Thực Hiện:</b> ' + (payload.total_actual || 0).toLocaleString('vi-VN') + ' ₫\n' +
      '• <b>% Đạt Team:</b> <b>' + payload.percent_achieved + '%</b> (vs Timegone ' + payload.timegone + '%)\n' +
      '━━━━━━━━━━━━━━━━━━━━\n' +
      '🏆 <b>BẢNG XẾP HẠNG NHÂN VIÊN:</b>\n';

    for (var i = 0; i < Math.min(10, topEmployees.length); i++) {
      var emp = topEmployees[i];
      var medal = i === 0 ? '🥇' : (i === 1 ? '🥈' : (i === 2 ? '🥉' : (i + 1) + '.'));
      var statusIcon = emp.percent >= payload.timegone ? '🟢' : (emp.percent >= payload.timegone * 0.75 ? '🟠' : '🔴');
      summaryMsg += medal + ' ' + statusIcon + ' <b>' + emp.name + '</b>: ' + emp.percent + '% (' + (emp.total_actual || 0).toLocaleString('vi-VN') + ' ₫)\n';
    }

    summaryMsg += 
      '━━━━━━━━━━━━━━━━━━━━\n' +
      '🔗 <b>Xem trực tiếp Google Sheet Online:</b>\n' + ssUrl + '\n\n' +
      '<i>Đã đính kèm file Excel (.xlsx) ở tin nhắn phía trên để bạn tải về hoặc chuyển tiếp cho quản lý!</i>';

    sendTelegramMessage(chatId, summaryMsg);

  } catch (err) {
    Logger.log('Lỗi tạo và gửi báo cáo: ' + err.toString());
    sendTelegramMessage(chatId, '❌ <b>Đã xảy ra lỗi trong quá trình tạo báo cáo:</b>\n' + err.message);
  }
}

/**
 * ============================================================================
 * CÁC HÀM GỌI TELEGRAM BOT API
 * ============================================================================
 */

/**
 * Gửi tin nhắn Text (hỗ trợ định dạng HTML)
 */
/**
 * Gửi tin nhắn Text (hỗ trợ định dạng HTML, tự động fallback nếu lỗi cú pháp)
 */
function sendTelegramMessage(chatId, text) {
  var token = getTelegramBotToken();
  var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  var payload = {
    chat_id: String(chatId),
    text: text,
    parse_mode: 'HTML',
    disable_web_page_preview: false
  };

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var resText = res.getContentText();
  try {
    var json = JSON.parse(resText);
    if (!json.ok) {
      console.error('❌ Lỗi gửi tin nhắn Telegram: ' + resText);
      // Tự động bỏ thẻ HTML và gửi lại dạng Plain Text
      delete payload.parse_mode;
      payload.text = text.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&');
      return UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
    }
  } catch (e) {
    console.error('Lỗi parse phản hồi Telegram: ' + e);
  }

  return res;
}

/**
 * Gửi hành động chat (ví dụ: 'typing', 'upload_document')
 */
function sendTelegramChatAction(chatId, action) {
  var token = getTelegramBotToken();
  var url = 'https://api.telegram.org/bot' + token + '/sendChatAction';
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: String(chatId), action: action || 'typing' }),
    muteHttpExceptions: true
  });
}

/**
 * Gửi File Document (File Excel .xlsx) về cho người dùng
 */
function sendTelegramDocument(chatId, blob, caption) {
  var token = getTelegramBotToken();
  var url = 'https://api.telegram.org/bot' + token + '/sendDocument';

  var payload = {
    chat_id: String(chatId),
    caption: caption || '',
    parse_mode: 'HTML',
    document: blob
  };

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });

  var resText = res.getContentText();
  try {
    var json = JSON.parse(resText);
    if (!json.ok) {
      console.error('❌ Lỗi gửi file Telegram: ' + resText);
      // Gửi lại không có HTML parse mode nếu lỗi entity
      delete payload.parse_mode;
      payload.caption = (caption || '').replace(/<[^>]*>/g, '').replace(/&amp;/g, '&');
      return UrlFetchApp.fetch(url, {
        method: 'post',
        payload: payload,
        muteHttpExceptions: true
      });
    }
  } catch (e) {
    console.error('Lỗi parse phản hồi gửi file: ' + e);
  }

  return res;
}

/**
 * Gửi ảnh (Photo) qua Telegram Bot
 */
function sendTelegramPhoto(chatId, blob, caption) {
  var token = getTelegramBotToken();
  var url = 'https://api.telegram.org/bot' + token + '/sendPhoto';

  var payload = {
    chat_id: String(chatId),
    caption: caption || '',
    parse_mode: 'HTML',
    photo: blob
  };

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });

  return res;
}

/**
 * Tạo ảnh biểu đồ thanh ngang Top 10 nhân viên & tiến độ team từ QuickChart API
 */
function generateRankingChartImage(payload) {
  try {
    var emps = (payload.employees || []).slice().sort(function(a, b) {
      return (b.percent || 0) - (a.percent || 0);
    });

    var labels = emps.map(function(e) { return e.name; });
    var pcts = emps.map(function(e) { return e.percent || 0; });
    var timegone = payload.timegone || 0;

    var chartConfig = {
      type: 'horizontalBar',
      data: {
        labels: labels,
        datasets: [{
          label: '% Đạt Doanh Số',
          data: pcts,
          backgroundColor: pcts.map(function(p) {
            return p >= timegone ? 'rgba(16, 185, 129, 0.85)' : (p >= timegone * 0.75 ? 'rgba(245, 158, 11, 0.85)' : 'rgba(239, 68, 68, 0.85)');
          }),
          borderColor: pcts.map(function(p) {
            return p >= timegone ? '#059669' : (p >= timegone * 0.75 ? '#D97706' : '#DC2626');
          }),
          borderWidth: 1.5
        }]
      },
      options: {
        title: {
          display: true,
          text: 'TIẾN ĐỘ DOANH SỐ TEAM ' + payload.team_lead.toUpperCase() + ' (Tháng ' + payload.month + '/' + payload.year + ' - Timegone: ' + timegone + '%)',
          fontSize: 16,
          fontColor: '#1E293B',
          fontStyle: 'bold'
        },
        legend: { display: false },
        scales: {
          xAxes: [{
            ticks: {
              beginAtZero: true,
              callback: function(val) { return val + '%'; }
            },
            gridLines: { color: 'rgba(226, 232, 240, 0.8)' }
          }],
          yAxes: [{
            ticks: { fontStyle: 'bold', fontColor: '#334155' },
            gridLines: { display: false }
          }]
        }
      }
    };

    var chartUrl = 'https://quickchart.io/chart?w=850&h=480&bkg=white&devicePixelRatio=2&c=' + encodeURIComponent(JSON.stringify(chartConfig));
    var res = UrlFetchApp.fetch(chartUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      return res.getBlob().setName('Tien_Do_Doanh_So_Team_' + payload.month + '_' + payload.year + '.png');
    }
  } catch (e) {
    Logger.log('Lỗi tạo ảnh biểu đồ QuickChart: ' + e);
  }
  return null;
}

/**
 * Lấy URL tải file từ Telegram dựa theo file_id
 */
function getTelegramFileUrl(fileId) {
  var token = getTelegramBotToken();
  var url = 'https://api.telegram.org/bot' + token + '/getFile?file_id=' + fileId;
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var data = JSON.parse(res.getContentText());

  if (data.ok && data.result && data.result.file_path) {
    return 'https://api.telegram.org/file/bot' + token + '/' + data.result.file_path;
  }
  throw new Error('Không thể lấy đường dẫn file từ Telegram: ' + (data.description || ''));
}

/**
 * Tải file từ Telegram về dưới dạng Google Apps Script Blob
 */
function downloadTelegramFile(fileId) {
  var fileUrl = getTelegramFileUrl(fileId);
  var res = UrlFetchApp.fetch(fileUrl, { muteHttpExceptions: true });
  return res.getBlob();
}
