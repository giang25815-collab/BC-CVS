const zlib = require('zlib');
const masterData = require('../master_data.js');

let allSkus = [];
if (masterData && masterData.fm_categories) {
  masterData.fm_categories.forEach(cat => allSkus.push(...cat.skus));
}

/**
 * Nén dữ liệu báo cáo thành chuỗi base64url siêu ngắn (~300 - 450 ký tự)
 * Tận dụng master_data có sẵn trong ứng dụng để chỉ lưu index và số liệu phát sinh.
 */
function encodeReportForShare(payload) {
  if (!payload) return '';

  // 1. Hubs (5 số)
  const h = (payload.bhx_hubs || []).map(hub => hub.amount !== undefined ? hub.amount : (hub.actual_sales || 0));

  // 2. Stores [ [storeIndex, sales], ... ]
  const s = [];
  (payload.stores || payload.cvs_stores || []).forEach((st, idx) => {
    const val = st.actual_sales || st.actual || 0;
    if (val > 0) {
      let storeIdx = idx;
      if (st.store_code && masterData && masterData.stores) {
        const found = masterData.stores.findIndex(ms => ms.store_code === st.store_code);
        if (found !== -1) storeIdx = found;
      }
      s.push([storeIdx, val]);
    }
  });

  // 3. FM matrix [ [fmStoreIndex, [ [skuIdx, amt], ... ]] ]
  const f = [];
  const fmStores = (payload.fm_sku_matrix && payload.fm_sku_matrix.stores) ? payload.fm_sku_matrix.stores : [];
  fmStores.forEach((st, sIdx) => {
    if (st.orders) {
      const ordersList = [];
      allSkus.forEach((sku, skuIdx) => {
        const amt = st.orders[sku.code] || 0;
        if (amt > 0) {
          ordersList.push([skuIdx, amt]);
        }
      });
      if (ordersList.length > 0) {
        f.push([sIdx, ordersList]);
      }
    }
  });

  const compactObj = {
    v: 2,
    m: payload.month || (new Date().getMonth() + 1),
    y: payload.year || new Date().getFullYear(),
    tg: payload.timegone || 0,
    h, s, f
  };

  const jsonStr = JSON.stringify(compactObj);
  const compressed = zlib.deflateSync(Buffer.from(jsonStr, 'utf8'));
  return compressed.toString('base64url');
}

/**
 * Lưu báo cáo vào Cloud Storage (Giải pháp 2C) để sinh đường link siêu ngắn, chuyên nghiệp
 * Ví dụ: https://bcgiang.vercel.app/r/DGD3mhlqTL
 */
async function generateCleanReportUrl(req, payload) {
  const host = (req && req.headers) ? (req.headers['x-forwarded-host'] || req.headers.host) : 'bcgiang.vercel.app';
  const proto = (req && req.headers) ? (req.headers['x-forwarded-proto'] || 'https') : 'https';
  const baseUrl = `${proto}://${host}`;

  // 1. Thử lưu vào Upstash Redis / Vercel KV nếu có cấu hình
  const kvUrl = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const kvToken = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if (kvUrl && kvToken) {
    try {
      const m = payload.month || (new Date().getMonth() + 1);
      const y = payload.year || new Date().getFullYear();
      const rand = Math.random().toString(36).substring(2, 6).toUpperCase();
      const customId = `CG${String(y).slice(-2)}${String(m).padStart(2, '0')}-${rand}`;

      const kvRes = await fetch(`${kvUrl}/set/${customId}`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${kvToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload),
        signal: AbortSignal.timeout(4000)
      });
      if (kvRes.ok) {
        return `${baseUrl}/r/${customId}`;
      }
    } catch (eKv) {
      console.warn('KV save failed, trying Cloud storage fallback:', eKv.message);
    }
  }

  // 2. Lưu vào Cloud Storage Bytebin (Global CDN, 0-config, ID 10 ký tự)
  try {
    const bbRes = await fetch('https://bytebin.lucko.me/post', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'User-Agent': 'SalesReportApp/2.0'
      },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(5000)
    });
    if (bbRes.ok) {
      const bbData = await bbRes.json();
      if (bbData && bbData.key) {
        return `${baseUrl}/r/${bbData.key}`;
      }
    }
  } catch (eBb) {
    console.warn('Bytebin cloud save failed:', eBb.message);
  }

  // 3. Fallback an toàn: Mã hóa nén trực tiếp vào URL nếu Cloud tạm thời không kết nối được
  const encoded = encodeReportForShare(payload);
  return `${baseUrl}/r#d=${encoded}`;
}

/**
 * Tự động tạo link rút gọn /r đẹp mắt trên Vercel theo domain hiện tại
 */
function generateReportUrl(req, payload) {
  const host = (req && req.headers) ? (req.headers['x-forwarded-host'] || req.headers.host) : 'bcgiang.vercel.app';
  const proto = (req && req.headers) ? (req.headers['x-forwarded-proto'] || 'https') : 'https';
  const baseUrl = `${proto}://${host}`;
  const encoded = encodeReportForShare(payload);
  return `${baseUrl}/r#d=${encoded}`;
}

module.exports = {
  encodeReportForShare,
  generateReportUrl,
  generateCleanReportUrl
};

