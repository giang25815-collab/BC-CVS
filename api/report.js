// api/report.js - Cloud Storage Service for Ultra-Short Executive Report Links (Solution 2C)
// Provides instant save & fetch for reports, generating clean URLs like: https://bcgiang.vercel.app/r/DGD3mhlqTL

const zlib = require('zlib');

module.exports = async (req, res) => {
  // CORS Headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  // Determine domain/origin
  const host = req.headers['x-forwarded-host'] || req.headers.host || 'bcgiang.vercel.app';
  const proto = req.headers['x-forwarded-proto'] || 'https';
  const baseUrl = `${proto}://${host}`;

  // =========================================================================
  // 1. GET: Lấy dữ liệu báo cáo theo ID
  // =========================================================================
  if (req.method === 'GET') {
    const id = (req.query.id || '').trim();
    if (!id) {
      return res.status(400).json({ ok: false, error: 'Thiếu tham số id' });
    }

    // 1.1 Kiểm tra Upstash Redis / Vercel KV nếu có cấu hình
    const kvUrl = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
    const kvToken = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
    if (kvUrl && kvToken) {
      try {
        const kvRes = await fetch(`${kvUrl}/get/${encodeURIComponent(id)}`, {
          headers: { Authorization: `Bearer ${kvToken}` },
          signal: AbortSignal.timeout(4000)
        });
        if (kvRes.ok) {
          const kvData = await kvRes.json();
          if (kvData && kvData.result) {
            let parsed = kvData.result;
            if (typeof parsed === 'string') {
              try { parsed = JSON.parse(parsed); } catch (e) {}
            }
            return res.status(200).json({ ok: true, data: parsed, source: 'kv' });
          }
        }
      } catch (eKv) {
        console.warn('Lỗi đọc từ KV:', eKv.message);
      }
    }

    // 1.2 Đọc từ Cloud Storage Bytebin (Global CDN)
    try {
      const bbRes = await fetch(`https://bytebin.lucko.me/${encodeURIComponent(id)}`, {
        signal: AbortSignal.timeout(5000)
      });
      if (bbRes.ok) {
        const data = await bbRes.json();
        return res.status(200).json({ ok: true, data: data, source: 'cloud' });
      }
    } catch (eBb) {
      console.warn('Lỗi đọc từ Bytebin:', eBb.message);
    }

    return res.status(404).json({ ok: false, error: 'Không tìm thấy báo cáo hoặc mã đã hết hạn' });
  }

  // =========================================================================
  // 2. POST: Lưu trữ dữ liệu báo cáo và sinh mã ID siêu ngắn
  // =========================================================================
  if (req.method === 'POST') {
    let body = req.body;
    if (typeof body === 'string') {
      try { body = JSON.parse(body); } catch (e) {}
    }

    if (!body || (!body.payload && !body.data && !body.compact)) {
      return res.status(400).json({ ok: false, error: 'Dữ liệu báo cáo không hợp lệ' });
    }

    const reportContent = body.payload || body.data || body.compact;

    // 2.1 Nếu có Vercel KV / Upstash, lưu với ID tùy biến theo tháng/năm
    const kvUrl = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
    const kvToken = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;

    if (kvUrl && kvToken) {
      try {
        const m = reportContent.month || (new Date().getMonth() + 1);
        const y = reportContent.year || new Date().getFullYear();
        const rand = Math.random().toString(36).substring(2, 6).toUpperCase();
        const customId = `CG${String(y).slice(-2)}${String(m).padStart(2, '0')}-${rand}`;

        const saveRes = await fetch(`${kvUrl}/set/${customId}`, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${kvToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(typeof reportContent === 'string' ? reportContent : JSON.stringify(reportContent)),
          signal: AbortSignal.timeout(5000)
        });

        if (saveRes.ok) {
          const cleanUrl = `${baseUrl}/r/${customId}`;
          return res.status(200).json({
            ok: true,
            id: customId,
            url: cleanUrl,
            shortUrl: cleanUrl,
            source: 'kv'
          });
        }
      } catch (eKvSave) {
        console.warn('Lỗi lưu vào KV, chuyển sang Cloud Storage:', eKvSave.message);
      }
    }

    // 2.2 Lưu vào Cloud Storage Bytebin (Siêu nhanh, ID 10 ký tự)
    try {
      const bbRes = await fetch('https://bytebin.lucko.me/post', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'User-Agent': 'SalesReportApp/2.0'
        },
        body: JSON.stringify(reportContent),
        signal: AbortSignal.timeout(6000)
      });

      if (bbRes.ok) {
        const bbData = await bbRes.json();
        if (bbData && bbData.key) {
          const reportId = bbData.key;
          const cleanUrl = `${baseUrl}/r/${reportId}`;
          return res.status(200).json({
            ok: true,
            id: reportId,
            url: cleanUrl,
            shortUrl: cleanUrl,
            source: 'cloud'
          });
        }
      }
    } catch (eBbSave) {
      console.warn('Lỗi lưu vào Bytebin:', eBbSave.message);
    }

    return res.status(500).json({
      ok: false,
      error: 'Không thể lưu báo cáo lên Cloud Storage lúc này'
    });
  }

  return res.status(405).json({ error: 'Method Not Allowed' });
};
