// api/shorten.js - Vercel Serverless Function to create clean, INSTANT redirect short URLs
// NO interstitial pages, NO 10-second countdowns, 100% direct 301/302 redirects.
module.exports = async (req, res) => {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  let targetUrl = '';
  if (req.method === 'GET') {
    targetUrl = req.query.url || '';
  } else {
    let body = req.body;
    if (typeof body === 'string') {
      try { body = JSON.parse(body); } catch (e) {}
    }
    targetUrl = (body && body.url) ? body.url : (req.query.url || '');
  }

  if (!targetUrl || !targetUrl.startsWith('http')) {
    return res.status(400).json({ ok: false, error: 'Tham số url không hợp lệ' });
  }

  // 1. Primary: da.gd (Instant 302 redirect, 0s countdown, ~18 chars)
  try {
    const dagdRes = await fetch(`https://da.gd/s?url=${encodeURIComponent(targetUrl)}`, {
      signal: AbortSignal.timeout(4000)
    });
    if (dagdRes.ok) {
      const short = (await dagdRes.text()).trim();
      if (short && short.startsWith('http')) {
        return res.status(200).json({ ok: true, shortUrl: short, provider: 'da.gd' });
      }
    }
  } catch (eDagd) {
    console.warn('da.gd failed:', eDagd.message);
  }

  // 2. Secondary fallback: cleanuri.com (Instant 301 redirect, 0s countdown, ~26 chars)
  try {
    const cleanRes = await fetch('https://cleanuri.com/api/v1/shorten', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ url: targetUrl }),
      signal: AbortSignal.timeout(4000)
    });
    if (cleanRes.ok) {
      const data = await cleanRes.json();
      if (data && data.result_url && data.result_url.startsWith('http')) {
        return res.status(200).json({ ok: true, shortUrl: data.result_url, provider: 'cleanuri' });
      }
    }
  } catch (eClean) {
    console.warn('cleanuri failed:', eClean.message);
  }

  // Fallback to original long URL if all shorteners fail
  return res.status(200).json({ ok: false, shortUrl: targetUrl, warning: 'Không thể rút gọn, dùng link gốc' });
};
