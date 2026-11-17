/**
 * api/proxy.mjs — Vercel Serverless Proxy → Google Apps Script Web App
 *
 * Tujuan:
 *   - Menghindari CORS (script.google.com tidak mengirim header CORS).
 *   - Menyembunyikan URL GAS dari publik.
 *   - Menyediakan endpoint stabil: /api/proxy?action=...
 *
 * Env variable wajib di Vercel:
 *   GAS_WEB_APP_URL = https://script.google.com/macros/s/XXXX/exec
 *
 * Pemakaian dari frontend:
 *   fetch('/api/proxy?action=getDashboardData')
 *   fetch('/api/proxy?action=getYearlyData&targetYear=2024')
 */

export default async function handler(req, res) {
  // CORS untuk pemakaian opsional dari origin lain (mis. preview deployments).
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Accept');

  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }

  const GAS_URL = process.env.GAS_WEB_APP_URL;
  if (!GAS_URL) {
    res.status(500).json({
      error: 'Konfigurasi server tidak lengkap: env GAS_WEB_APP_URL belum di-set di Vercel.'
    });
    return;
  }

  try {
    const target = new URL(GAS_URL);

    // 1. Teruskan query parameter
    const query = req.query || {};
    for (const key of Object.keys(query)) {
      const val = query[key];
      if (Array.isArray(val)) {
        for (const v of val) target.searchParams.append(key, v);
      } else if (val != null) {
        target.searchParams.set(key, val);
      }
    }

    // 2. Bangun request ke GAS
    const forwardHeaders = { Accept: 'application/json' };
    let body;
    if (req.method === 'POST') {
      forwardHeaders['Content-Type'] = 'application/json';
      body = JSON.stringify(req.body || {});
    }

    const gasRes = await fetch(target.toString(), {
      method: req.method === 'POST' ? 'POST' : 'GET',
      headers: forwardHeaders,
      body,
      redirect: 'follow'
    });

    const raw = await gasRes.text();

    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store, max-age=0');
    res.setHeader('X-Upstream-Status', String(gasRes.status));

    if (!gasRes.ok) {
      res.status(502).send(
        JSON.stringify({
          error: 'GAS upstream error (HTTP ' + gasRes.status + ')',
          detail: raw ? raw.slice(0, 500) : null
        })
      );
      return;
    }

    res.status(200).send(raw);
  } catch (err) {
    res.status(502).json({
      error: 'Proxy error: ' + (err && err.message ? err.message : String(err))
    });
  }
}