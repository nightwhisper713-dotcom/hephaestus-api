// ============================================================================
//  幻翔 — 圖片生成後端代理  (/generate/image)
//  目的：讓客戶前端 (CALLIOPE 輪播圖) 生成圖片時，OpenAI Key 只留在伺服器端，
//        絕不出現在瀏覽器。前端只送 { companyId, prompt, size, quality }。
//
//  部署位置：加進 hephaestus-api（Render）的 Express app。
//  需求：Render 環境變數  OPENAI_API_KEY = sk-...   （在 Render → Environment 設定）
//  Node 18+ 內建 fetch；若是舊版 Node，請 `npm i node-fetch` 並取消最下方註解。
// ============================================================================

// const fetch = (...a) => import('node-fetch').then(({default: f}) => f(...a)); // 舊版 Node 才需要

/**
 * 用法 A（最簡單）：在 server.js 直接掛上這個 router
 *   const imageProxy = require('./hephaestus-image-proxy');
 *   app.use(imageProxy);            // 提供 POST /generate/image
 *
 * 用法 B：若你已有自己的 router，把下面 handler 的內容複製進你的路由即可。
 *
 * 注意：確保你的 CORS 設定允許客戶前端網域（Netlify / GitHub Pages）。
 *       hephaestus-api 既有的 /generate/ppt 等端點能跨網域運作，代表 CORS 已設好；
 *       這個端點沿用同一套即可。
 */

const express = require('express');
const router = express.Router();

// 確保能解析 JSON（若 server.js 已有 app.use(express.json()) 可省略）
router.use(express.json({ limit: '1mb' }));

router.post('/generate/image', async (req, res) => {
  try {
    const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
    if (!OPENAI_API_KEY) {
      return res.status(500).json({ error: '伺服器未設定 OPENAI_API_KEY 環境變數' });
    }

    const { prompt, size, quality, companyId } = req.body || {};
    if (!prompt || typeof prompt !== 'string') {
      return res.status(400).json({ error: '缺少 prompt' });
    }
    // 基本防濫用：限制長度
    const safePrompt = prompt.slice(0, 4000);
    const safeSize = ['1024x1536', '1024x1024', '1536x1024'].includes(size) ? size : '1024x1536';
    const safeQuality = ['low', 'medium', 'high'].includes(quality) ? quality : 'low';

    // （可選）在這裡加入身分驗證：
    //   - 驗證 req.headers.authorization 內的 Supabase JWT，或
    //   - 用 service-role key 查 companyId 是否存在 / 是否超額。
    //   目前先以「Key 不外洩」為首要目標；驗證可後續加強。
    void companyId;

    const oaResp = await fetch('https://api.openai.com/v1/images/generations', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + OPENAI_API_KEY,
      },
      body: JSON.stringify({
        model: 'gpt-image-1',
        prompt: safePrompt,
        n: 1,
        size: safeSize,
        quality: safeQuality,
      }),
    });

    const data = await oaResp.json().catch(() => ({}));
    if (!oaResp.ok || data.error) {
      const msg = (data.error && (data.error.message || data.error)) || ('OpenAI HTTP ' + oaResp.status);
      return res.status(502).json({ error: msg });
    }

    const b64 = data.data && data.data[0] && data.data[0].b64_json;
    if (!b64) return res.status(502).json({ error: 'OpenAI 未回傳圖片資料' });

    // 只回傳圖片，不回傳任何金鑰
    return res.json({ b64_json: b64 });
  } catch (e) {
    return res.status(500).json({ error: (e && e.message) || '圖片代理失敗' });
  }
});

module.exports = router;
