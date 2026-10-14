/**
 * sw.js — 天線選型 PWA 的 Service Worker (完全離線版)
 *
 * 策略:
 *  - 安裝時預先快取所有 App 檔案 (資料、圖片、字型皆內嵌於 index.html)，
 *    第一次開啟後即可完全離線使用。
 *  - 頁面導覽 (開啟 App)：有網路時優先載入最新版 (逾時 3.5 秒或離線則用快取)，
 *    發佈新版後第一次開啟即生效；使用說明的示範畫面直接使用快取。
 *  - 其他同源檔案：cache-first，背景更新。
 *  - 產品照片 (photos.json)：安裝時預先快取，離線也能放大檢視。
 *  - version.json：一律走網路，供 App 偵測新版本；離線時回 503。
 *  - 不使用任何外部網域資源。
 *
 * CACHE_VERSION 由 build_data.py --version 自動改寫，請勿手動修改。
 * App 的「強制更新」會先確認伺服器可連線，才清除快取並重新載入。
 */
/** 開啟 App 時等待網路的上限 (毫秒)；超過即改用快取，確保開啟速度 */
const NAV_TIMEOUT_MS = 3500;

const CACHE_VERSION = 'antenna-selector-v3.16.0';

const APP_SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png',
  './screenshot-wide.png',
  './screenshot-narrow.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_VERSION);
    // cache: 'reload' 確保拿到伺服器最新檔案，而非瀏覽器 HTTP 快取
    await cache.addAll(APP_SHELL.map((u) => new Request(u, { cache: 'reload' })));
    // 放大檢視用照片 (photos.json，由 build_data.py 產生)：一併預先快取以便離線檢視；
    // 檔案不存在或下載失敗時不影響安裝
    try { await cache.add(new Request('./photos.json', { cache: 'reload' })); } catch { /* 忽略 */ }
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

/** 背景更新單一資源 (失敗時安靜略過) */
async function refresh(cache, req, key) {
  try {
    const res = await fetch(req, { cache: 'no-cache' });
    if (res.ok) await cache.put(key || req, res.clone());
    return res;
  } catch {
    return null;
  }
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;   // 不處理外部請求

  // version.json：只走網路
  if (url.pathname.endsWith('/version.json')) {
    event.respondWith(
      fetch(req, { cache: 'no-store' }).catch(() => new Response('{}', { status: 503, headers: { 'Content-Type': 'application/json' } }))
    );
    return;
  }

  // 開啟 App (導覽請求)
  //  - 一般開啟：有網路時優先取最新版 (逾時 NAV_TIMEOUT_MS 改用快取)，離線時用快取
  //    → 發佈新版後，使用者「第一次」開啟就會看到新版
  //  - 使用說明內的示範畫面 (?embed=)：直接用快取，避免重複下載
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_VERSION);
      const cached = (await cache.match('./index.html')) || (await cache.match('./'));
      if (url.searchParams.has('embed') && cached) return cached;
      const network = refresh(cache, new Request('./index.html'), './index.html');
      if (!cached) {
        const res = await network;
        return res || new Response('<h1>Offline</h1><p>請先在有網路時開啟一次。</p>', {
          status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        });
      }
      const timeout = new Promise((r) => setTimeout(() => r(null), NAV_TIMEOUT_MS));
      const res = await Promise.race([network, timeout]);
      if (res && res.ok) return res;
      event.waitUntil(network);           // 逾時：先給快取，背景繼續更新
      return cached;
    })());
    return;
  }

  // 其他同源檔案：cache-first + 背景更新
  event.respondWith((async () => {
    const cache = await caches.open(CACHE_VERSION);
    const cached = await cache.match(req, { ignoreSearch: true });
    if (cached) {
      event.waitUntil(refresh(cache, req));
      return cached;
    }
    const res = await refresh(cache, req);
    return res || new Response('', { status: 504 });
  })());
});

// App 要求立即啟用新版 Service Worker
self.addEventListener('message', (event) => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});
