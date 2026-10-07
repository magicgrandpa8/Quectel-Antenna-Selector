/**
 * sw.js — 天線選型 PWA 的 Service Worker (完全離線版)
 *
 * 策略:
 *  - 安裝時預先快取所有 App 檔案 (資料、圖片、字型皆內嵌於 index.html)，
 *    第一次開啟後即可完全離線使用。
 *  - 頁面導覽 (開啟 App)：一律先回快取的 index.html (秒開、離線可用)，
 *    有網路時於背景更新快取，下次開啟生效。
 *  - 其他同源檔案：cache-first，背景更新。
 *  - version.json：一律走網路，供 App 偵測新版本；離線時回 503。
 *  - 不使用任何外部網域資源。
 *
 * CACHE_VERSION 由 build_data.py --version 自動改寫，請勿手動修改。
 * App 的「強制更新」會先確認伺服器可連線，才清除快取並重新載入。
 */
const CACHE_VERSION = 'antenna-selector-v3.8.0';

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

  // 開啟 App (導覽請求)：回快取的 index.html，背景更新
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_VERSION);
      const cached = (await cache.match('./index.html')) || (await cache.match('./'));
      if (cached) {
        event.waitUntil(refresh(cache, new Request('./index.html'), './index.html'));
        return cached;
      }
      const res = await refresh(cache, new Request('./index.html'), './index.html');
      return res || new Response('<h1>Offline</h1><p>請先在有網路時開啟一次。</p>', {
        status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' },
      });
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
