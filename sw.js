/**
 * sw.js — 天線選型 PWA 的 Service Worker
 *
 * 策略:
 *  - App shell (index.html、manifest、icons)：安裝時預先快取，之後
 *    stale-while-revalidate (先回快取、背景更新)，確保離線可用。
 *  - Google Fonts：runtime cache-first，離線時退回系統字型。
 *
 * 更新資料時：重新執行 build_data.py 產生 index.html，並遞增 CACHE_VERSION，
 * 使用者下次開啟即會取得新版。
 */
const CACHE_VERSION = 'antenna-selector-v1';
const FONT_CACHE = 'antenna-selector-fonts-v1';

const APP_SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(APP_SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  // 清除舊版快取
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== CACHE_VERSION && k !== FONT_CACHE).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Google Fonts：cache-first
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(
      caches.open(FONT_CACHE).then(async (cache) => {
        const hit = await cache.match(req);
        if (hit) return hit;
        try {
          const res = await fetch(req);
          if (res.ok || res.type === 'opaque') cache.put(req, res.clone());
          return res;
        } catch {
          return new Response('', { status: 504 });
        }
      })
    );
    return;
  }

  // 同源資源：stale-while-revalidate
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.open(CACHE_VERSION).then(async (cache) => {
        const cached = await cache.match(req, { ignoreSearch: true });
        const network = fetch(req)
          .then((res) => { if (res.ok) cache.put(req, res.clone()); return res; })
          .catch(() => null);
        if (cached) { event.waitUntil(network); return cached; }
        const res = await network;
        return res || (await cache.match('./index.html')) || new Response('Offline', { status: 503 });
      })
    );
  }
});
