/* RAB Rumah & Gedung — Service Worker (by DG) */
const VERSION = 'rab-v1.3.1';
const SHELL = [
  './', 'index.html', 'style.css', 'calc.js', 'app.js', 'manifest.json', 'favicon.ico',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon.svg', 'icons/apple-touch-icon.png',
  'logo/logo-light.png', 'logo/logo-dark.png',
  'vendor/jspdf.umd.min.js', 'vendor/jspdf.plugin.autotable.min.js', 'vendor/xlsx.full.min.js'
];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  // API Apps Script & Google Drive: selalu langsung ke jaringan
  if (/script\.google(usercontent)?\.com|googleapis\.com\/(drive|upload)/.test(url.host + url.pathname) && !/fonts\.googleapis/.test(url.host)) return;
  // Font Google: cache saat pertama dipakai
  if (/fonts\.(googleapis|gstatic)\.com/.test(url.host)) {
    e.respondWith(caches.open(VERSION + '-font').then(c => c.match(e.request).then(r => r || fetch(e.request).then(res => { c.put(e.request, res.clone()); return res; }).catch(() => r))));
    return;
  }
  if (url.origin !== location.origin) return;
  // Halaman: network-first (agar update cepat), fallback cache
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request).then(res => { caches.open(VERSION).then(c => c.put('index.html', res.clone())); return res; }).catch(() => caches.match('index.html')));
    return;
  }
  // Aset: cache-first
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => {
    if (res.ok) caches.open(VERSION).then(c => c.put(e.request, res.clone()));
    return res;
  })));
});
