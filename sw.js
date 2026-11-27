// Service worker ของ F Office
// - หน้าเว็บ (HTML): โหลดจากอินเทอร์เน็ตก่อน เพื่อให้ได้เวอร์ชันล่าสุดทุกครั้งที่อัปเดตบน GitHub
//   ถ้าออฟไลน์จะใช้ฉบับที่เก็บไว้
// - ไอคอน/manifest: ใช้ที่เก็บไว้ก่อน แล้วอัปเดตเบื้องหลัง
// - ข้อมูลระบบจาก Google Sheet (คนละโดเมน) ไม่ผ่านตัวนี้ จึงสดใหม่เสมอ
const CACHE = 'foffice-v1';
const SHELL = ['./', 'index.html', 'manifest.json', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;   // ปล่อยคำขอข้ามโดเมน (API, ฟอนต์) ผ่านตามปกติ

  if (req.mode === 'navigate' || req.destination === 'document') {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put('index.html', copy));
          return res;
        })
        .catch(() => caches.match('index.html'))
    );
    return;
  }

  e.respondWith(
    caches.match(req).then((hit) => {
      const net = fetch(req).then((res) => {
        if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
        return res;
      }).catch(() => hit);
      return hit || net;
    })
  );
});
