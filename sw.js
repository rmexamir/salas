/* سرویس‌ورکر: ابتدا شبکه (تا نسخه‌ی جدید همیشه برسد)، اگر آفلاین بود نسخه‌ی ذخیره‌شده */
const CACHE = 'salas-v1.09';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png', './icon-maskable-512.png'];
// هر فایل جداگانه ذخیره می‌شود تا نبودِ یک فایل (مثلاً آیکون) نصب کل سرویس‌ورکر را خراب نکند
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {})))).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin) return;          // درخواست‌های Google Apps Script هیچ‌وقت ذخیره نمی‌شوند
  e.respondWith(fetch(r).then(res => { if (res.ok) { const cp = res.clone(); caches.open(CACHE).then(c => c.put(r, cp)); } return res; })
    .catch(() => caches.match(r).then(m => m || caches.match('./index.html'))));
});

/* لمس اعلان یادآور: برنامه را باز یا فعال کن */
self.addEventListener('notificationclick', e => { e.notification.close(); e.waitUntil(self.clients.matchAll({ type:'window', includeUncontrolled:true }).then(l => { for (const c of l) if ('focus' in c) return c.focus(); return self.clients.openWindow('./'); })); });
