/* سرویس‌ورکر: شبکه‌ی اول با مهلت ۴ ثانیه‌ای (اینترنت ضعیف صفحه را معطل نکند)، بعد نسخه‌ی ذخیره‌شده */
const CACHE = 'salas-v1.52';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png', './icon-maskable-512.png'];
const NET_TIMEOUT = 4000;
// هر فایل جداگانه ذخیره می‌شود تا نبودِ یک فایل (مثلاً آیکون) نصب کل سرویس‌ورکر را خراب نکند
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {})))).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });

// صفحه‌ی اصلی: اگر شبکه تا ۴ ثانیه جواب نداد، نسخه‌ی ذخیره‌شده نشان داده می‌شود و دانلود در پس‌زمینه ادامه می‌یابد
async function pageFirst(e, r) {
  const cache = await caches.open(CACHE);
  const cached = (await cache.match(r, { ignoreSearch: true })) || (await cache.match('./index.html'));
  const net = fetch(r).then(res => { if (res && res.ok) cache.put(r, res.clone()); return res; });
  if (!cached) return net;                                   // بار اول: چاره‌ای جز منتظر ماندن نیست
  try {
    return await Promise.race([net, new Promise((_, rej) => setTimeout(rej, NET_TIMEOUT))]);
  } catch (err) {
    e.waitUntil(net.catch(() => {}));                        // به‌روزرسانی کش در پس‌زمینه تمام شود
    return cached;
  }
}
// آیکون و manifest: فوراً از کش، و در پس‌زمینه تازه می‌شود
async function assetSWR(e, r) {
  const cache = await caches.open(CACHE), cached = await cache.match(r);
  const net = fetch(r).then(res => { if (res && res.ok) cache.put(r, res.clone()); return res; }).catch(() => null);
  if (cached) { e.waitUntil(net); return cached; }
  return (await net) || Response.error();
}
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin) return;          // درخواست‌های Google Apps Script هیچ‌وقت ذخیره نمی‌شوند
  const isPage = r.mode === 'navigate' || /\/(index\.html)?$/.test(u.pathname);
  e.respondWith(isPage ? pageFirst(e, r) : assetSWR(e, r));
});

/* لمس اعلان یادآور: برنامه را باز یا فعال کن */
self.addEventListener('notificationclick', e => { e.notification.close(); e.waitUntil(self.clients.matchAll({ type:'window', includeUncontrolled:true }).then(l => { for (const c of l) if ('focus' in c) return c.focus(); return self.clients.openWindow('./'); })); });
