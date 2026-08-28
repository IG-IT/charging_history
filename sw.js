const CACHE='g9-charge-v9';
// No personal data file here on purpose — the repo ships code only, data is imported
// client-side (see app.js SEED_FILE comment). Precaching a data/ path that doesn't
// exist on the hosted site would fail cache.addAll() and break the whole install.
const ASSETS=['./','./index.html','./styles.css','./app.js','./vendor/xlsx.full.min.js','./manifest.webmanifest','./icon.svg','./apple-touch-icon.png'];
self.addEventListener('install',e=>e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)).then(()=>self.skipWaiting())));
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',e=>{ if(e.request.method!=='GET')return; e.respondWith(caches.match(e.request).then(cached=>cached||fetch(e.request).then(r=>{ if(new URL(e.request.url).origin===location.origin){const copy=r.clone(); caches.open(CACHE).then(c=>c.put(e.request,copy));} return r; }).catch(()=>cached))); });
