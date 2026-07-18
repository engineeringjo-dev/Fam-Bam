/* FAM BAM in TRABZAM — Service Worker (offline + map tile caching) */
const APP_CACHE = 'fambam-app-v45';
const TILE_CACHE = 'fambam-tiles-v4';
const LIB_CACHE = 'fambam-libs-v4';
const MAX_TILES = 1500; // cap cached tiles so storage doesn't grow forever

const APP_SHELL = ['./', './index.html', './anthem.mp3'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(APP_CACHE).then(c => c.addAll(APP_SHELL)).catch(()=>{}));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => ![APP_CACHE, TILE_CACHE, LIB_CACHE].includes(k)).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

function isTile(url){
  return /tile|basemaps|\bmt[0-9]\.google|khms|cartocdn|cartodb|openstreetmap|arcgisonline|\/vt\//i.test(url);
}
function isLib(url){
  return /cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|fonts\.googleapis\.com|fonts\.gstatic\.com|unpkg\.com/i.test(url);
}

async function trimCache(name, max){
  const c = await caches.open(name);
  const keys = await c.keys();
  if (keys.length > max){ for (let i=0; i<keys.length-max; i++) await c.delete(keys[i]); }
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = req.url;

  // App navigations: network-first, fall back to cached shell (works offline)
  if (req.mode === 'navigate'){
    e.respondWith((async () => {
      try { const net = await fetch(req); const c = await caches.open(APP_CACHE); c.put('./index.html', net.clone()); return net; }
      catch(_){ return (await caches.match('./index.html')) || (await caches.match('./')) || Response.error(); }
    })());
    return;
  }

  // Map tiles: cache-first, then network (caches whatever areas you pan over)
  if (isTile(url)){
    e.respondWith((async () => {
      const hit = await caches.match(req);
      if (hit) return hit;
      try { const net = await fetch(req, {mode:'no-cors'}); const c = await caches.open(TILE_CACHE); c.put(req, net.clone()); trimCache(TILE_CACHE, MAX_TILES); return net; }
      catch(_){ return hit || Response.error(); }
    })());
    return;
  }

  // Libraries/fonts (Leaflet, Supabase, Tajawal): cache-first
  if (isLib(url)){
    e.respondWith((async () => {
      const hit = await caches.match(req);
      if (hit) return hit;
      try { const net = await fetch(req); const c = await caches.open(LIB_CACHE); c.put(req, net.clone()); return net; }
      catch(_){ return hit || Response.error(); }
    })());
    return;
  }
});

// focus/open the app when a notification is tapped
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil((async () => {
    const all = await self.clients.matchAll({type:'window', includeUncontrolled:true});
    if (all.length){ all[0].focus(); }
    else { self.clients.openWindow('./'); }
  })());
});
