const CACHE='lasgum-v44-shell';
const SHELL=['./','./APLIKASI_UJIAN_V44_FINAL_DISTRIBUTION.html','./lasgum.webmanifest'];
const RUNTIME_ORIGINS=['https://cdn.tailwindcss.com','https://cdn.jsdelivr.net'];
self.addEventListener('install',event=>{
  event.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL).catch(()=>{})).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',event=>{
  event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
self.addEventListener('fetch',event=>{
  const req=event.request, u=new URL(req.url);
  const same=u.origin===location.origin, runtime=RUNTIME_ORIGINS.includes(u.origin);
  if(!same && !runtime)return;
  event.respondWith(caches.match(req).then(cached=>{
    if(cached)return cached;
    return fetch(req).then(r=>{
      if(r && (r.ok || r.type==='opaque')){
        const copy=r.clone(); caches.open(CACHE).then(c=>c.put(req,copy)).catch(()=>{});
      }
      return r;
    }).catch(()=>cached || new Response('',{status:503,statusText:'Offline'}));
  }));
});
