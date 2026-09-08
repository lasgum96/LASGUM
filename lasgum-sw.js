const CACHE='lasgum-v64-nokey-fix3-shell';
const RUNTIME_CACHE='lasgum-v64-nokey-fix3-runtime';

const SHELL=[
  './',
  './index.html',
];

const RUNTIME_URLS=[
  'https://cdn.tailwindcss.com',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
];

self.addEventListener('install', event => {
  event.waitUntil((async()=>{
    const cache=await caches.open(CACHE);
    await cache.addAll(SHELL);

    // CDN files are stored as opaque responses when the browser permits it.
    // If a CDN is temporarily unavailable, installation still succeeds.
    for(const url of RUNTIME_URLS){
      try{
        const res=await fetch(url,{mode:'no-cors',cache:'no-store'});
        if(res && (res.ok || res.type==='opaque')){
          const rc=await caches.open(RUNTIME_CACHE);
          await rc.put(url,res.clone());
        }
      }catch(e){}
    }
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async()=>{
    const keep=new Set([CACHE,RUNTIME_CACHE]);
    const keys=await caches.keys();
    await Promise.all(keys.filter(k=>!keep.has(k)).map(k=>caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const req=event.request;
  if(req.method!=='GET') return;

  const url=new URL(req.url);

  if(url.origin===location.origin){
    event.respondWith((async()=>{
      const cached=await caches.match(req);
      try{
        // Network-first agar perangkat segera menerima versi HTML terbaru.
        const fresh=await fetch(req,{cache:'no-store'});
        if(fresh && fresh.ok){
          const cache=await caches.open(CACHE);
          await cache.put(req,fresh.clone());
          return fresh;
        }
      }catch(e){}
      return cached || Response.error();
    })());
    return;
  }

  if(RUNTIME_URLS.includes(req.url)){
    event.respondWith((async()=>{
      const rc=await caches.open(RUNTIME_CACHE);
      const cached=await rc.match(req.url);
      try{
        const fresh=await fetch(req,{mode:'no-cors'});
        if(fresh && (fresh.ok || fresh.type==='opaque')){
          rc.put(req.url,fresh.clone());
          return fresh;
        }
      }catch(e){}
      return cached || Response.error();
    })());
  }
});
