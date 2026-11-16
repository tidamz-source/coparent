/* Service worker : l'app marche hors ligne et se met à jour toute seule */
const CACHE='coparent-v1';
const SHELL=['./','index.html','manifest.webmanifest','icons/icon-192.png','icons/icon-512.png','icons/apple-touch-icon.png',
  'avatars/papa.webp','avatars/maman.webp','avatars/enfant1.webp','avatars/enfant2.webp'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const r=e.request;if(r.method!=='GET')return;
  const u=new URL(r.url);
  if(u.hostname==='data.education.gouv.fr')return; // API vacances : toujours en direct
  if(r.mode==='navigate'||u.pathname.endsWith('/')||u.pathname.endsWith('index.html')){
    // réseau d'abord pour avoir la dernière version, cache si hors ligne
    e.respondWith(fetch(r).then(res=>{const c=res.clone();caches.open(CACHE).then(x=>x.put('./',c));return res}).catch(()=>caches.match('./')));
    return;
  }
  // le reste (images, polices) : cache d'abord, mis à jour en arrière-plan
  e.respondWith(caches.match(r).then(hit=>{
    const net=fetch(r).then(res=>{if(res.ok||res.type==='opaque'){const c=res.clone();caches.open(CACHE).then(x=>x.put(r,c))}return res}).catch(()=>hit);
    return hit||net;
  }));
});
