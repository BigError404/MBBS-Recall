const CACHE='mbbs-recall-pro-v7.4-header-and-bridge-fix';
const CORE=['./','./index.html','./manifest.json','./icon.svg','./engine.js','./app.js','./workbook-seed.js','./xlsx-import.js'];
const NETWORK_TIMEOUT_MS=4000;
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(CORE)).then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith('mbbs-recall-pro-')&&k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
// Network-first: a fix pushed to GitHub Pages reaches the phone on the next online open.
// The cache is only the offline fallback (and is used if the network takes longer than 4 s).
self.addEventListener('fetch',e=>{
  const r=e.request;
  if(r.method!=='GET'||new URL(r.url).origin!==self.location.origin)return;
  const network=fetch(r,{cache:'no-cache'}).then(res=>{
    if(res&&res.ok){const copy=res.clone();caches.open(CACHE).then(c=>c.put(r,copy));}
    return res;
  }).catch(()=>null);
  const timeout=new Promise(resolve=>setTimeout(()=>resolve(null),NETWORK_TIMEOUT_MS));
  e.respondWith(Promise.race([network,timeout]).then(res=>res||caches.match(r).then(hit=>hit||network).then(hit=>hit||(r.mode==='navigate'?caches.match('./index.html'):Response.error()))));
});
