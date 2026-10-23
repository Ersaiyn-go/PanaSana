const C="uchet-v5";
const F=["./","index.html","manifest.webmanifest","icon-180.png","icon-192.png","icon-512.png"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(F)).then(()=>self.skipWaiting()))});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C&&x!=="share").map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener("fetch",e=>{
  const u=new URL(e.request.url);
  if(e.request.method==="POST"&&u.pathname.endsWith("/share")){
    e.respondWith((async()=>{
      try{
        const f=await e.request.formData(),c=await caches.open("share");
        await c.delete("__sf");await c.delete("__st");
        const file=f.get("file");
        if(file&&file.size)await c.put("__sf",new Response(file,{headers:{"Content-Type":file.type||"application/octet-stream"}}));
        const t=[f.get("title"),f.get("text")].filter(Boolean).join("\n");
        if(t)await c.put("__st",new Response(t));
      }catch(x){}
      return Response.redirect(new URL("./?share=1",self.registration.scope).href,303);
    })());
    return;
  }
  if(e.request.method!=="GET")return;
  if(u.origin==="https://cdnjs.cloudflare.com"){
    e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(n=>{const c=n.clone();caches.open(C).then(x=>x.put(e.request,c));return n})));
    return;
  }
  if(u.origin!==location.origin)return;
  e.respondWith(
    fetch(e.request).then(r=>{const c=r.clone();caches.open(C).then(x=>x.put(e.request,c));return r})
    .catch(()=>caches.match(e.request,{ignoreSearch:true}).then(r=>r||caches.match("index.html")))
  );
});
