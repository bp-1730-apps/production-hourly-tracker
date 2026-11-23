// Hourly Tracker offline cache: network first (always fresh when online), falls back to the last copy when offline.
var C="ht-v1";
self.addEventListener("install",function(e){self.skipWaiting();e.waitUntil(caches.open(C).then(function(c){return c.addAll(["./","index.html","xlsx.min.js","manifest.webmanifest","icon-192.png","icon-512.png"]).catch(function(){})}))});
self.addEventListener("activate",function(e){e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==C}).map(function(k){return caches.delete(k)}))}).then(function(){return self.clients.claim()}))});
self.addEventListener("fetch",function(e){
  var r=e.request,u=new URL(r.url);
  if(r.method!=="GET"||u.origin!==location.origin)return;   // Supabase / L2L calls go straight to the network
  e.respondWith(new Promise(function(res){
    var done=false,timer=setTimeout(function(){caches.match(r,{ignoreSearch:true}).then(function(m){if(m&&!done){done=true;res(m)}})},4000);
    fetch(r).then(function(n){clearTimeout(timer);if(n&&n.ok){var cp=n.clone();caches.open(C).then(function(c){c.put(new Request(u.origin+u.pathname),cp)})}if(!done){done=true;res(n)}})
    .catch(function(){clearTimeout(timer);caches.match(r,{ignoreSearch:true}).then(function(m){if(!done){done=true;res(m||new Response("Offline",{status:503}))}})});
  }));
});
