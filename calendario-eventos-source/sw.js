const CACHE = "ic-eventos-v2";
const ASSETS = ["/", "/index.html", "/manifest.json", "/logo.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  if (e.request.url.includes("firebaseio.com") || e.request.url.includes("googleapis.com")) return;

  // Navegação/HTML: rede primeiro, sempre buscando a versão publicada mais recente no Netlify.
  // Era aqui que o cache-first antigo prendia todo mundo numa versão desatualizada do app.
  if (e.request.mode === "navigate" || e.request.destination === "document") {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy));
          return res;
        })
        .catch(() => caches.match(e.request).then(c => c || caches.match("/index.html")))
    );
    return;
  }

  // Demais arquivos estáticos (logo, manifest) — cache primeiro, com respaldo offline
  e.respondWith(caches.match(e.request).then(c => c || fetch(e.request).catch(() => caches.match("/index.html"))));
});

self.addEventListener("notificationclick", e => {
  e.notification.close();
  e.waitUntil(clients.matchAll({ type: "window" }).then(cs => {
    if (cs.length) return cs[0].focus();
    return clients.openWindow("/");
  }));
});

self.addEventListener("push", e => {
  const data = e.data ? e.data.json() : {};
  e.waitUntil(self.registration.showNotification(data.title || "Calendário de Eventos", {
    body: data.body || "",
    icon: "/logo.png",
    badge: "/logo.png"
  }));
});
