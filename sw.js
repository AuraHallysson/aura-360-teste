/* Service Worker — Aura Compras (PWA)
   Estratégia: cache-first para o "app shell" (funciona offline após o 1º acesso),
   com atualização em segundo plano. NÃO armazena dados de solicitações — estes
   ficam no localStorage / SharePoint da aplicação. */
const CACHE = "aura-compras-v1";
const CORE = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => Promise.allSettled(CORE.map((u) => c.add(u))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return; // nunca intercepta POST/PATCH (ex.: Microsoft Graph)
  const url = new URL(req.url);
  // Não intercepta chamadas de API (Graph/login) — sempre rede
  if (/graph\.microsoft\.com|login\.microsoftonline\.com|msauth\.net/.test(url.host)) return;
  e.respondWith(
    caches.match(req).then((hit) => {
      const rede = fetch(req).then((res) => {
        if (res && res.status === 200) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => { try { c.put(req, copy); } catch (_) {} });
        }
        return res;
      }).catch(() => hit || caches.match("./index.html"));
      return hit || rede;
    })
  );
});
