/* Service worker: deixa o app abrir sem internet e sempre busca a versão nova quando há internet */
const CACHE = "atividades-v3";
const ARQUIVOS = ["./", "index.html", "app.js", "config.js", "manifest.json", "icon-192.png", "icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE)
    .then(c => c.addAll(ARQUIVOS.map(a => new Request(a, { cache: "reload" }))))
    .then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET") return;
  const u = new URL(r.url);
  const proprio = u.origin === location.origin && !u.pathname.startsWith("/__/");
  const biblioteca = u.hostname === "www.gstatic.com" || u.hostname === "cdnjs.cloudflare.com";
  if (proprio) e.respondWith(redePrimeiro(r));
  else if (biblioteca) e.respondWith(cachePrimeiro(r));
  // O banco de dados vai direto para a internet (o próprio Firebase guarda uma cópia para usar sem internet)
});

// Arquivos do app: sempre confere no servidor se há versão nova; sem internet, usa a guardada
async function redePrimeiro(r) {
  try {
    const resp = await fetch(r, { cache: "no-cache" });
    if (resp.ok) (await caches.open(CACHE)).put(r, resp.clone());
    return resp;
  } catch (e) {
    return (await caches.match(r)) || (r.mode === "navigate" ? caches.match("index.html") : Response.error());
  }
}

// Bibliotecas com versão fixa: usa a guardada; se não tiver, baixa e guarda
async function cachePrimeiro(r) {
  const c = await caches.match(r);
  if (c) return c;
  const resp = await fetch(r);
  if (resp.ok) (await caches.open(CACHE)).put(r, resp.clone());
  return resp;
}
