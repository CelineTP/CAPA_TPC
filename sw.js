// Service worker Capa Stock
// - Pages et scripts de l'appli : le réseau d'abord (les mises à jour arrivent tout de suite),
//   avec repli sur le cache hors-ligne ou si le réseau met plus de 4 s à répondre.
// - Bibliothèques externes (XLSX, ZXing) : cache d'abord, elles ne changent jamais.
// - Installation tolérante : un fichier manquant (ex. une icône) n'empêche plus l'installation.
const CACHE_NAME = "capa-stock-v6";
const CORE_ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png"
];
const CDN_LIBS = [
  "https://cdn.sheetjs.com/",
  "https://unpkg.com/@zxing/library"
];
const NETWORK_TIMEOUT_MS = 4000;

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      Promise.allSettled(CORE_ASSETS.map((url) => cache.add(new Request(url, { cache: "reload" }))))
    )
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

function putInCache(req, res) {
  if (res && (res.ok || res.type === "opaque")) {
    const clone = res.clone();
    caches.open(CACHE_NAME).then((cache) => cache.put(req, clone)).catch(() => {});
  }
}

function fromCache(req) {
  return caches.match(req).then((cached) => {
    if (cached) return cached;
    if (req.mode === "navigate") return caches.match("./index.html");
    return null;
  });
}

function networkFirst(req) {
  return new Promise((resolve) => {
    let settled = false;
    const done = (res) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(res);
    };
    const timer = setTimeout(() => {
      fromCache(req).then((cached) => {
        if (cached) done(cached);
      });
    }, NETWORK_TIMEOUT_MS);

    fetch(req, { cache: "no-cache" })
      .then((res) => {
        putInCache(req, res);
        done(res);
      })
      .catch(() => {
        fromCache(req).then((cached) => done(cached || Response.error()));
      });
  });
}

function cacheFirst(req) {
  return caches.match(req).then((cached) => {
    if (cached) return cached;
    return fetch(req).then((res) => {
      putInCache(req, res);
      return res;
    });
  });
}

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const sameOrigin = new URL(req.url).origin === self.location.origin;
  const isLib = CDN_LIBS.some((prefix) => req.url.startsWith(prefix));

  if (isLib) {
    event.respondWith(cacheFirst(req));
  } else if (sameOrigin) {
    event.respondWith(networkFirst(req));
  }
  // Autres domaines (Google, polices…) : on laisse le navigateur faire.
});
