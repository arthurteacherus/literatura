// Сервис-воркер: после первой загрузки программа открывается и работает без интернета.
// Стратегия: сначала из кэша, а свежую версию тихо подтягиваем в фоне (обновление — со 2-го запуска).
// Версию сборки и список файлов подставляет build.js.
const CACHE = "lit-4307ff2698";
const ASSETS = ["./","app.js","icon-192.png","icon-512.png","index.html","manifest.webmanifest","platform.js","tablet-ui.js","vendor/fonts/fonts.css","vendor/fonts/jizaRExUiTo99u79D0-ExdGM.woff2","vendor/fonts/jizaRExUiTo99u79D0aExdGM.woff2","vendor/fonts/jizaRExUiTo99u79D0KExQ.woff2","vendor/fonts/jizaRExUiTo99u79D0yExdGM.woff2","vendor/fonts/jizfRExUiTo99u79B_mh0O6tLQ.woff2","vendor/fonts/jizfRExUiTo99u79B_mh0OCtLQ0Z.woff2","vendor/fonts/jizfRExUiTo99u79B_mh0OOtLQ0Z.woff2","vendor/fonts/jizfRExUiTo99u79B_mh0OqtLQ0Z.woff2","vendor/fonts/nuFiD-vYSZviVYUb_rj3ij__anPXDTjYgFE_.woff2","vendor/fonts/nuFiD-vYSZviVYUb_rj3ij__anPXDTLYgFE_.woff2","vendor/fonts/nuFiD-vYSZviVYUb_rj3ij__anPXDTPYgFE_.woff2","vendor/fonts/nuFiD-vYSZviVYUb_rj3ij__anPXDTzYgA.woff2","vendor/logo.png","vendor/pdf.min.js","vendor/pdf.worker.min.js","vendor/react-dom.production.min.js","vendor/react.production.min.js"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS.map((u) => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith("lit-") && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // Яндекс.Диск и прочее — напрямую
  event.respondWith(
    caches.open(CACHE).then(async (cache) => {
      const cached = (await cache.match(req, { ignoreSearch: true })) || (req.mode === "navigate" ? await cache.match("index.html") : null);
      const network = fetch(req)
        .then((res) => {
          if (res && res.ok) cache.put(req, res.clone());
          return res;
        })
        .catch(() => null);
      return cached || (await network) || new Response("Нет сети, а страницы нет в кэше", { status: 503 });
    })
  );
});
