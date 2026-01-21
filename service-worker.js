const CACHE_NAME = "qtp-cache-v2";
const PRECACHE_URLS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./service-worker.js",
  "./vendor/mammoth.browser.min.js",
  "./vendor/expr-eval.bundle.min.js",
  "./vendor/jspdf.umd.min.js",
  "./vendor/html2canvas.min.js",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    (async () => {
      const cache = await caches.open(CACHE_NAME);
      await cache.addAll(PRECACHE_URLS.map((u) => new Request(u, { cache: "reload" })));
      await self.skipWaiting();
    })(),
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)));
      await self.clients.claim();
    })(),
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  // Only handle same-origin requests.
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    (async () => {
      const cache = await caches.open(CACHE_NAME);
      const cached = await cache.match(req, { ignoreSearch: false });
      if (cached) return cached;

      try {
        const fresh = await fetch(req);
        // Cache successful basic responses.
        if (fresh && fresh.status === 200 && (fresh.type === "basic" || fresh.type === "default")) {
          cache.put(req, fresh.clone());
        }
        return fresh;
      } catch (e) {
        // Fallback to app shell if available.
        const shell = await cache.match("./index.html");
        if (shell) return shell;
        throw e;
      }
    })(),
  );
});
