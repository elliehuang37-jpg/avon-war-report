const CACHE_NAME = "avon-fragrance-consultant-v3";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./assets/all-p01-01.png",
  "./assets/all-p02-04.png",
  "./assets/all-p02-06.png",
  "./assets/all-p02-09.png",
  "./assets/all-p03-11.png",
  "./assets/all-p03-12.png",
  "./assets/all-p04-17.png",
  "./assets/all-p04-18.png",
  "./assets/all-p05-22.png",
  "./assets/all-p05-23.png",
  "./assets/all-p05-24.png",
  "./assets/all-p06-27.png",
  "./assets/all-p06-28.png",
  "./assets/all-p06-29.png",
  "./assets/all-p07-34.png",
  "./assets/all-p07-35.png",
  "./assets/all-p08-39.png",
  "./assets/all-p08-40.png",
  "./assets/all-p08-41.png",
  "./assets/all-p09-43.png",
  "./assets/all-p09-45.png",
  "./assets/all-p09-47.png",
  "./assets/all-p10-50.png",
  "./assets/all-p10-51.png",
  "./assets/all-p10-53.png",
  "./assets/all-p10-54.png",
  "./assets/all-p11-56.png",
  "./assets/all-p11-57.png",
  "./assets/all-p11-58.png",
  "./assets/all-p13-61.png",
  "./assets/all-p13-64.png",
  "./assets/all-p14-67.png",
  "./assets/all-p14-69.png",
  "./assets/all-p15-72.png"
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});
