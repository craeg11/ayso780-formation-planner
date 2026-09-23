const CACHE_NAME = 'lineup-planner-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json'
];

self.addEventListener('fetch', (event) => {
  if (event.request.url.includes('script.google.com')) {
    return; // Let the browser handle the request natively
  }

  event.respondWith(
    caches.match(event.request).then((response) => {
      return response || fetch(event.request);
    })
  );
});
