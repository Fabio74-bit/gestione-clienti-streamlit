self.addEventListener('install', event => {
  console.log('Service Worker installato');
  event.waitUntil(
    caches.open('gestione-clienti-cache').then(cache => {
      return cache.addAll(['/', '/static/manifest.json']);
    })
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request).then(response => {
      return response || fetch(event.request);
    })
  );
});

