const CACHE = 'fighter-hq-v1';
const ASSETS = ['./', './index.html', './site-config.json', './manifest.webmanifest'];

self.addEventListener('install', function(e) {
  e.waitUntil(caches.open(CACHE).then(function(c) { return c.addAll(ASSETS); }));
});

self.addEventListener('fetch', function(e) {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request).catch(function() { return caches.match('./index.html'); });
    })
  );
});
