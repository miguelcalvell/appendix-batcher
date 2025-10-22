// bump cache to v6.9 to force refresh
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open('appendix-batcher-v6.9').then(cache => cache.addAll([
    './', './index.html', './style.css', './app.js', './manifest.json',
    'https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js',
    'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js'
  ])));
});
self.addEventListener('fetch', (e) => {
  e.respondWith(caches.match(e.request).then(resp => resp || fetch(e.request)));
});
