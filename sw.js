/* Service worker: keeps the app itself (not the site data) available offline.
   The site data is saved separately inside the app (IndexedDB).
   When you upload a new version of any app file to GitHub, change VERSION below
   so phones pick up the new files. */
const VERSION = 'erps-app-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './jspdf.umd.min.js'];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(VERSION).then(function (cache) { return cache.addAll(SHELL); }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k.indexOf('erps-app-') === 0 && k !== VERSION; })
                             .map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);

  // Never touch calls to Google's web app: the app handles its own data
  if (/(^|\.)google\.com$/.test(url.hostname) && url.hostname.indexOf('fonts.') !== 0) return;
  if (/googleusercontent\.com$/.test(url.hostname)) return;

  // The app's own files: show the saved copy at once, refresh it in the background
  if (url.origin === self.location.origin) {
    event.respondWith(caches.open(VERSION).then(function (cache) {
      return cache.match(req, { ignoreSearch: true }).then(function (cached) {
        var network = fetch(req).then(function (res) {
          if (res && res.ok) cache.put(req, res.clone());
          return res;
        }).catch(function () { return null; });
        if (cached) { event.waitUntil(network); return cached; }
        return network.then(function (res) {
          if (res) return res;
          if (req.mode === 'navigate') return cache.match('./index.html').then(function (i) { return i || new Response('Offline', { status: 503 }); });
          return new Response('Offline', { status: 503 });
        });
      });
    }));
    return;
  }

  // Fonts: keep a copy after the first load
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(caches.open(VERSION).then(function (cache) {
      return cache.match(req).then(function (cached) {
        if (cached) return cached;
        return fetch(req).then(function (res) {
          if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
          return res;
        }).catch(function () { return Response.error(); });
      });
    }));
  }
});
