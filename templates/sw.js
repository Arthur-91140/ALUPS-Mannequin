/* Service worker ALUPS — Suivi des mannequins
   Servi par Flask afin que les chemins respectent le sous-dossier de
   déploiement. Bump CACHE pour forcer la mise à jour du cache. */

const CACHE = 'alups-v5';
const OFFLINE_URL = '{{ url_for("offline") }}';
const STATIC_PREFIX = '{{ url_for("static", filename="") }}';

const PRECACHE = [
  OFFLINE_URL,
  '{{ url_for("manifest") }}',
  '{{ url_for("static", filename="css/style.css") }}',
  '{{ url_for("static", filename="img/logo.png") }}',
  '{{ url_for("static", filename="icons/icon-192.png") }}',
  '{{ url_for("static", filename="icons/icon-512.png") }}',
  '{{ url_for("static", filename="icons/apple-touch-icon.png") }}'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE)
      .then((cache) => cache.addAll(PRECACHE))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Navigations : réseau d'abord, repli sur le cache puis la page hors ligne.
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match(req).then((r) => r || caches.match(OFFLINE_URL)))
    );
    return;
  }

  // Assets statiques (css, icônes, images) : stale-while-revalidate.
  if (url.origin === self.location.origin) {
    if (!url.pathname.startsWith(STATIC_PREFIX)) return; // BD, photos, exports : pas de cache
    event.respondWith(
      caches.match(req).then((cached) => {
        const network = fetch(req)
          .then((res) => {
            if (res && res.status === 200) {
              const copy = res.clone();
              caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
            }
            return res;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
    return;
  }

  // Cross-origin (polices, CDN icônes, signature_pad) : cache d'abord.
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req)
        .then((res) => {
          if (res && (res.status === 200 || res.type === 'opaque')) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => cached);
    })
  );
});
