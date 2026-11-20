// Service worker – aplikacja działa offline po pierwszym uruchomieniu (także z paska Windows).
const CACHE = 'tm-v54';
const SHELL = ['./', './index.html', './styles.css', './theme.css', './fonts/geist-latin-wght-normal.woff2', './fonts/geist-latin-ext-wght-normal.woff2', './fonts/geist-mono-latin-wght-normal.woff2', './fonts/geist-mono-latin-ext-wght-normal.woff2', './app.js', './config.js', './manifest.webmanifest',
  './icons/tm-a-logo.svg', './icons/tm-a-16.png', './icons/tm-a-32.png', './icons/tm-a-192.png', './icons/tm-a-512.png', './icons/tm-a-maskable-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin || u.pathname.startsWith('/__')) return; // /__ics = import kalendarza, zawsze z sieci // API Google/Graph zawsze z sieci
  // network-first dla plików aplikacji (szybki rozwój), fallback na cache offline
  e.respondWith(fetch(e.request).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request, {ignoreSearch: true})));
});
