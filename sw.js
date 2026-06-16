// ── Cuero Viejo Service Worker ───────────────────────────────
const CACHE_NAME  = 'cuero-viejo-v11';
const DATA_CACHE  = 'cuero-viejo-data-v11';

// Recursos de la shell de la app (se cachean en la instalación)
const SHELL_URLS = [
  '/',
  '/index.html',
  '/manifest.json',
  'https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,600;1,300;1,400&family=DM+Mono:wght@300;400;500&family=DM+Sans:wght@300;400;500&display=swap',
];

// ── install: pre-cachear la shell ─────────────────────────────
self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(SHELL_URLS).catch(() => {}))
  );
});

// ── activate: limpiar caches viejos ──────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => k !== CACHE_NAME && k !== DATA_CACHE)
          .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// ── fetch: estrategia por tipo de request ────────────────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Ignorar esquemas no cacheables (chrome-extension, data:, etc.)
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return;

  // Requests a Apps Script (API): network-first, fallback a cache
  if (url.hostname === 'script.google.com') {
    event.respondWith(networkFirstData(event.request));
    return;
  }

  // Fonts de Google: cache-first
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(cacheFirst(event.request));
    return;
  }

  // index.html: siempre network-first para recibir actualizaciones
  if (url.pathname.endsWith('/') || url.pathname.endsWith('index.html')) {
    event.respondWith(networkFirstShell(event.request));
    return;
  }

  // Resto de la shell: cache-first
  event.respondWith(cacheFirst(event.request));
});

// ── Estrategia: network-first para datos ─────────────────────
async function networkFirstData(request) {
  const cache = await caches.open(DATA_CACHE);
  try {
    const response = await fetch(request);
    if (response.ok) {
      // Solo cachear GETs de tipo getXxx (lecturas)
      const url = new URL(request.url);
      const action = url.searchParams.get('action') || '';
      if (request.method === 'GET' && action.startsWith('get')) {
        cache.put(request, response.clone());
      }
    }
    return response;
  } catch {
    const cached = await cache.match(request);
    if (cached) return cached;
    // Sin red y sin caché: devolver respuesta de error legible
    return new Response(
      JSON.stringify({ ok: false, error: 'Sin conexión. Mostrando datos guardados.' }),
      { headers: { 'Content-Type': 'application/json' } }
    );
  }
}

// ── Estrategia: network-first para HTML (siempre actualizado) ─
async function networkFirstShell(request) {
  const cache = await caches.open(CACHE_NAME);
  try {
    const response = await fetch(request);
    if (response.ok) cache.put(request, response.clone());
    return response;
  } catch {
    const cached = await cache.match(request);
    return cached || new Response('Offline', { status: 503 });
  }
}

// ── Estrategia: cache-first para shell y fuentes ─────────────
async function cacheFirst(request) {
  const cached = await caches.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    return new Response('Offline', { status: 503 });
  }
}
