// Test page service worker: keeps a copy of the page on the phone (works offline)
// and receives files shared from WhatsApp. The shared file is handled here, on the phone; it is never sent anywhere.
const VERSION = 'share-test-1';
const FILES = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== VERSION && k !== 'shared') await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method === 'POST' && url.pathname.endsWith('/share-target')) {
    e.respondWith((async () => {
      const fd = await e.request.formData();
      const files = fd.getAll('chat').filter(f => f && typeof f !== 'string');
      const info = {
        at: Date.now(),
        fields: [...fd.keys()],
        title: String(fd.get('title') || ''),
        files: files.map(f => ({name: f.name, type: f.type, size: f.size}))
      };
      const c = await caches.open('shared');
      await c.put('./shared-info', new Response(JSON.stringify(info), {headers: {'content-type': 'application/json'}}));
      if (files[0]) await c.put('./shared-file', new Response(files[0]));
      return Response.redirect(new URL('./?shared=1', self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  e.respondWith(caches.match(e.request, {ignoreSearch: true}).then(r => r || fetch(e.request)));
});
