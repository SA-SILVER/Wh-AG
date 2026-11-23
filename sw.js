// "איפה הכסף" – service worker.
// 1. Keeps a full copy of the app on the phone, so it opens without internet.
// 2. Receives chats shared from WhatsApp. The shared file is handled here, inside the phone,
//    and handed to the app; it is never sent anywhere.
// 3. A new version is downloaded in the background but only takes over when the user taps "update".
const VERSION = 'kesef-1.13';
const FILES = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES.map(f => new Request(f, {cache: 'reload'})))));
});

self.addEventListener('message', e => { if (e.data === 'skip-waiting') self.skipWaiting(); });

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== VERSION && k !== 'shared') await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return;
  if (e.request.method === 'POST' && url.pathname.endsWith('/share-target')) {
    e.respondWith((async () => {
      const fd = await e.request.formData();
      const files = fd.getAll('chat').filter(f => f && typeof f !== 'string' && f.size);
      const c = await caches.open('shared');
      await c.put('./shared-info', new Response(JSON.stringify({at: Date.now(), files: files.map(f => ({name: f.name, type: f.type, size: f.size}))}), {headers: {'content-type': 'application/json'}}));
      if (files[0]) await c.put('./shared-file', new Response(files[0]));
      else await c.delete('./shared-file');
      return Response.redirect(new URL('./?shared=1', self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== 'GET') return;
  if (url.pathname.endsWith('/version.json')) return;            // always from the network (update notes)
  e.respondWith(caches.open(VERSION).then(c => c.match(e.request, {ignoreSearch: true})).then(r => r || fetch(e.request)));
});
