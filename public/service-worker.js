/* eslint-disable no-restricted-globals */

const CACHE_NAME = 'evidence-vault-v1';
const urlsToCache = ['/', '/index.html'];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request).catch(() =>
      caches.match(event.request)
    )
  );
});

self.addEventListener('push', event => {
  const data = event.data?.json() || {};
  const title = data.title || 'BidElevate Evidence Vault';
  const options = {
    body: data.body || "Have you added any evidence this week? It only takes a minute.",
    icon: '/icon-192.png',
    badge: '/icon-192.png',
    tag: 'weekly-reminder',
    renotify: true,
    data: { url: self.location.origin }
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window' }).then(clientList => {
      if (clientList.length > 0) return clientList[0].focus();
      return clients.openWindow(event.notification.data?.url || '/');
    })
  );
});
