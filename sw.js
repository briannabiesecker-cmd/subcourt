// Minimal service worker — makes Rally installable as a PWA.
// Network-first: no caching of app content, so HTML updates are always live.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
