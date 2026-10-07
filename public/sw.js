// Feed Mail service worker: lets phones install the app. It deliberately
// caches nothing, so no email ever gets stored on the device.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {}); // network only
