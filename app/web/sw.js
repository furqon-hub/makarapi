// Service worker minimal (syarat PWA). Tidak menyimpan cache: pemeriksaan selalu butuh server.
self.addEventListener("install", function(){ self.skipWaiting(); });
self.addEventListener("activate", function(e){ e.waitUntil(self.clients.claim()); });
self.addEventListener("fetch", function(e){ e.respondWith(fetch(e.request)); });
