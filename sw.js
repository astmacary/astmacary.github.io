// RCI Serve — service worker.
// Registered at the top of the site ("/"), but it only answers for RCI Serve's own files.
// The team apps (/rci-productions/, /rci-security/, /rci-ushers/, /rci-worship-arts/) have their own
// service workers and are never touched here. Always asks for the newest copy first; the saved copy is only used offline.
const CACHE = "rci-serve-v38";
const SHELL = ["/", "/index.html", "/communications.js", "/rosters.js", "/manifest.webmanifest",
  "/icons/icon-192.png", "/icons/icon-512.png", "/icons/icon-maskable-512.png",
  "/icons/apple-touch-icon.png", "/icons/favicon.png", "/icons/wordmark.png", "/icons/wordmark-dark.png"];
const OWN = new Set(SHELL);

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE)
    .then(c => c.addAll(SHELL.map(u => new Request(u, { cache: "reload" }))))
    .then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith("rci-serve-") && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

function newestFirst(request, cacheKey){
  return fetch(request.url, { cache: "no-cache", credentials: "same-origin" })
    .then(res => {
      if (res.ok){ const copy = res.clone(); caches.open(CACHE).then(c => c.put(cacheKey, copy)); }
      return res;
    })
    .catch(() => caches.match(cacheKey).then(hit => hit || (request.mode === "navigate" ? caches.match("/index.html") : Response.error())));
}

self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== self.location.origin) return;
  if (!OWN.has(url.pathname)) return;   // a team app's file: leave it alone
  const key = url.pathname === "/" ? "/index.html" : url.pathname;
  e.respondWith(newestFirst(e.request, new Request(url.origin + key)));
});
