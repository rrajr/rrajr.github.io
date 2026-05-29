const cacheName = 'cache-rrajr-v1';
const cacheResources = [
	'index.html', 
];
async function precache(){
	const cache = await caches.open(cacheName);
	return cache.addAll(cacheResources);
}
self.addEventListener("install", event => {
	event.waitUntil(precache());
});
async function networkFirst(request) {
	try {
		const networkResponse = await fetch(request);
		if (networkResponse.ok) {
			const cache = await caches.open(cacheName);
			cache.put(request, networkResponse.clone());
		}
		return networkResponse;
	} catch (error) {
		const cachedResponse = await caches.match(request);
		return cachedResponse || Response.error();
	}
}
async function cacheFirst(request) {
	const cachedResponse = await caches.match(request);
	if (cachedResponse) {
		return cachedResponse;
	}
	try {
		const networkResponse = await fetch(request);
		if (networkResponse.ok) {
			const cache = await caches.open("cache-rrajr-v1");
			cache.put(request, networkResponse.clone());
		}
		return networkResponse;
	} catch (error) {
		return Response.error();
	}
}
self.addEventListener("fetch", (event) => {
	const url = new URL(event.request.url);
	if (url.pathname.match(/^\/webfonts/)) {
		event.respondWith(cacheFirst(event.request));
	} else {
		event.respondWith(networkFirst(event.request));
	}
});