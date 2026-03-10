const CACHE_VERSION = 'v1';
const APP_SHELL_CACHE = `accounting-forecast-app-shell-${CACHE_VERSION}`;
const RUNTIME_CACHE = `accounting-forecast-runtime-${CACHE_VERSION}`;
const ACTIVE_CACHES = [APP_SHELL_CACHE, RUNTIME_CACHE];
const APP_SHELL_ASSETS = [
	'./',
	'./index.html',
	'./styles.css',
	'./app.js',
	'./manifest.webmanifest',
	'./vendor/chart.umd.min.js',
	'./apple-touch-icon.png',
	'./favicon-16x16.png',
	'./favicon-32x32.png',
	'./favicon.ico',
	'./pwa-icons/icon-192.png',
	'./pwa-icons/icon-512.png',
	'./pwa-icons/maskable-192.png',
	'./pwa-icons/maskable-512.png',
];

self.addEventListener('install', event => {
	event.waitUntil(
		(async () => {
			const cache = await caches.open(APP_SHELL_CACHE);
			await cache.addAll(APP_SHELL_ASSETS.map(asset => new URL(asset, self.registration.scope).toString()));
			await self.skipWaiting();
		})()
	);
});

self.addEventListener('activate', event => {
	event.waitUntil(
		(async () => {
			const cacheKeys = await caches.keys();
			await Promise.all(cacheKeys.filter(cacheName => !ACTIVE_CACHES.includes(cacheName)).map(cacheName => caches.delete(cacheName)));
			await self.clients.claim();
		})()
	);
});

self.addEventListener('fetch', event => {
	const { request } = event;
	if (request.method !== 'GET') {
		return;
	}
	if (request.cache === 'only-if-cached' && request.mode !== 'same-origin') {
		return;
	}

	const requestUrl = new URL(request.url);
	if (requestUrl.origin !== self.location.origin) {
		return;
	}

	if (request.mode === 'navigate') {
		event.respondWith(handleNavigationRequest(request));
		return;
	}

	if (shouldHandleStaticAsset(request)) {
		event.respondWith(handleStaticAssetRequest(event));
	}
});

function shouldHandleStaticAsset(request) {
	const destinations = new Set(['script', 'style', 'image', 'font', 'manifest']);
	if (destinations.has(request.destination)) {
		return true;
	}

	return /\.(?:css|js|png|jpg|jpeg|svg|webp|ico|webmanifest)$/i.test(new URL(request.url).pathname);
}

async function handleNavigationRequest(request) {
	try {
		const response = await fetch(request);
		return response;
	} catch (error) {
		const cache = await caches.open(APP_SHELL_CACHE);
		return (
			(await cache.match(new URL('./index.html', self.registration.scope).toString())) ||
			(await cache.match(new URL('./', self.registration.scope).toString()))
		);
	}
}

async function handleStaticAssetRequest(event) {
	const { request } = event;
	const cachedResponse = await caches.match(request);
	const fetchPromise = fetch(request)
		.then(async response => {
			if (!response || !response.ok) {
				return response;
			}

			const runtimeCache = await caches.open(RUNTIME_CACHE);
			await runtimeCache.put(request, response.clone());
			return response;
		})
		.catch(() => null);

	if (cachedResponse) {
		event.waitUntil(fetchPromise);
		return cachedResponse;
	}

	const networkResponse = await fetchPromise;
	if (networkResponse) {
		return networkResponse;
	}

	return Response.error();
}
