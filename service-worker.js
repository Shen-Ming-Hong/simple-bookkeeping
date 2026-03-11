const CACHE_VERSION = 'v2';
const APP_SHELL_CACHE = `accounting-forecast-app-shell-${CACHE_VERSION}`;
const RUNTIME_CACHE = `accounting-forecast-runtime-${CACHE_VERSION}`;
const ACTIVE_CACHES = [APP_SHELL_CACHE, RUNTIME_CACHE];
const NETWORK_FIRST_ASSETS = ['./', './index.html', './app.js', './styles.css'].map(asset =>
	new URL(asset, self.registration.scope).pathname
);
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

	if (shouldHandleNetworkFirstAsset(request)) {
		event.respondWith(handleNetworkFirstAssetRequest(request));
		return;
	}

	if (shouldHandleStaticAsset(request)) {
		event.respondWith(handleStaticAssetRequest(event));
	}
});

function shouldHandleNetworkFirstAsset(request) {
	return NETWORK_FIRST_ASSETS.includes(new URL(request.url).pathname);
}

function shouldHandleStaticAsset(request) {
	const destinations = new Set(['script', 'style', 'image', 'font', 'manifest']);
	if (destinations.has(request.destination)) {
		return true;
	}

	return /\.(?:css|js|png|jpg|jpeg|svg|webp|ico|webmanifest)$/i.test(new URL(request.url).pathname);
}

async function handleNavigationRequest(request) {
	try {
		return await fetchAndCache(request, APP_SHELL_CACHE);
	} catch (error) {
		const cache = await caches.open(APP_SHELL_CACHE);
		return (
			(await cache.match(new URL('./index.html', self.registration.scope).toString())) ||
			(await cache.match(new URL('./', self.registration.scope).toString()))
		);
	}
}

async function handleNetworkFirstAssetRequest(request) {
	try {
		return await fetchAndCache(request, APP_SHELL_CACHE);
	} catch (error) {
		const cachedResponse = await caches.match(request);
		if (cachedResponse) {
			return cachedResponse;
		}

		return Response.error();
	}
}

async function handleStaticAssetRequest(event) {
	const { request } = event;
	const cachedResponse = await caches.match(request);
	const fetchPromise = fetchAndCache(request, RUNTIME_CACHE).catch(() => null);

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

async function fetchAndCache(request, cacheName) {
	const response = await fetch(request);
	if (!response || !response.ok) {
		return response;
	}

	const cache = await caches.open(cacheName);
	await cache.put(request, response.clone());
	return response;
}
