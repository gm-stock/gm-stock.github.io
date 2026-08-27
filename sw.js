// Service Worker - 항상 최신 버전 유지
const CACHE_NAME = 'gm-stock-v1';

// 설치 시 캐시 없이 네트워크 우선
self.addEventListener('install', event => {
  self.skipWaiting();
});

// 활성화 시 기존 캐시 전부 삭제
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.map(key => caches.delete(key)))
    ).then(() => self.clients.claim())
  );
});

// 요청 시 항상 네트워크 우선, 실패 시만 캐시
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request)
      .then(response => {
        // 성공하면 캐시에 저장
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
