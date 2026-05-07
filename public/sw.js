// プール点検アプリ Service Worker
// ============================================================
// キャッシュ戦略
//   - 自己オリジンの HTML / JS / JSON  → Network First（最新優先）
//   - 自己オリジンのその他アセット      → Stale-While-Revalidate
//   - 外部CDN                          → Cache First（起動高速化）
//   - Firebase / Google API             → SWをバイパス（従来どおり）
// バージョンを更新すると activate 時に旧キャッシュを一括削除します。
// ============================================================

const APP_VERSION = 'v11-20260502-03';           // デプロイのたびに更新
const CACHE_NAME  = `pool-app-${APP_VERSION}`;

// プリキャッシュ対象（起動時に必要な静的ファイル）
const PRECACHE_URLS = [
  '/',
  '/index.html',
  '/manifest.json',
  '/favicon.png',
  '/icon-48.png',
  '/icon-192.png',
  '/icon-512.png',
  '/icon-maskable-192.png',
  '/icon-maskable-512.png',
  '/pool_mapping.json',
  '/weather_judge.js',
];

// Network First で扱う拡張子（更新頻度が高いもの）
function isNetworkFirstPath(pathname) {
  return (
    pathname === '/' ||
    pathname.endsWith('.html') ||
    pathname.endsWith('.js')   ||
    pathname.endsWith('.json')
  );
}

// ------------------------------------------------------------
// install : 新SWを即 waiting に、必要ファイルを個別フォールバックで投入
// ------------------------------------------------------------
self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await Promise.all(
      PRECACHE_URLS.map(async url => {
        try {
          // no-cache で取り直し、必ず最新をプリキャッシュ
          const resp = await fetch(url, { cache: 'no-cache' });
          if (resp && resp.status === 200) {
            await cache.put(url, resp.clone());
          }
        } catch (err) {
          console.warn('[SW] precache fail:', url, err);
        }
      })
    );
  })());
});

// ------------------------------------------------------------
// activate : 旧バージョンのキャッシュを一括削除 + 即 claim
// ------------------------------------------------------------
self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(
      keys
        .filter(k => k !== CACHE_NAME)
        .map(k => {
          console.log('[SW] Deleting old cache:', k);
          return caches.delete(k);
        })
    );
    // Navigation Preload が使える環境では有効化（初回表示を高速に）
    if (self.registration.navigationPreload) {
      try { await self.registration.navigationPreload.enable(); } catch (_) {}
    }
    await self.clients.claim();
  })());
});

// ------------------------------------------------------------
// message : クライアントからの指示に対応
//   - SKIP_WAITING : 手動で待機中SWを昇格
//   - CLEAR_CACHES : Cache Storage を全削除
// ------------------------------------------------------------
self.addEventListener('message', event => {
  const data = event.data || {};
  if (data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  if (data.type === 'CLEAR_CACHES') {
    event.waitUntil((async () => {
      const keys = await caches.keys();
      await Promise.all(keys.map(k => caches.delete(k)));
    })());
  }
});

// ------------------------------------------------------------
// fetch : リクエストを分類してルーティング
// ------------------------------------------------------------
self.addEventListener('fetch', event => {
  const req = event.request;

  // GET 以外（POST等）はSWを通さない
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Firebase / Google API 系はSWを介さない（従来どおり）
  if (url.href.includes('firebaseio.com')    ||
      url.href.includes('googleapis.com')    ||
      url.href.includes('identitytoolkit')   ||
      url.href.includes('securetoken')) {
    return;
  }

  // ナビゲーションリクエスト（アドレスバー遷移 / PWA起動）は Network First
  if (req.mode === 'navigate') {
    event.respondWith(networkFirst(req, event));
    return;
  }

  // 外部オリジン（SheetJS CDN など）は Cache First
  if (url.origin !== self.location.origin) {
    event.respondWith(cacheFirst(req));
    return;
  }

  // 自己オリジンの HTML / JS / JSON は Network First
  if (isNetworkFirstPath(url.pathname)) {
    event.respondWith(networkFirst(req, event));
    return;
  }

  // それ以外（アイコン・画像・CSS等）は Stale-While-Revalidate
  event.respondWith(staleWhileRevalidate(req));
});

// ============================================================
// 戦略実装（すべて async / await でノンブロッキング）
// ============================================================

// Network First : 最新を取りに行き、失敗時はキャッシュへフォールバック
async function networkFirst(req, event) {
  const cache = await caches.open(CACHE_NAME);
  try {
    // Navigation Preload があればそれを優先利用
    const preload = event && event.preloadResponse ? await event.preloadResponse : null;
    const fresh = preload || await fetch(req, { cache: 'no-store' });
    if (fresh && fresh.status === 200 && fresh.type !== 'opaque') {
      cache.put(req, fresh.clone()).catch(() => {});
    }
    return fresh;
  } catch (err) {
    const cached = await cache.match(req);
    if (cached) return cached;
    // ナビゲーションなら最終手段として index.html を返す
    if (req.mode === 'navigate') {
      const fallback = await cache.match('/index.html') || await cache.match('/');
      if (fallback) return fallback;
    }
    throw err;
  }
}

// Stale-While-Revalidate : まずキャッシュで高速表示し、裏で更新
async function staleWhileRevalidate(req) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(req);
  const networkPromise = fetch(req)
    .then(resp => {
      if (resp && resp.status === 200 && resp.type !== 'opaque') {
        cache.put(req, resp.clone()).catch(() => {});
      }
      return resp;
    })
    .catch(() => cached);
  return cached || networkPromise;
}

// Cache First : あればキャッシュ、なければネットワーク（取得後に保存）
async function cacheFirst(req) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(req);
  if (cached) return cached;
  try {
    const resp = await fetch(req);
    if (resp && resp.status === 200) {
      cache.put(req, resp.clone()).catch(() => {});
    }
    return resp;
  } catch (err) {
    if (cached) return cached;
    throw err;
  }
}
