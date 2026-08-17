/* 領収書まとめ — サービスワーカー

   役割は2つだけです。
   1. ホーム画面から起動できるようにする
   2. 一度開いたあとは、通信できない場所でも起動して使えるようにする

   撮影した画像やデータをここで保存することはありません。
   キャッシュしているのは、画面を組み立てるためのファイルだけです。 */

const CACHE = "receipt-pwa-v1";

/* 画面の表示に必要なファイル。
   外部の部品は版を固定しています（版を上げるときは CACHE の名前も変えます）。 */
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/apple-touch-icon.png",
  "https://unpkg.com/react@18.3.1/umd/react.production.min.js",
  "https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js",
  "https://unpkg.com/@babel/standalone@7.29.8/babel.min.js",
  "https://cdn.tailwindcss.com/3.4.16",
  "https://unpkg.com/jspdf@2.5.2/dist/jspdf.umd.min.js",
  "https://unpkg.com/encoding-japanese@2.2.0/encoding.min.js",
];

/* 1件でも取得に失敗したら導入全体が失敗する、という状態は避けます。
   通信が不安定な場所で開いた場合でも、取れたものだけを残します。 */
async function precache() {
  const cache = await caches.open(CACHE);
  await Promise.all(
    ASSETS.map(async (url) => {
      let res = null;
      try {
        res = await fetch(new Request(url, { cache: "reload" }));
      } catch (e) {
        try {
          res = await fetch(new Request(url, { mode: "no-cors" }));
        } catch (e2) {
          res = null;
        }
      }
      if (res && (res.ok || res.type === "opaque")) {
        try {
          await cache.put(url, res);
        } catch (e) {
          /* 保存できなかった1件は次回の起動時に取り直します */
        }
      }
    })
  );
}

self.addEventListener("install", (event) => {
  event.waitUntil(precache().then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const names = await caches.keys();
      await Promise.all(
        names.map((n) => (n === CACHE ? null : caches.delete(n)))
      );
      await self.clients.claim();
    })()
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  /* 画面そのものの読み込み。通信できるときは新しいものを、
     できないときは保存してあるものを返します。 */
  if (req.mode === "navigate") {
    event.respondWith(
      (async () => {
        try {
          return await fetch(req);
        } catch (e) {
          const cache = await caches.open(CACHE);
          const hit =
            (await cache.match("./index.html")) || (await cache.match("./"));
          return hit || Response.error();
        }
      })()
    );
    return;
  }

  /* 部品は保存してあるものを優先します。起動が速くなり、圏外でも動きます。 */
  event.respondWith(
    (async () => {
      const cache = await caches.open(CACHE);
      const hit = await cache.match(req);
      if (hit) return hit;
      try {
        const res = await fetch(req);
        if (res && (res.ok || res.type === "opaque")) {
          cache.put(req, res.clone()).catch(() => {});
        }
        return res;
      } catch (e) {
        return Response.error();
      }
    })()
  );
});
