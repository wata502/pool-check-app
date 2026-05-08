# プール点検PWA 完全インストール計画書

## 現状の問題点

Windows上で「インストールされているアプリ」には表示されるが、スタートメニューにアイコンが出ない。

## 原因分析

現在の `manifest.json` と アイコンに以下の問題がある。

### 1. アイコンが透過PNGになっている（角丸マスク付き）

現在のicon-192.png / icon-512.png は角が透明（RGBA）になっている。
Windowsのスタートメニューやタスクバーでは透過部分が黒くなったり、
アイコンとして認識されない場合がある。

**対策**: 背景を完全に塗りつぶした不透過（RGB）のアイコンを用意する。

### 2. `maskable` アイコンが未定義

Chromeは `purpose: "maskable"` のアイコンを使ってOS向けのアイコンを生成する。
現在は `"purpose": "any"` しかないため、OSがアイコンを正しく切り出せない。

**対策**: maskable用アイコン（セーフゾーン内に絵を収めたもの）を追加する。
maskableアイコンは上下左右10%がOSによって切り取られる可能性があるため、
中央80%の範囲に絵を配置する。

### 3. manifest.json に `id` フィールドがない

`id` はPWAの一意識別子。これがないとブラウザがアプリのIDを
URLベースで推測するため、再インストール時に別アプリと認識されることがある。

**対策**: `"id": "pool-inspection"` を追加する。

### 4. アイコンサイズが2種類しかない

192と512のみ。Windowsでは他のサイズ（48, 72, 96, 128, 256, 384）も
参照されることがある。最低限 48, 192, 512 の3サイズは必要。

**対策**: 追加サイズのアイコンを生成する。

---

## 修正計画

### Step 1: アイコンの再生成

タイトルなし.pngをベースに以下を作成する。

| ファイル名       | サイズ   | 用途       | 透過 |
|:----------------|:---------|:-----------|:-----|
| icon-48.png     | 48x48    | any        | なし |
| icon-192.png    | 192x192  | any        | なし |
| icon-512.png    | 512x512  | any        | なし |
| icon-maskable-192.png | 192x192 | maskable | なし |
| icon-maskable-512.png | 512x512 | maskable | なし |

- `any` アイコン: 正方形、角丸なし、背景ベタ塗り（#0277bd）、透過なし
- `maskable` アイコン: 正方形、背景ベタ塗り、絵を中央80%に縮小配置

### Step 2: manifest.json の更新

```json
{
  "id": "pool-inspection",
  "name": "プール点検",
  "short_name": "プール点検",
  "description": "プール付属電気設備点検入力アプリ",
  "start_url": "/",
  "scope": "/",
  "display": "standalone",
  "orientation": "portrait",
  "background_color": "#0277bd",
  "theme_color": "#0277bd",
  "icons": [
    { "src": "icon-48.png",  "sizes": "48x48",   "type": "image/png", "purpose": "any" },
    { "src": "icon-192.png", "sizes": "192x192",  "type": "image/png", "purpose": "any" },
    { "src": "icon-512.png", "sizes": "512x512",  "type": "image/png", "purpose": "any" },
    { "src": "icon-maskable-192.png", "sizes": "192x192", "type": "image/png", "purpose": "maskable" },
    { "src": "icon-maskable-512.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable" }
  ]
}
```

### Step 3: Service Worker キャッシュ更新

sw.js のキャッシュ名を `pool-app-v4` に更新し、新しいアイコンファイルも
キャッシュ対象に追加する。

### Step 4: デプロイ

`スマホアプリ更新.bat` を実行。

### Step 5: 既存PWAのアンインストールと再インストール

1. Windowsの「設定」→「アプリ」→「インストールされているアプリ」で
   「プール点検」をアンインストール
2. chrome://apps からも削除（あれば）
3. Chromeで pool-inspection-app.web.app を開く
4. アドレスバー右端のインストールアイコン（+マーク）をクリック
5. 「インストール」を押す
6. スタートメニューに水泳者アイコンが表示されることを確認
