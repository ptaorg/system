# PTAホームページ・入会申込・管理アプリ

PTAの任意加入、明示申込、学校名簿非流用を前提に、保護者用Googleサイト、入会申込Googleフォーム、役員用管理アプリをまとめて使うための配布用セットです。

## まず渡すURL

https://github.com/ptaorg/system

## このセットの完成形

```text
保護者
↓
Googleサイト
↓
入会申込Googleフォーム
↓
管理スプレッドシート
↓
役員用管理アプリ
```

保護者に見せるのは、Googleサイトと入会申込フォームだけです。
管理アプリURL、管理スプレッドシートURL、編集用フォームURLは保護者へ共有しません。

## Googleサイトはそのまま使える形にする

まず使うファイルはこれです。

```text
site/google-site-ready.html
```

この1ファイルに、次の内容をまとめています。

```text
トップページ
お知らせ
活動紹介
資料室
入会案内
重要事項
FAQ
お問い合わせ
```

Googleサイトでページを作り、`挿入 → 埋め込む → 埋め込みコード` に全文を貼れば使えます。

各PTAは、次の部分だけ差し替えます。

```text
○○学校PTA
年度
会費
支払い方法
問い合わせ先メール
FORM_URL_HERE
PDF資料リンク
```

## 管理アプリはアプリとして使う

役員側は、表を直接触るのではなく、Apps Scriptの管理アプリ画面を使います。

使うファイルはこの2つです。

```text
src/Code.gs
src/Code.webapp-addon.gs
src/Index.html
```

`Code.gs` が基本処理、`Code.webapp-addon.gs` が管理アプリ画面との接続、`Index.html` が役員用の画面です。

Apps Scriptに貼った後、Webアプリとしてデプロイすれば、役員用の管理アプリURLができます。

## 基本原則

- PTAへの加入は任意です。
- 加入しない方は、入会申込フォームを提出する必要はありません。
- 加入しないことにより、児童・生徒が学校教育活動上の不利益を受けることはありません。
- 学校名簿を流用せず、申込者本人が入力した情報だけで会員管理を行います。
- この仕組みは、非加入者情報を集めるためのものではありません。

## ファイル構成

```text
/
├── index.html                       # 配布説明ページ
├── README.md                        # この説明
├── site/
│   └── google-site-ready.html       # Googleサイトへそのまま貼る原本HTML
├── src/
│   ├── Code.gs                      # Apps Script本体
│   ├── Code.webapp-addon.gs         # 管理アプリ画面用の追加コード
│   ├── Index.html                   # 管理アプリ画面HTML
│   └── appsscript.json              # Apps Script設定例
└── docs/
    ├── GOOGLE_SITE_READY.md         # Googleサイト原本の使い方
    ├── SETUP.md                     # セットアップ手順
    ├── form-template.md             # Googleフォーム項目例
    ├── consent-template.md          # 重要事項説明テンプレート
    └── operation.md                 # 日常運用手順
```

## 最短手順

1. `site/google-site-ready.html` をGoogleサイトに埋め込む。
2. 入会申込Googleフォームを作る。
3. 管理スプレッドシートを作る。
4. スプレッドシートのApps Scriptに `src/Code.gs` を貼る。
5. 同じApps Scriptに `src/Code.webapp-addon.gs` も貼る。
6. HTMLファイル `Index` を作り、`src/Index.html` を貼る。
7. フォームURLをGoogleサイトの `FORM_URL_HERE` と差し替える。
8. Webアプリとしてデプロイする。
9. テスト申込を送る。
10. 管理アプリで会員確定できるか確認する。

## やってはいけないこと

- 加入しない方にフォーム提出を求めること。
- 非加入届として使うこと。
- 学校名簿をPTA会員名簿として流用すること。
- 申込のない保護者を会員として登録すること。
- 管理アプリURLや管理シートURLを保護者全体へ共有すること。
