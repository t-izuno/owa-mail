# 開発ガイド

## 前提条件

- Node.js 18 以上
- npm

## セットアップ

```bash
git clone <repository-url>
cd owa-mail
npm install
npm install -g @playwright/cli@latest
```

## ビルド

```bash
npm run build
```

TypeScript ソース (`src/`) を `dist/` にコンパイルする。
通常の `npm install` 時は `prepare` スクリプトで自動ビルドされる。
`npm install --package-lock-only` のような lockfile 更新専用実行では `prepare` はスキップされる。

## 開発時の実行

```bash
# tsx で直接実行（ビルド不要）
npm run dev -- login
npm run dev -- list --unread
```

## テスト

```bash
npm test
```

vitest で `src/**/*.test.ts` を実行する。

## プロジェクト構成

```text
owa-mail/
├── src/
│   ├── index.ts          # メインソース（CLI + OWAクライアント）
│   └── index.test.ts     # ユニットテスト
├── dist/                 # ビルド出力（.gitignore）
├── package.json
├── tsconfig.json
└── .gitignore
```

## 技術スタック

| 項目 | 技術 |
| --- | --- |
| 言語 | TypeScript (ES2022, NodeNext) |
| ランタイム | Node.js |
| ブラウザ自動化 | @playwright/cli (optional) |
| テスト | vitest |
| Lint | strict + noUncheckedIndexedAccess |

## 設計方針

- `@playwright/cli` は `login` コマンドでのみ使用し、`optionalDependencies` として管理
- セッション情報は `~/.config/owa-mail/session.json` に `playwright-cli state-save`
  形式で保存
- `login` 以外のコマンドは保存済みクッキーで OWA service.svc を直接呼び出す
- 出力は全て JSON 形式で stdout に出力し、パイプや他ツールとの連携を容易にする

## コミット前チェック

```bash
npm run build && npm test
```

ビルドエラーとテスト失敗がないことを確認してからコミットする。
