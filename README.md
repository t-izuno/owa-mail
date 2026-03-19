# owa-mail

OWA (Outlook Web Access) 経由で Office365 のメールを参照・操作する CLI ツール。
Playwright でブラウザログイン → セッションを保存し、以降は OWA service.svc を直接叩く。

## セットアップ

```bash
brew tap t-izuno/homebrew-tap
brew install owa-mail
brew install playwright-cli
```

Homebrew を使わない場合は、GitHub Releases から配布物を取得して `owa-mail` バイナリを `PATH` の通った場所に配置する。

## 使い方

### ログイン（初回・セッション切れ時）

```bash
owa-mail login
```

ブラウザが開くので Office365 にログインする。
ログイン完了後、セッションが `~/.config/owa-mail/session.json` に保存される。

### フォルダ一覧

```bash
owa-mail folders
```

### メール一覧

```bash
# 受信トレイの最新20件
owa-mail list

# 未読のみ
owa-mail list --unread

# 差出人でフィルタ
owa-mail list --from tanaka@example.com

# 件名でフィルタ
owa-mail list --subject "月次レポート"

# 組み合わせ（未読 + 差出人 + フォルダ指定）
owa-mail list --unread --from john@example.com --folder inbox --count 50
```

### メール本文

```bash
owa-mail read <item_id>
```

`item_id` は `owa-mail list` で取得できる。

### メール検索

```bash
owa-mail search "請求書"
owa-mail search "tanaka@example.com"
owa-mail search "キーワード" --folder 送信済み
```

### 添付ファイル取得

```bash
owa-mail attachment <item_id>
owa-mail attachment <item_id> --save-dir /tmp/mail
```

### 既読にする

```bash
# 単一
owa-mail mark-read <item_id>

# 複数まとめて
owa-mail mark-read <item_id1> <item_id2> <item_id3>
```

## エラーと対処

| エラー | 対処 |
| --- | --- |
| セッションが見つかりません | `owa-mail login` を実行 |
| X-OWA-CANARY が見つかりません | `owa-mail login` で再ログイン |
| セッション期限切れ | `owa-mail login` で再ログイン |
| フォルダが見つかりません | `owa-mail folders` で正確な名前を確認 |

## ライセンス

[MIT](LICENSE)
