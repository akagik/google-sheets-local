# Google Sheets ローカル操作スキル

Claude Code から Google Sheets API を使ってスプレッドシートの読み書きを行うスキルです。

## セットアップ

### 1. Google Cloud プロジェクトの準備

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成（または既存を使用）。
2. Google Sheets API を有効化。
3. 「認証情報」→「OAuth 2.0 クライアント ID」を作成（デスクトップアプリ）。
4. JSON をダウンロードし、`config/credentials.json` として配置。

### 2. Python 仮想環境のセットアップ (uv)

```bash
cd .claude/skills/google-sheets-local
uv venv .venv
uv pip install -r requirements.txt
```

> uv が未インストールの場合: `brew install uv`

### 3. 初回認証

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/auth_setup.py
```

ブラウザが開くので Google アカウントで認証してください。
`config/token.json` が生成されれば完了です。

## 使い方

Claude Code に対して、以下のように依頼するだけです:

- 「`<URL>` のシート名一覧を教えて」
- 「`<URL>` の key 名一覧を取得して」
- 「`<URL>` の A1:D10 を見せて」
- 「`<URL>` の Sheet1 に `["田中", "tanaka@example.com"]` を追加して」

## ファイル構成

```
google-sheets-local/
├── SKILL.md              # Claude Code が参照するスキル定義
├── README.md             # このファイル
├── scripts/
│   ├── auth_setup.py     # OAuth 認証セットアップ
│   └── sheets_tool.py    # Sheets API 操作 CLI
├── config/
│   ├── credentials.json  # OAuth クライアント秘密鍵（要配置）
│   └── token.json        # 認証後に自動生成
└── examples/
    └── sample_prompt.md  # プロンプト例
```

## 注意事項

- `config/credentials.json` と `config/token.json` は `.gitignore` に追加してください。
- トークンの有効期限が切れた場合は `auth_setup.py` を再実行してください。
