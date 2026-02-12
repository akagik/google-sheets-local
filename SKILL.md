---
name: google-sheets-local
description: |
  ローカル環境から Google Sheets API を用いて、指定されたスプレッドシートの内容を閲覧・編集するスキルです。
  「この URL の key 名一覧を取得して」「このシートに 1 行追加して」などの依頼に対して使用します。
version: 0.1.0
---

# Google Sheets ローカル操作スキル

## Overview

このスキルは、ユーザーが指定した Google スプレッドシート URL をもとに、
ローカル環境の Google Sheets API クライアント（Python）を使ってデータの読み書きを行います。

スクリプトの場所: `.claude/skills/google-sheets-local/scripts/sheets_tool.py`

## Preconditions

- `config/credentials.json` が `.claude/skills/google-sheets-local/config/` に配置されていること。
- 初回に `python .claude/skills/google-sheets-local/scripts/auth_setup.py` を実行して `token.json` が作成済みであること。
- Python パッケージ `google-api-python-client`, `google-auth-httplib2`, `google-auth-oauthlib` がインストール済みであること。

## Supported operations

| コマンド | 説明 |
|---------|------|
| `sheets` | スプレッドシート内のシート名一覧を取得 |
| `headers` | 指定シートの 1 行目（ヘッダー / key 名）を取得 |
| `read` | 指定範囲のセル値を取得 |
| `append` | シート末尾に 1 行追加 |
| `update` | 指定範囲のセルを上書き更新 |
| `filter` | 指定列が特定の値と一致する行だけを抽出（シート全体を走査） |
| `gid-to-name` | URL 内の gid パラメータからシート名を取得 |



## How to handle a request

すべての操作は以下の形式で `sheets_tool.py` を Bash 経由で実行します:

```
python .claude/skills/google-sheets-local/scripts/sheets_tool.py <コマンド> <URL> [引数...]
```

### 1. シート名一覧を取得

ユーザーが「このスプレッドシートにどんなシートがあるか教えて」と依頼した場合:

```bash
python .claude/skills/google-sheets-local/scripts/sheets_tool.py sheets "<URL>"
```

### 2. ヘッダー（key 名）一覧を取得

ユーザーが「key 名一覧を取得して」と依頼した場合:

```bash
# シート名指定なし（最初のシート）
python .claude/skills/google-sheets-local/scripts/sheets_tool.py headers "<URL>"

# シート名指定あり
python .claude/skills/google-sheets-local/scripts/sheets_tool.py headers "<URL>" "シート名"
```

返却される JSON 配列をそのまま整形してユーザーに提示する。

### 3. 指定範囲のデータを読み取る

ユーザーが「A1:C10 のデータを見せて」と依頼した場合:

```bash
python .claude/skills/google-sheets-local/scripts/sheets_tool.py read "<URL>" 'Sheet1!A1:C10'
```

- 範囲にシート名を含められる（例: `Sheet1!A1:C10`）。
- 返却は 2 次元 JSON 配列。テーブル形式に整形して返すと見やすい。

### 4. 行を追加（append）

ユーザーが「このシートに 1 行追加して」と依頼した場合:

1. まず `headers` コマンドでヘッダーを取得して列順を確認する。
2. ユーザーが指定した値をヘッダー順の配列に変換する（未指定の列は空文字 `""` にする）。
3. 以下を実行:

```bash
python .claude/skills/google-sheets-local/scripts/sheets_tool.py append "<URL>" "シート名" '["値1", "値2", "値3"]'
```

4. 追加結果を確認してユーザーに報告する。

### 5. セル範囲を更新（update）

ユーザーが「A2:C2 を書き換えて」と依頼した場合:

```bash
python .claude/skills/google-sheets-local/scripts/sheets_tool.py update "<URL>" 'Sheet1!A2:C2' '[["新値1", "新値2", "新値3"]]'
```

- values は 2 次元配列（行×列）で渡す。
- 更新前に `read` で現在の値を確認し、変更内容をユーザーに確認してから実行すること。

### 6. 特定列の値でフィルタリング（filter）

ユーザーが「バージョン 3.10.0 のデータだけ取得して」のように、特定の列の値でデータを絞り込みたい場合:

```bash
python .claude/skills/google-sheets-local/scripts/sheets_tool.py filter "<URL>" "シート名" "列名" "値"
```

例:
```bash
python .claude/skills/google-sheets-local/scripts/sheets_tool.py filter "<URL>" "ui" "バージョン" "3.10.0"
```

- シート全体のデータを取得し、指定列が値と **完全一致** する行だけを返す。
- 結果はヘッダーをキーとした辞書の JSON 配列で返却される。
- `read` コマンドのように範囲を指定する必要がないため、行数が不明な場合に便利。
- 指定した列名がヘッダーに存在しない場合はエラーになる。

### 7. URL の gid からシート名を取得（gid-to-name）

ユーザーが URL 付きで操作を依頼した場合、URL に `gid=` パラメータが含まれていれば、**最初に** このコマンドでシート名を特定する。
これにより全シートを順に検索する必要がなくなる。

```bash
python .claude/skills/google-sheets-local/scripts/sheets_tool.py gid-to-name "<URL>"
```

- URL 内の `gid` パラメータに対応するシート名を文字列で返す。
- `gid` が無い場合は最初のシート名を返す。
- **重要**: ユーザーが URL を指定した場合、`filter` や `read` の前にまずこのコマンドを実行してシート名を確定させること。

## フィルタリング時の注意事項

特定の列の値（バージョン番号など）でデータをフィルタリングする場合:

- **各行の値を厳密にチェックすること**。A列が空欄の行を「直前のバージョンと同じ」と仮定してはならない。
- 必ず **対象列に値が明示されている行だけ** を抽出する。
- セクション単位の切り出し（「次の別バージョンが現れるまで」）は行わない。シートの構造を事前に確認せずに推測でフィルタリングしない。
- フィルタリング前に、対象列の実データ構造（各行に値があるのか、セクション先頭のみか）を少量サンプルで確認することを推奨する。

## Safety & Error handling

- **認証エラー**: `token.json` が無いか期限切れの場合、ユーザーに `auth_setup.py` の再実行を案内する。
- **権限不足**: スプレッドシートの共有設定を確認してもらうよう伝える。
- **不正な URL**: Google Sheets 形式でない URL の場合は処理を行わずにエラーを伝える。
- **書き込み前の確認**: `append` や `update` を実行する前に、必ず変更内容をユーザーに提示し、確認を取ること。
- **大量データ**: 一度に大量のセルを取得しないよう、必要な範囲のみを指定するように誘導する。
