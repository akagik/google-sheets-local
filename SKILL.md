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
- 初回に `.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/auth_setup.py` を実行して `token.json` が作成済みであること。
- `.claude/skills/google-sheets-local/.venv/` に仮想環境が作成済みで、`uv pip install -r requirements.txt` でパッケージがインストール済みであること。

## Supported operations

| コマンド | 説明 |
|---------|------|
| `sheets` | スプレッドシート内のシート名一覧を取得 |
| `headers` | 指定シートの 1 行目（ヘッダー / key 名）を取得 |
| `read` | 指定範囲のセル値を取得 |
| `append` | シート末尾に 1 行追加 |
| `update` | 指定範囲のセルを上書き更新 |
| `insert-rows` | 指定行に新しい行を挿入してデータを書き込む（後続行は押し下げられる） |
| `filter` | 指定列が特定の値と一致する行だけを抽出（シート全体を走査） |
| `gid-to-name` | URL 内の gid パラメータからシート名を取得 |
| `notes` | ヘッダー行の列名・型（2行目）・メモを一括取得 |
| `lookup` | 呼び名からスプレッドシートの索引エントリを検索（URL 不要） |
| `registry` | 登録済み全スプレッドシートの一覧を表示（URL 不要） |



## レジストリ（索引表）

頻繁に使うスプレッドシートを「呼び名」で参照できる仕組み。

### ファイル構成

- `registry/index.json`: 全シートのメタ情報（key, aliases, URL, シート名, gid, 説明）
- `registry/sheets/<key>.md`: 各シートの詳細説明（列説明、注意事項、操作例）

### index.json の構造

```json
[
  {
    "key": "artifact-master",
    "aliases": ["アーティファクトテーブル", "ArtifactMaster"],
    "url": "https://docs.google.com/spreadsheets/d/.../edit?gid=123#gid=123",
    "sheet_name": "ArtifactMaster",
    "gid": 123,
    "description": "遠征アーティファクトのマスターデータ",
    "detail_file": "artifact-master.md"
  }
]
```

### 詳細ファイル（registry/sheets/*.md）の書き方

列名・型・メモは `notes` コマンドでスプレッドシートから直接取得できるため、md ファイルには二重管理しない。
md ファイルには **スプレッドシートのメモには書けない情報** だけを記載する。

- `# <シート名>` で始める
- `## 概要`: シートの役割を簡潔に
- `## 列情報の取得方法`: `notes` コマンドの実行例を記載
- `## 注意事項`: データ構造の癖やハマりどころ、列間の制約（併用不可など）、Deprecated な列の情報
- `## よく使う操作例`: 具体的なコマンド例

### 新しいシートの登録手順

1. `registry/index.json` にエントリを追加
2. `registry/sheets/<key>.md` に詳細ファイルを作成
3. `notes` コマンドで列情報を確認し、注意事項を md に記載

## How to handle a request

すべての操作は以下の形式で `sheets_tool.py` を Bash 経由で実行します:

```
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py <コマンド> [引数...]
```

### 0. ユーザーが呼び名でシートを指定した場合

ユーザーが URL ではなく「アーティファクトテーブル」「モンスターマスター」のような呼び名でシートを指定した場合:

1. `lookup` で URL とシート名を解決する:

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py lookup "アーティファクト"
```

2. `registry/sheets/` の詳細ファイルを Read で参照し、注意事項を確認する。
3. 列構造の確認が必要な場合は `notes` コマンドで型・メモを取得する。
4. 取得した URL を使って通常の `read` / `update` 等を実行する。

### 0a. 登録済みシート一覧を確認

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py registry
```

- 各エントリの key, aliases, sheet_name, description を一覧表示する。

### 1. シート名一覧を取得

ユーザーが「このスプレッドシートにどんなシートがあるか教えて」と依頼した場合:

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py sheets "<URL>"
```

### 2. ヘッダー（key 名）一覧を取得

ユーザーが「key 名一覧を取得して」と依頼した場合:

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py headers "<URL>" "シート名"
```

返却される JSON 配列をそのまま整形してユーザーに提示する。

**重要**: 列の構造を把握する必要がある場合（書き込み前、初めて触るシートなど）は、`headers` ではなく `notes` を使うこと。

### 2a. ヘッダーの列名・型・メモを一括取得（notes）

**シートの列情報を確認する際は、常にこのコマンドを使う。** `headers` は列名だけだが、`notes` は型とメモも一緒に取得できる。

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py notes "<URL>" "シート名"
```

返却は JSON 配列。各要素は `{"name": "列名", "type": "int", "note": "メモ内容"}` の形式。

- **type**: 2 行目のセル値をそのまま返す（例: `int`, `string`, `RarityType`, `AbilityType` 等）。2 行目が空なら空文字。
- **note**: ヘッダーセルに付与されたメモ。複数行の場合は `\n` 区切り。メモが無ければ空文字。
- 名前が空の列はスキップされる。

**使うべきタイミング**:
- 初めて操作するシートの列構造を把握するとき
- `append` / `insert-rows` / `update` の前に列の型・意味を確認するとき
- ユーザーにシートの構造を説明するとき

### 3. 指定範囲のデータを読み取る

ユーザーが「A1:C10 のデータを見せて」と依頼した場合:

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py read "<URL>" 'Sheet1!A1:C10'
```

- 範囲にシート名を含められる（例: `Sheet1!A1:C10`）。
- 返却は 2 次元 JSON 配列。テーブル形式に整形して返すと見やすい。

### 4. 行を追加（append）

ユーザーが「このシートに 1 行追加して」と依頼した場合:

1. まず `notes` コマンドでヘッダー・型・メモを取得して列順と各列の意味を確認する。
2. ユーザーが指定した値をヘッダー順の配列に変換する（未指定の列は空文字 `""` にする）。
3. 以下を実行:

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py append "<URL>" "シート名" '["値1", "値2", "値3"]'
```

4. 追加結果を確認してユーザーに報告する。

### 5. セル範囲を更新（update）

ユーザーが「A2:C2 を書き換えて」と依頼した場合:

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py update "<URL>" 'Sheet1!A2:C2' '[["新値1", "新値2", "新値3"]]'
```

- values は 2 次元配列（行×列）で渡す。
- 更新前に `read` で現在の値を確認し、変更内容をユーザーに確認してから実行すること。

#### 行番号の特定に関する重要な注意

`update` で特定の行を書き換える場合、**行番号の特定は慎重に行うこと**。
`read` の結果（配列）から行番号を目視でカウントすると off-by-one エラーを起こしやすい。

**推奨手順:**
1. 対象の列を `read` で取得する（例: `'Sheet!B1:B30'`）。
2. 結果の配列から対象値のインデックスを特定する（0-indexed）。
3. 行番号 = インデックス + 1（範囲が A1 始まりの場合）。
4. **更新実行前に、対象行の周辺数行を `read` で取得し、正しい行を指しているか必ず検証する。**

```
# 例: B列で "2001" を探して行番号を特定した後、更新前に必ず確認
read "ExpeditionDungeon!B11:I11"  # 対象行のデータを確認
# → dungeonId が "2001" であることを目視確認してから update を実行
```

**絶対にやってはいけないこと:**
- 配列の要素数を目視で数えて行番号を決定する（off-by-one の原因）
- 行番号の検証をせずに `update` を実行する（誤った行を破壊する危険がある）

### 6. 行を挿入（insert-rows）

既存データの間に新しい行を挿入したい場合に使う。`update` と違い、後続の行を押し下げて新しい行を差し込む。

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py insert-rows "<URL>" "シート名" <行番号> '<JSON 2次元配列>' [説明]
```

- **行番号**: 1-indexed。指定した行番号の位置に新しい行が挿入される（既存の行は下に押し下げられる）。
- **values**: 2 次元配列。複数行を一度に挿入可能。
- 後続の行やセクションヘッダーが壊れることなく、安全に行を差し込める。

例:
```bash
# シート "ArtifactMaster" の行 374 に 1 行挿入
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py insert-rows "<URL>" "ArtifactMaster" 374 '[["値1", "値2", "値3"]]' "新アーティファクト追加"
```

**`update` との使い分け:**
- 空行に書き込む場合 → `update` で OK
- 既存データの間に行を差し込む場合 → `insert-rows` を使う

### 7. 特定列の値でフィルタリング（filter）

ユーザーが「バージョン 3.10.0 のデータだけ取得して」のように、特定の列の値でデータを絞り込みたい場合:

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py filter "<URL>" "シート名" "列名" "値"
```

例:
```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py filter "<URL>" "ui" "バージョン" "3.10.0"
```

- シート全体のデータを取得し、指定列が値と **完全一致** する行だけを返す。
- 結果はヘッダーをキーとした辞書の JSON 配列で返却される。
- `read` コマンドのように範囲を指定する必要がないため、行数が不明な場合に便利。
- 指定した列名がヘッダーに存在しない場合はエラーになる。

### 8. URL の gid からシート名を取得（gid-to-name）

ユーザーが URL 付きで操作を依頼した場合、URL に `gid=` パラメータが含まれていれば、**最初に** このコマンドでシート名を特定する。
これにより全シートを順に検索する必要がなくなる。

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py gid-to-name "<URL>"
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

## Changelog（変更履歴の自動記録）

`update`、`append`、`insert-rows` を実行すると、変更内容が自動的にローカルの changelog ファイルに記録される。

- **保存先**: `.claude/skills/google-sheets-local/changelogs/`
- **ファイル名**: `YYYYMMDD_HHMM_<説明>.md`
- **記録内容**: 日時、操作種別、スプレッドシートへのリンク（gid付き）、範囲、変更前の値、変更後の値

### 説明の指定方法

`update` / `append` / `insert-rows` の最後の引数にオプションで説明を渡せる:

```bash
.claude/skills/google-sheets-local/.venv/bin/python .claude/skills/google-sheets-local/scripts/sheets_tool.py update "<URL>" 'Sheet!A1' '[["値"]]' "Expeditionのdepth更新"
```

説明を省略した場合はシート名と操作から自動生成される（例: `ExpeditionDungeon_update`）。

### 重要

- **`update` / `append` / `insert-rows` を呼ぶ際は、必ず変更内容がわかる説明を渡すこと。**
- changelog には変更前の値も記録されるため、万が一誤った更新をしても復元が可能。

## enum 値の書き込みに関する注意

スプレッドシートの列の型が enum 型（例: `RecommendType`, `RarityType`, `AbilityType` 等）の場合、
**数値ではなく enum の識別子（文字列）を書き込むこと。**

- OK: `"ExpeditionFastPass"`, `"MegaBundle"`, `"None"`
- NG: `"21"`, `"20"`, `"0"`

SheetSync のインポート時に enum 識別子から自動的に数値に変換されるため、
スプレッドシート側には常に人間が読める識別子を記入する。
型が enum かどうかは `notes` コマンドで 2 行目の型情報を確認すること。

## Safety & Error handling

- **認証エラー**: `token.json` が無いか期限切れの場合、ユーザーに `auth_setup.py` の再実行を案内する。
- **権限不足**: スプレッドシートの共有設定を確認してもらうよう伝える。
- **不正な URL**: Google Sheets 形式でない URL の場合は処理を行わずにエラーを伝える。
- **書き込み前の確認**: `append` や `update` を実行する前に、必ず変更内容をユーザーに提示し、確認を取ること。
- **大量データ**: 一度に大量のセルを取得しないよう、必要な範囲のみを指定するように誘導する。

## 致命的な禁止事項

### `append` コマンドで型行（2行目）やヘッダー行（1行目）を破壊してはならない

**2025-02-13 に発生したインシデント:**

`append` コマンドを使って localize_master (master-2) シートにデータを追加した際、
Google Sheets API の `values.append` が**シート末尾ではなく row 2（型の行）にデータを書き込み、
元の型情報を上書き・破壊してしまった**。

これにより SheetSync のインポートが正常に動作しなくなる致命的な問題が発生した。

**根本原因:**
- `append` API はヘッダー行の直後の空行を「テーブルの末尾」と判定する場合がある
- シートの構造（ヘッダー → 型行 → データ行）を事前に確認せず `append` を実行した

**再発防止策:**
1. **`append` は使わない。** 代わりに `read` で最終行を特定し、`update` で明示的な行番号に書き込む。
2. **書き込み前に必ず対象行の現在の内容を `read` で確認する。**
3. **row 1（ヘッダー行）と row 2（型行）は絶対に書き換えてはならない。**
4. **`append` の `updatedRange` レスポンスを必ず確認し、意図した行に書き込まれたかを検証する。**
