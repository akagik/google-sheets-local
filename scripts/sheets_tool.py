"""
Google Sheets 操作ユーティリティ。

URL を渡すと spreadsheetId を抽出し、読み取り・書き込みを行う。
Claude Code から Python 実行で呼び出す前提。

使い方（CLI）:
    python scripts/sheets_tool.py headers      <URL> [シート名]
    python scripts/sheets_tool.py read         <URL> <範囲>
    python scripts/sheets_tool.py append       <URL> <シート名> <JSON配列>
    python scripts/sheets_tool.py update       <URL> <範囲> <JSON 2次元配列>
    python scripts/sheets_tool.py insert-rows  <URL> <シート名> <行番号> <JSON 2次元配列> [説明]
    python scripts/sheets_tool.py sheets       <URL>
    python scripts/sheets_tool.py filter       <URL> <シート名> <列名> <値>
    python scripts/sheets_tool.py gid-to-name  <URL>
    python scripts/sheets_tool.py notes        <URL> [シート名]
    python scripts/sheets_tool.py lookup       <検索キーワード>
    python scripts/sheets_tool.py registry
"""

import datetime
import json
import os
import re
import sys

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)
TOKEN_PATH = os.path.join(SKILL_DIR, "config", "token.json")
CHANGELOGS_DIR = os.path.join(SKILL_DIR, "changelogs")
REGISTRY_PATH = os.path.join(SKILL_DIR, "registry", "index.json")


# ---------------------------------------------------------------------------
# レジストリ（索引表）
# ---------------------------------------------------------------------------

def _load_registry() -> list[dict]:
    """registry/index.json を読み込んで返す。ファイルが無ければ空リスト。"""
    if not os.path.exists(REGISTRY_PATH):
        return []
    with open(REGISTRY_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def lookup(keyword: str) -> dict | None:
    """呼び名（keyword）から索引エントリを検索して返す。

    key の完全一致 → aliases の部分一致（大文字小文字無視）の順で検索。
    一致しなければ None。
    """
    entries = _load_registry()
    kw_lower = keyword.lower()

    # key の完全一致
    for entry in entries:
        if entry["key"] == keyword:
            return entry

    # aliases の部分一致
    for entry in entries:
        for alias in entry.get("aliases", []):
            if kw_lower in alias.lower() or alias.lower() in kw_lower:
                return entry

    return None


def list_registry() -> list[dict]:
    """登録済み全エントリの概要を返す。"""
    entries = _load_registry()
    return [
        {
            "key": e["key"],
            "aliases": e.get("aliases", []),
            "sheet_name": e.get("sheet_name", ""),
            "description": e.get("description", ""),
        }
        for e in entries
    ]


# ---------------------------------------------------------------------------
# 共通ヘルパー
# ---------------------------------------------------------------------------

def _sanitize_range(range_: str) -> str:
    r"""範囲文字列に含まれるシェル由来のエスケープ (\!) を除去する。"""
    return range_.replace("\\!", "!")


def extract_sheet_id(url: str) -> str:
    """URL から spreadsheetId を抽出する。"""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not m:
        raise ValueError(f"Google Sheets の URL として認識できません: {url}")
    return m.group(1)


def extract_gid(url: str) -> int | None:
    """URL から gid パラメータを抽出する。見つからなければ None。"""
    m = re.search(r"[#&?]gid=(\d+)", url)
    return int(m.group(1)) if m else None


def get_service():
    """認証済みの Sheets API サービスを返す。"""
    if not os.path.exists(TOKEN_PATH):
        raise FileNotFoundError(
            f"token.json が見つかりません ({TOKEN_PATH})。"
            "先に auth_setup.py を実行してください。"
        )
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    return build("sheets", "v4", credentials=creds).spreadsheets()


# ---------------------------------------------------------------------------
# 読み取り系
# ---------------------------------------------------------------------------

def list_sheet_names(url: str) -> list[str]:
    """スプレッドシート内のシート名一覧を返す。"""
    sheet_id = extract_sheet_id(url)
    service = get_service()
    meta = service.get(spreadsheetId=sheet_id).execute()
    return [s["properties"]["title"] for s in meta.get("sheets", [])]


def resolve_sheet_name_by_gid(url: str) -> str:
    """URL 内の gid からシート名を返す。gid が無い場合は最初のシート名を返す。"""
    sheet_id = extract_sheet_id(url)
    gid = extract_gid(url)
    service = get_service()
    meta = service.get(spreadsheetId=sheet_id).execute()
    sheets = meta.get("sheets", [])
    if not sheets:
        raise ValueError("スプレッドシートにシートが存在しません。")
    if gid is None:
        return sheets[0]["properties"]["title"]
    for s in sheets:
        if s["properties"]["sheetId"] == gid:
            return s["properties"]["title"]
    raise ValueError(f"gid={gid} に対応するシートが見つかりません。")


def get_header_keys(url: str, sheet_name: str | None = None) -> list[str]:
    """1 行目のヘッダー（key 名）を配列で返す。"""
    sheet_id = extract_sheet_id(url)
    service = get_service()
    range_ = f"'{sheet_name}'!1:1" if sheet_name else "1:1"
    result = service.values().get(
        spreadsheetId=sheet_id,
        range=range_,
    ).execute()
    values = result.get("values", [[]])
    return values[0] if values else []


def get_header_notes(url: str, sheet_name: str | None = None) -> list[dict]:
    """ヘッダー行の列名・型（2行目）・メモを取得して返す。

    返却: [{"name": "列名", "type": "int", "note": "メモ\n複数行"}, ...]
    メモや型が無い列も含む（空文字）。名前が空の列はスキップ。
    """
    sheet_id = extract_sheet_id(url)
    service = get_service()
    range_ = f"'{sheet_name}'!1:2" if sheet_name else "1:2"
    result = service.get(
        spreadsheetId=sheet_id,
        ranges=range_,
        fields="sheets.data.rowData.values(formattedValue,note)",
    ).execute()

    row_data = result["sheets"][0]["data"][0].get("rowData", [])
    headers = row_data[0].get("values", []) if len(row_data) > 0 else []
    type_row = row_data[1].get("values", []) if len(row_data) > 1 else []

    columns = []
    for i, cell in enumerate(headers):
        name = cell.get("formattedValue", "")
        if not name:
            continue
        note = cell.get("note", "")
        type_cell = type_row[i] if i < len(type_row) else {}
        type_str = type_cell.get("formattedValue", "")
        columns.append({"name": name, "type": type_str, "note": note})
    return columns


def read_range(url: str, range_: str) -> list[list[str]]:
    """指定範囲のセル値を 2 次元配列で返す。

    range_ の例: "Sheet1!A1:C10", "A1:F100"
    """
    sheet_id = extract_sheet_id(url)
    range_ = _sanitize_range(range_)
    service = get_service()
    result = service.values().get(
        spreadsheetId=sheet_id,
        range=range_,
    ).execute()
    return result.get("values", [])


def filter_by_column(url: str, sheet_name: str, column_name: str, value: str) -> list[dict]:
    """指定シートの全データを読み取り、指定列が value と一致する行だけを返す。

    結果はヘッダーをキーとした辞書のリスト。
    """
    sheet_id = extract_sheet_id(url)
    service = get_service()
    range_ = f"'{sheet_name}'"
    result = service.values().get(
        spreadsheetId=sheet_id,
        range=range_,
    ).execute()
    rows = result.get("values", [])
    if not rows:
        return []

    headers = rows[0]
    if column_name not in headers:
        raise ValueError(
            f"列 '{column_name}' が見つかりません。利用可能な列: {headers}"
        )
    col_idx = headers.index(column_name)

    matched = []
    for row in rows[1:]:
        if col_idx < len(row) and row[col_idx] == value:
            record = {}
            for i, h in enumerate(headers):
                record[h] = row[i] if i < len(row) else ""
            matched.append(record)
    return matched


# ---------------------------------------------------------------------------
# 書き込み系
# ---------------------------------------------------------------------------

def append_row(url: str, sheet_name: str, row_values: list) -> dict:
    """シート末尾に 1 行追加する。"""
    sheet_id = extract_sheet_id(url)
    service = get_service()
    range_ = f"'{sheet_name}'!A1"
    body = {"values": [row_values]}
    result = service.values().append(
        spreadsheetId=sheet_id,
        range=range_,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
    return result.get("updates", {})


def _resolve_sheet_id_by_name(spreadsheet_id: str, sheet_name: str) -> int:
    """シート名から sheetId (gid) を取得する。"""
    service = get_service()
    meta = service.get(spreadsheetId=spreadsheet_id).execute()
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == sheet_name:
            return s["properties"]["sheetId"]
    raise ValueError(f"シート '{sheet_name}' が見つかりません。")


def insert_rows(url: str, sheet_name: str, row_number: int, values: list[list]) -> dict:
    """指定行に新しい行を挿入してデータを書き込む。

    row_number: 1-indexed の行番号。その行の前に挿入される。
    values: 書き込むデータの 2 次元配列。
    """
    spreadsheet_id = extract_sheet_id(url)
    sid = _resolve_sheet_id_by_name(spreadsheet_id, sheet_name)
    service = get_service()
    num_rows = len(values)

    # 1. 空行を挿入
    insert_req = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": sid,
                        "dimension": "ROWS",
                        "startIndex": row_number - 1,  # 0-indexed
                        "endIndex": row_number - 1 + num_rows,
                    },
                    "inheritFromBefore": True,
                }
            }
        ]
    }
    service.batchUpdate(spreadsheetId=spreadsheet_id, body=insert_req).execute()

    # 2. 挿入した行にデータを書き込む
    write_range = f"'{sheet_name}'!A{row_number}"
    body = {"values": values}
    result = service.values().update(
        spreadsheetId=spreadsheet_id,
        range=write_range,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
    return result


def update_range(url: str, range_: str, values: list[list]) -> dict:
    """指定範囲のセルを上書き更新する。

    range_ の例: "Sheet1!A2:C2"
    values の例: [["val1", "val2", "val3"]]
    """
    sheet_id = extract_sheet_id(url)
    range_ = _sanitize_range(range_)
    service = get_service()
    body = {"values": values}
    result = service.values().update(
        spreadsheetId=sheet_id,
        range=range_,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
    return result


# ---------------------------------------------------------------------------
# Changelog
# ---------------------------------------------------------------------------

def _extract_sheet_name_from_range(range_: str) -> str:
    """範囲文字列からシート名を抽出する。"""
    r = _sanitize_range(range_)
    if "!" in r:
        return r.split("!")[0].strip("'")
    return "unknown"


def _resolve_gid_by_sheet_name(sheet_id: str, sheet_name: str) -> int | None:
    """シート名から gid を逆引きする。見つからなければ None。"""
    try:
        service = get_service()
        meta = service.get(spreadsheetId=sheet_id).execute()
        for s in meta.get("sheets", []):
            if s["properties"]["title"] == sheet_name:
                return s["properties"]["sheetId"]
    except Exception:
        pass
    return None


def _save_changelog(
    url: str,
    range_: str,
    operation: str,
    old_values=None,
    new_values=None,
    description: str | None = None,
) -> str:
    """変更ログを changelogs/ ディレクトリに Markdown ファイルとして保存する。"""
    os.makedirs(CHANGELOGS_DIR, exist_ok=True)

    now = datetime.datetime.now()
    timestamp_str = now.strftime("%Y%m%d_%H%M")
    sanitized_range = _sanitize_range(range_)
    sheet_name = _extract_sheet_name_from_range(range_)

    if description is None:
        description = f"{sheet_name}の{operation}"

    filename = f"{timestamp_str}_{description}.md"
    filepath = os.path.join(CHANGELOGS_DIR, filename)

    # ファイル名衝突を回避
    counter = 1
    while os.path.exists(filepath):
        filename = f"{timestamp_str}_{description}_{counter}.md"
        filepath = os.path.join(CHANGELOGS_DIR, filename)
        counter += 1

    sheet_id = extract_sheet_id(url)
    gid = _resolve_gid_by_sheet_name(sheet_id, sheet_name)
    if gid is not None:
        spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit?gid={gid}#gid={gid}"
    else:
        spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    lines = [
        f"# {description}",
        "",
        f"- **日時**: {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"- **操作**: {operation}",
        f"- **スプレッドシート**: [{sheet_name}]({spreadsheet_url})",
        f"- **範囲**: `{sanitized_range}`",
        "",
    ]

    if old_values is not None:
        lines += ["## 変更前", "", "```json", json.dumps(old_values, ensure_ascii=False, indent=2), "```", ""]

    lines += ["## 変更後", "", "```json", json.dumps(new_values, ensure_ascii=False, indent=2), "```", ""]

    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    return filepath


# ---------------------------------------------------------------------------
# CLI エントリポイント
# ---------------------------------------------------------------------------

def _print_json(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]

    # --- URL 不要のコマンド ---

    if cmd == "registry":
        _print_json(list_registry())
        return

    if cmd == "lookup":
        if len(sys.argv) < 3:
            print("使い方: sheets_tool.py lookup <検索キーワード>")
            sys.exit(1)
        keyword = sys.argv[2]
        result = lookup(keyword)
        if result:
            _print_json(result)
        else:
            entries = list_registry()
            print(f"エラー: '{keyword}' に一致するエントリが見つかりません。", file=sys.stderr)
            print("\n登録済みエントリ一覧:", file=sys.stderr)
            for e in entries:
                print(f"  - {e['key']}: {e['description']} (aliases: {', '.join(e['aliases'])})", file=sys.stderr)
            sys.exit(1)
        return

    # --- URL 必須のコマンド ---

    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    url = sys.argv[2]

    if cmd == "sheets":
        _print_json(list_sheet_names(url))

    elif cmd == "headers":
        sheet_name = sys.argv[3] if len(sys.argv) > 3 else None
        _print_json(get_header_keys(url, sheet_name))

    elif cmd == "notes":
        sheet_name = sys.argv[3] if len(sys.argv) > 3 else None
        _print_json(get_header_notes(url, sheet_name))

    elif cmd == "read":
        if len(sys.argv) < 4:
            print("使い方: sheets_tool.py read <URL> <範囲>")
            sys.exit(1)
        _print_json(read_range(url, sys.argv[3]))

    elif cmd == "append":
        if len(sys.argv) < 5:
            print("使い方: sheets_tool.py append <URL> <シート名> '<JSON配列>' [説明]")
            sys.exit(1)
        sheet_name = sys.argv[3]
        row_values = json.loads(sys.argv[4])
        desc = sys.argv[5] if len(sys.argv) > 5 else None
        result = append_row(url, sheet_name, row_values)
        log_range = f"'{sheet_name}'!append"
        filepath = _save_changelog(url, log_range, "append", None, [row_values], desc)
        result["changelog"] = filepath
        _print_json(result)

    elif cmd == "update":
        if len(sys.argv) < 5:
            print("使い方: sheets_tool.py update <URL> <範囲> '<JSON 2次元配列>' [説明]")
            sys.exit(1)
        range_ = sys.argv[3]
        values = json.loads(sys.argv[4])
        desc = sys.argv[5] if len(sys.argv) > 5 else None
        # 更新前の値を記録
        old_values = read_range(url, range_)
        result = update_range(url, range_, values)
        filepath = _save_changelog(url, range_, "update", old_values, values, desc)
        result["changelog"] = filepath
        _print_json(result)

    elif cmd == "insert-rows":
        if len(sys.argv) < 6:
            print("使い方: sheets_tool.py insert-rows <URL> <シート名> <行番号> '<JSON 2次元配列>' [説明]")
            sys.exit(1)
        sheet_name = sys.argv[3]
        row_number = int(sys.argv[4])
        values = json.loads(sys.argv[5])
        desc = sys.argv[6] if len(sys.argv) > 6 else None
        result = insert_rows(url, sheet_name, row_number, values)
        log_range = f"'{sheet_name}'!A{row_number}"
        filepath = _save_changelog(url, log_range, "insert-rows", None, values, desc)
        result["changelog"] = filepath
        _print_json(result)

    elif cmd == "filter":
        if len(sys.argv) < 6:
            print("使い方: sheets_tool.py filter <URL> <シート名> <列名> <値>")
            sys.exit(1)
        sheet_name = sys.argv[3]
        column_name = sys.argv[4]
        value = sys.argv[5]
        _print_json(filter_by_column(url, sheet_name, column_name, value))

    elif cmd == "gid-to-name":
        print(resolve_sheet_name_by_gid(url))

    else:
        print(f"不明なコマンド: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
