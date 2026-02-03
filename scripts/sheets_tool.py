"""
Google Sheets 操作ユーティリティ。

URL を渡すと spreadsheetId を抽出し、読み取り・書き込みを行う。
Claude Code から Python 実行で呼び出す前提。

使い方（CLI）:
    python scripts/sheets_tool.py headers  <URL> [シート名]
    python scripts/sheets_tool.py read     <URL> <範囲>
    python scripts/sheets_tool.py append   <URL> <シート名> <JSON配列>
    python scripts/sheets_tool.py update   <URL> <範囲> <JSON 2次元配列>
    python scripts/sheets_tool.py sheets   <URL>
    python scripts/sheets_tool.py filter   <URL> <シート名> <列名> <値>
"""

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


# ---------------------------------------------------------------------------
# 共通ヘルパー
# ---------------------------------------------------------------------------

def extract_sheet_id(url: str) -> str:
    """URL から spreadsheetId を抽出する。"""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not m:
        raise ValueError(f"Google Sheets の URL として認識できません: {url}")
    return m.group(1)


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


def read_range(url: str, range_: str) -> list[list[str]]:
    """指定範囲のセル値を 2 次元配列で返す。

    range_ の例: "Sheet1!A1:C10", "A1:F100"
    """
    sheet_id = extract_sheet_id(url)
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


def update_range(url: str, range_: str, values: list[list]) -> dict:
    """指定範囲のセルを上書き更新する。

    range_ の例: "Sheet1!A2:C2"
    values の例: [["val1", "val2", "val3"]]
    """
    sheet_id = extract_sheet_id(url)
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
# CLI エントリポイント
# ---------------------------------------------------------------------------

def _print_json(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2))


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]
    url = sys.argv[2]

    if cmd == "sheets":
        _print_json(list_sheet_names(url))

    elif cmd == "headers":
        sheet_name = sys.argv[3] if len(sys.argv) > 3 else None
        _print_json(get_header_keys(url, sheet_name))

    elif cmd == "read":
        if len(sys.argv) < 4:
            print("使い方: sheets_tool.py read <URL> <範囲>")
            sys.exit(1)
        _print_json(read_range(url, sys.argv[3]))

    elif cmd == "append":
        if len(sys.argv) < 5:
            print("使い方: sheets_tool.py append <URL> <シート名> '<JSON配列>'")
            sys.exit(1)
        sheet_name = sys.argv[3]
        row_values = json.loads(sys.argv[4])
        _print_json(append_row(url, sheet_name, row_values))

    elif cmd == "update":
        if len(sys.argv) < 5:
            print("使い方: sheets_tool.py update <URL> <範囲> '<JSON 2次元配列>'")
            sys.exit(1)
        range_ = sys.argv[3]
        values = json.loads(sys.argv[4])
        _print_json(update_range(url, range_, values))

    elif cmd == "filter":
        if len(sys.argv) < 6:
            print("使い方: sheets_tool.py filter <URL> <シート名> <列名> <値>")
            sys.exit(1)
        sheet_name = sys.argv[3]
        column_name = sys.argv[4]
        value = sys.argv[5]
        _print_json(filter_by_column(url, sheet_name, column_name, value))

    else:
        print(f"不明なコマンド: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
