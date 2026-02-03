"""
Google Sheets API OAuth 認証セットアップスクリプト。

初回のみターミナルから実行する:
    python scripts/auth_setup.py

ブラウザが開き Google アカウントで認証すると、
同ディレクトリに token.json が保存される。
以降は token.json を再利用して API を呼び出す。
"""

import os
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# このスクリプトがあるディレクトリを基準にパスを解決
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)
CREDENTIALS_PATH = os.path.join(SKILL_DIR, "config", "credentials.json")
TOKEN_PATH = os.path.join(SKILL_DIR, "config", "token.json")


def main():
    if not os.path.exists(CREDENTIALS_PATH):
        print(f"エラー: {CREDENTIALS_PATH} が見つかりません。")
        print("Google Cloud Console から OAuth クライアント ID の JSON をダウンロードし、")
        print("config/credentials.json として配置してください。")
        sys.exit(1)

    creds = None

    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("トークンの有効期限切れ。リフレッシュ中...")
            creds.refresh(Request())
        else:
            print("ブラウザで Google アカウント認証を行います...")
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_PATH, SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, "w") as token_file:
            token_file.write(creds.to_json())
        print(f"認証成功。トークンを保存しました: {TOKEN_PATH}")
    else:
        print("既存のトークンは有効です。再認証は不要です。")


if __name__ == "__main__":
    main()
