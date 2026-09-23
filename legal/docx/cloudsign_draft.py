#!/usr/bin/env python3
"""
クラウドサイン Web API で「下書き」を4件作成するスクリプト（送信はしない）。

前提
  - クラウドサインの企業向けプラン（Web API が使えるプラン）で、管理画面から
    API クライアント ID を発行していること。
  - 各 Word ファイルを PDF に変換して legal/docx/pdf/ に同名（拡張子 .pdf）で置いてあること。
    クラウドサイン API がアップロードできるのは PDF のみ。
  - 環境変数 CLOUDSIGN_CLIENT_ID にクライアント ID を入れて実行する。
    サンドボックス環境を使う場合は CLOUDSIGN_BASE=https://api-sandbox.cloudsign.jp を指定する。

使い方
  CLOUDSIGN_CLIENT_ID=xxxx python3 cloudsign_draft.py            # 下書き作成
  CLOUDSIGN_CLIENT_ID=xxxx python3 cloudsign_draft.py --dry-run  # API を呼ばず内容だけ表示

このスクリプトは POST /documents/{id}（送信）を一切呼ばない。送信はクラウドサインの画面で、
押印欄の位置と宛先を確認してから手動で行う。

API リファレンス: https://app.swaggerhub.com/apis/CloudSign/cloudsign-web_api/
※ widgets（押印欄の入力項目）の widget_type の値と座標系は、上記リファレンスで確認してから
  WIDGET_TYPE_SEAL の値を直すこと。位置は画面上で調整する前提で、ここでは置かない。
"""
import os
import sys
import json
import urllib.request
import urllib.parse
import uuid

BASE = os.environ.get("CLOUDSIGN_BASE", "https://api.cloudsign.jp")
CLIENT_ID = os.environ.get("CLOUDSIGN_CLIENT_ID", "")
HERE = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(HERE, "pdf")
DRY_RUN = "--dry-run" in sys.argv

SENDER_NOTE = "子会社（株式会社イーフルケアパートナーズ）設立および代表取締役兼務に関する承認書類"

# 送信単位ごとの定義。participants は署名（押印）する相手。
# メールアドレスは株主間契約 別紙6.7 の通知先。実際に操作する担当者は先方に確認する。
DRAFTS = [
    {
        "title": "【EEFUL】子会社設立・代表取締役兼務に関する承認書（グロービス）",
        "message": (
            "株式会社EEFULホールディングスの森山です。\n"
            "株主間契約第2.3条第1項および第2.4条第2項に基づく通知書、承認書、\n"
            "ならびに株主総会の書面決議に関する提案書・同意書をお送りします。\n"
            "ご確認のうえ、承認書と同意書への押印をお願いいたします。"
        ),
        "files": [
            "01_通知書兼承認依頼書_投資者宛.pdf",
            "02_承認書_グロービス.pdf",
            "04_株主総会提案書.pdf",
            "07_同意書_GlobisFundVII.pdf",
            "08_同意書_グロービス7号ファンド.pdf",
        ],
        "participants": [
            {"email": "ryohei.minami@globis.com", "name": "南 良平",
             "organization": "グロービス・キャピタル・パートナーズ株式会社"},
        ],
    },
    {
        "title": "【EEFUL】子会社設立・代表取締役兼務に関する承認書（DBC）",
        "message": (
            "株式会社EEFULホールディングスの森山です。\n"
            "株主間契約第2.3条第1項および第2.4条第2項に基づく通知書、承認書、\n"
            "ならびに株主総会の書面決議に関する提案書・同意書をお送りします。\n"
            "ご確認のうえ、承認書と同意書への押印をお願いいたします。"
        ),
        "files": [
            "01_通知書兼承認依頼書_投資者宛.pdf",
            "03_承認書_DBC.pdf",
            "04_株主総会提案書.pdf",
            "09_同意書_DBC1号.pdf",
        ],
        "participants": [
            {"email": "shusuke.terada@dbcapital.jp", "name": "寺田 修輔",
             "organization": "DBC1号投資事業有限責任組合"},
        ],
    },
    {
        "title": "【EEFUL】株主総会書面決議 同意書（Future Identity Capital）",
        "message": (
            "株式会社EEFULホールディングスの森山です。\n"
            "会社法第319条第1項に基づく株主総会の書面決議について、提案書と同意書をお送りします。\n"
            "ご確認のうえ、同意書への押印をお願いいたします。"
        ),
        "files": [
            "04_株主総会提案書.pdf",
            "06_同意書_FutureIdentityCapital.pdf",
        ],
        "participants": [
            # メールアドレスは要確認
            {"email": os.environ.get("FIC_EMAIL", "TODO@example.com"), "name": "森山 博暢",
             "organization": "Future Identity Capital投資事業有限責任組合"},
        ],
    },
    {
        "title": "【EEFUL】株主総会書面決議 同意書・議事録（森山穂貴）",
        "message": "自社控え。提案書、同意書（森山穂貴）、臨時株主総会議事録。",
        "files": [
            "04_株主総会提案書.pdf",
            "05_同意書_森山穂貴.pdf",
            "10_臨時株主総会議事録.pdf",
        ],
        "participants": [
            {"email": os.environ.get("MORIYAMA_EMAIL", "h_moriyama@eeful-hd.com"), "name": "森山 穂貴",
             "organization": "株式会社EEFULホールディングス"},
        ],
    },
]


def call(method, path, token=None, form=None, files=None):
    """最小限の HTTP クライアント（外部ライブラリ不要）。"""
    url = BASE + path
    headers = {"Accept": "application/json"}
    if token:
        headers["Authorization"] = "Bearer " + token
    data = None
    if files:
        boundary = "----cs" + uuid.uuid4().hex
        body = b""
        for k, v in (form or {}).items():
            body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n").encode()
        for k, (fname, content, ctype) in files.items():
            body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"; filename=\"{fname}\"\r\n"
                     f"Content-Type: {ctype}\r\n\r\n").encode() + content + b"\r\n"
        body += f"--{boundary}--\r\n".encode()
        headers["Content-Type"] = "multipart/form-data; boundary=" + boundary
        data = body
    elif form is not None:
        data = urllib.parse.urlencode(form).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    with urllib.request.urlopen(req, timeout=60) as r:
        raw = r.read()
        return json.loads(raw) if raw else {}


def main():
    if not DRY_RUN and not CLIENT_ID:
        sys.exit("CLOUDSIGN_CLIENT_ID が未設定です。--dry-run で内容確認だけ行えます。")

    # ファイルの存在確認
    missing = [f for d in DRAFTS for f in d["files"] if not os.path.exists(os.path.join(PDF_DIR, f))]
    if missing:
        print("PDF が見つかりません（Word から PDF 保存して legal/docx/pdf/ に置いてください）:")
        for m in sorted(set(missing)):
            print("  -", m)
        if not DRY_RUN:
            sys.exit(1)

    token = None
    if not DRY_RUN:
        token = call("POST", "/token?client_id=" + urllib.parse.quote(CLIENT_ID))["access_token"]

    for d in DRAFTS:
        print("\n=== 下書き:", d["title"])
        if DRY_RUN:
            for f in d["files"]:
                print("  file:", f)
            for p in d["participants"]:
                print("  participant:", p["name"], p["email"], p["organization"])
            continue

        doc = call("POST", "/documents", token, form={
            "title": d["title"], "note": SENDER_NOTE, "message": d["message"], "can_transfer": "true",
        })
        doc_id = doc["id"]
        print("  document id:", doc_id)
        for f in d["files"]:
            with open(os.path.join(PDF_DIR, f), "rb") as fh:
                content = fh.read()
            res = call("POST", f"/documents/{doc_id}/files", token,
                       form={"name": f[:-4]}, files={"uploadfile": (f, content, "application/pdf")})
            print("  file id:", res.get("id"), f)
        for p in d["participants"]:
            res = call("POST", f"/documents/{doc_id}/participants", token, form={
                "email": p["email"], "name": p["name"], "organization": p["organization"],
                "language_code": "ja",
            })
            print("  participant id:", res.get("id"), p["name"])
        print("  -> 下書きのまま。押印欄の配置と送信はクラウドサインの画面で行う。")

    print("\n完了。送信は行っていません。")


if __name__ == "__main__":
    main()
