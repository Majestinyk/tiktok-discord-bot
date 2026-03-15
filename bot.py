"""
TikTok Analytics Discord Bot - 編集者別シート管理版
@メンションで編集者を判別し、個別シート＋合算シートに追記するボット
"""

import os
import re
import json
import base64
import discord
import httpx
from datetime import datetime, timezone, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── 設定 ─────────────────────────────────────────────
DISCORD_TOKEN        = os.environ.get("DISCORD_BOT_TOKEN", "")
ANTHROPIC_API_KEY    = os.environ.get("ANTHROPIC_API_KEY", "")
TARGET_CHANNEL_ID    = int(os.environ.get("TARGET_CHANNEL_ID", "0"))
SPREADSHEET_ID       = os.environ.get("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_FILE = os.environ.get("SERVICE_ACCOUNT_FILE", "service_account.json")

JST = timezone(timedelta(hours=9))

# ── 編集者設定 ────────────────────────────────────────
EDITORS = {
    "みゃも": {
        "sheet": "みゃも",
        "color": {"red": 0.18, "green": 0.62, "blue": 0.35},
        "color_hex": "2E9E59",
    },
    "まぜし": {
        "sheet": "まぜし",
        "color": {"red": 0.45, "green": 0.18, "blue": 0.69},
        "color_hex": "7330B0",
    },
}
SUMMARY_SHEET = os.environ.get("SUMMARY_SHEET", "合算")

# ── ヘッダー ──────────────────────────────────────────
HEADERS = [
    "画像読み込み日時",
    "動画投稿日",
    "編集者",
    "動画視聴数",
    "総再生時間",
    "平均視聴時間",
    "動画をフル視聴(%)",
    "新規フォロワー数",
    "推定報酬額",
    "視聴維持率・計測時間",
    "視聴維持率(%)",
]

# ── Google Sheets ─────────────────────────────────────
def get_credentials():
    google_creds_json = os.environ.get("GOOGLE_CREDENTIALS", "")
    if google_creds_json:
        info = json.loads(google_creds_json)
        return service_account.Credentials.from_service_account_info(
            info,
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
    return service_account.Credentials.from_service_account_file(
        os.environ.get("SERVICE_ACCOUNT_FILE", "service_account.json"),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )

def get_sheets_service():
    return build("sheets", "v4", credentials=get_credentials())


def ensure_sheets(service):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    existing = {s["properties"]["title"] for s in meta["sheets"]}

    requests = []
    if SUMMARY_SHEET not in existing:
        requests.append({"addSheet": {"properties": {
            "title": SUMMARY_SHEET,
            "tabColor": {"red": 0.95, "green": 0.60, "blue": 0.07},
        }}})
    for name, cfg in EDITORS.items():
        if cfg["sheet"] not in existing:
            requests.append({"addSheet": {"properties": {
                "title": cfg["sheet"],
                "tabColor": cfg["color"],
            }}})
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests},
        ).execute()

    for sheet_name in [SUMMARY_SHEET] + [c["sheet"] for c in EDITORS.values()]:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1:K1",
        ).execute()
        if not result.get("values"):
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{sheet_name}!A1",
                valueInputOption="RAW",
                body={"values": [HEADERS]},
            ).execute()
            _style_header(service, sheet_name)


def _style_header(service, sheet_name: str):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id = next(
        s["properties"]["sheetId"]
        for s in meta["sheets"]
        if s["properties"]["title"] == sheet_name
    )
    if sheet_name == SUMMARY_SHEET:
        bg = {"red": 0.95, "green": 0.60, "blue": 0.07}
    else:
        cfg = next((c for c in EDITORS.values() if c["sheet"] == sheet_name), None)
        bg = cfg["color"] if cfg else {"red": 0.3, "green": 0.3, "blue": 0.3}

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": bg,
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            }
        }]},
    ).execute()


def append_row(service, sheet_name: str, row: list):
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [row]},
    ).execute()


def color_editor_cell_in_summary(service, editor_cfg: dict):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id = next(
        s["properties"]["sheetId"]
        for s in meta["sheets"]
        if s["properties"]["title"] == SUMMARY_SHEET
    )
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SUMMARY_SHEET}!C:C",
    ).execute()
    last_row = len(result.get("values", [])) - 1

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": last_row,
                    "endRowIndex": last_row + 1,
                    "startColumnIndex": 2,  # C列（編集者）
                    "endColumnIndex": 3,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": editor_cfg["color"],
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat",
            }
        }]},
    ).execute()


# ── @メンションから編集者を判別 ──────────────────────
def detect_editor(message: discord.Message):
    for member in message.mentions:
        display = (member.display_name or "").lower()
        global_name = (member.global_name or "").lower()
        nick = (member.name or "").lower()
        for editor_name, cfg in EDITORS.items():
            key = editor_name.lower()
            if key in (display, global_name, nick):
                return editor_name, cfg

    text = message.content.lower()
    for editor_name, cfg in EDITORS.items():
        if f"@{editor_name.lower()}" in text:
            return editor_name, cfg

    return None, None


# ── 画像 → base64 ────────────────────────────────────
async def image_to_base64(url: str):
    async with httpx.AsyncClient() as http:
        r = await http.get(url)
        r.raise_for_status()
        media_type = r.headers.get("content-type", "image/png").split(";")[0]
        return base64.standard_b64encode(r.content).decode(), media_type


# ── Claude API で画像解析 ─────────────────────────────
async def extract_data(image_b64: str, media_type: str) -> dict:
    prompt = """この画像はTikTokの動画分析画面です。画像タイプを判定してJSONのみ返してください。

【タイプ1: 統計サマリー画面（数値KPI一覧）】
画面上部に「2026/3/15に投稿」のような投稿日が表示されています。正確に読み取ってください。
{
  "type": "stats",
  "投稿日": "文字列（例: 2026/3/15）",
  "動画視聴数": "数値文字列（例: 2600）",
  "総再生時間": "文字列（例: 14h:42m:9s）",
  "平均視聴時間": "文字列（例: 17.83s）",
  "動画をフル視聴": "文字列（例: 10.3%）",
  "新規フォロワー数": "数値文字列（例: 3）",
  "推定報酬額": "文字列（例: 円0）"
}

【タイプ2: 視聴維持率グラフ画面】
画面左下に「0:05 (32%)」のような形式でテキストが表示されています。括弧内の数字が視聴維持率です。この数字を注意深く読み取ってください。
{
  "type": "retention",
  "時間": "文字列（例: 0:05）",
  "視聴維持率": "文字列（括弧内の数字をそのまま、例: 32%）"
}

JSONのみ返してください。マークダウン・説明文は不要です。"""

    payload = {
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 400,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": image_b64}},
                {"type": "text", "text": prompt},
            ],
        }],
    }

    async with httpx.AsyncClient(timeout=30) as http:
        r = await http.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json=payload,
        )
        r.raise_for_status()

    raw = r.json()["content"][0]["text"].strip()
    raw = re.sub(r"^```json\s*|^```\s*|```$", "", raw, flags=re.MULTILINE).strip()
    return json.loads(raw)


# ── Discord クライアント ──────────────────────────────
intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)


@client.event
async def on_ready():
    print(f"✅ ボット起動: {client.user}")
    print(f"   監視チャンネルID : {TARGET_CHANNEL_ID}")
    print(f"   スプレッドシートID: {SPREADSHEET_ID}")
    print(f"   編集者設定: {list(EDITORS.keys())}")
    try:
        service = get_sheets_service()
        ensure_sheets(service)
        print("   シート初期化: OK")
    except Exception as e:
        print(f"   ⚠️ Sheets接続エラー: {e}")


@client.event
async def on_message(message: discord.Message):
    if message.author.bot:
        return
    if TARGET_CHANNEL_ID and message.channel.id != TARGET_CHANNEL_ID:
        return

    images = [a for a in message.attachments if a.content_type and a.content_type.startswith("image/")]
    if not images:
        return

    editor_name, editor_cfg = detect_editor(message)
    if not editor_name:
        await message.reply(
            "⚠️ 編集者のメンションが見つかりませんでした。\n"
            f"投稿時に `@みゃも` または `@まぜし` をメンションしてください。"
        )
        return

    if len(images) < 2:
        await message.reply("⚠️ 画像が1枚しか検出されませんでした。統計サマリーと視聴維持率の**2枚を同時**に投稿してください。")
        return

    processing_msg = await message.reply(f"⏳ **{editor_name}** の画像を解析中...")

    try:
        results = []
        for img in images[:2]:
            b64, media_type = await image_to_base64(img.url)
            data = await extract_data(b64, media_type)
            results.append(data)

        stats     = next((r for r in results if r.get("type") == "stats"), None)
        retention = next((r for r in results if r.get("type") == "retention"), None)

        if not stats or not retention:
            await processing_msg.edit(content="❌ 画像の種類を正しく判別できませんでした。統計サマリー画面と視聴維持率グラフ画面の2枚を送ってください。")
            return

        now_jst    = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
        posted_jst = message.created_at.astimezone(JST).strftime("%Y-%m-%d %H:%M:%S")

        row = [
            now_jst,
            stats.get("投稿日", ""),
            editor_name,
            stats.get("動画視聴数", ""),
            stats.get("総再生時間", ""),
            stats.get("平均視聴時間", ""),
            stats.get("動画をフル視聴", ""),
            stats.get("新規フォロワー数", ""),
            stats.get("推定報酬額", ""),
            retention.get("時間", ""),
            retention.get("視聴維持率", ""),
        ]

        service = get_sheets_service()
        append_row(service, editor_cfg["sheet"], row)
        append_row(service, SUMMARY_SHEET, row)
        color_editor_cell_in_summary(service, editor_cfg)

        sheet_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
        color_emoji = "🟢" if editor_cfg["color_hex"] == "2E9E59" else "🟣"
        await processing_msg.edit(
            content=(
                f"{color_emoji} **{editor_name}** のデータを記録しました！\n"
                f"📊 動画視聴数: **{stats.get('動画視聴数')}**　"
                f"フォロワー増: **{stats.get('新規フォロワー数')}**　"
                f"視聴維持率: **{retention.get('視聴維持率')}**（{retention.get('時間')}地点）\n"
                f"📝 記録先: `{editor_cfg['sheet']}` シート ＋ `{SUMMARY_SHEET}` シート\n"
                f"🔗 {sheet_url}"
            )
        )

    except Exception as e:
        await processing_msg.edit(content=f"❌ エラーが発生しました: `{e}`")
        raise


def main():
    if not DISCORD_TOKEN:
        raise ValueError("環境変数 DISCORD_BOT_TOKEN が未設定")
    if not ANTHROPIC_API_KEY:
        raise ValueError("環境変数 ANTHROPIC_API_KEY が未設定")
    if not SPREADSHEET_ID:
        raise ValueError("環境変数 SPREADSHEET_ID が未設定")
    client.run(DISCORD_TOKEN)


if __name__ == "__main__":
    main()
