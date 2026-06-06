"""
TikTok Analytics Discord Bot - 動的編集者管理版
@メンションで編集者を自動判別・シートを自動作成するボット
"""

import os
import re
import json
import base64
import asyncio
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
SUMMARY_SHEET        = os.environ.get("SUMMARY_SHEET", "合算")

JST = timezone(timedelta(hours=9))

# ── 色のローテーション（編集者が増えるたびに自動割り当て） ──
COLORS = [
    {"red": 0.18, "green": 0.62, "blue": 0.35},  # 緑
    {"red": 0.45, "green": 0.18, "blue": 0.69},  # 紫
    {"red": 0.20, "green": 0.45, "blue": 0.75},  # 青
    {"red": 0.85, "green": 0.35, "blue": 0.15},  # オレンジ
    {"red": 0.75, "green": 0.15, "blue": 0.35},  # ピンク
    {"red": 0.15, "green": 0.60, "blue": 0.60},  # シアン
]

# ── ヘッダー ──────────────────────────────────────────
HEADERS = [
    "画像読み込み日時",
    "動画投稿日",
    "動画タイトル",
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


def get_existing_sheets(service) -> dict:
    """既存シート名とIDの辞書を返す"""
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}


def ensure_summary_sheet(service):
    """合算シートがなければ作成"""
    existing = get_existing_sheets(service)
    if SUMMARY_SHEET not in existing:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {
                "title": SUMMARY_SHEET,
                "tabColor": {"red": 0.95, "green": 0.60, "blue": 0.07},
            }}}]},
        ).execute()
        _write_header(service, SUMMARY_SHEET, {"red": 0.95, "green": 0.60, "blue": 0.07})


def ensure_editor_sheet(service, editor_name: str) -> dict:
    """編集者シートがなければ自動作成して色を割り当てる"""
    existing = get_existing_sheets(service)

    if editor_name in existing:
        # 既存シートの色情報を返す（色インデックスで管理）
        editor_count = len([k for k in existing if k not in [SUMMARY_SHEET]])
        color = COLORS[(list(existing.keys()).index(editor_name) - 1) % len(COLORS)]
        return {"color": color}

    # 新規作成：既存の編集者シート数で色を決定
    editor_sheets = [k for k in existing if k not in [SUMMARY_SHEET]]
    color = COLORS[len(editor_sheets) % len(COLORS)]

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"addSheet": {"properties": {
            "title": editor_name,
            "tabColor": color,
        }}}]},
    ).execute()
    _write_header(service, editor_name, color)
    print(f"✅ 新しい編集者シートを作成: {editor_name}")
    return {"color": color}


def _write_header(service, sheet_name: str, bg_color: dict):
    """ヘッダー行を書いてスタイルを設定"""
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": [HEADERS]},
    ).execute()

    existing = get_existing_sheets(service)
    sheet_id = existing[sheet_name]

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": bg_color,
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


def color_editor_cell_in_summary(service, editor_name: str, color: dict):
    """合算シートの編集者列（D列）に色付きテキストを適用"""
    existing = get_existing_sheets(service)
    sheet_id = existing[SUMMARY_SHEET]

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SUMMARY_SHEET}!D:D",
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
                    "startColumnIndex": 3,  # D列
                    "endColumnIndex": 4,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": color,
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat",
            }
        }]},
    ).execute()


# ── @メンションから編集者名を取得 ────────────────────
def detect_editor(message: discord.Message):
    """メンションされたユーザーの表示名を返す"""
    if message.mentions:
        member = message.mentions[0]
        name = member.display_name or member.global_name or member.name
        return name
    return None


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
画面上部に動画タイトルと「2026/3/15に投稿」のような投稿日が表示されています。正確に読み取ってください。
{
  "type": "stats",
  "投稿日": "文字列（例: 2026/3/15）",
  "動画タイトル": "文字列（例: 集まれ愛知県民#なかま #ジオゲッサー）",
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
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 400,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": image_b64}},
                {"type": "text", "text": prompt},
            ],
        }],
    }

    async with httpx.AsyncClient(timeout=60) as http:
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
    try:
        service = get_sheets_service()
        ensure_summary_sheet(service)
        print("   合算シート確認: OK")
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

    editor_name = detect_editor(message)
    if not editor_name:
        await message.reply(
            "⚠️ 編集者のメンションが見つかりませんでした。\n"
            "投稿時に編集者を `@メンション` してください。"
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

        now_jst = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")

        row = [
            now_jst,
            stats.get("投稿日", ""),
            stats.get("動画タイトル", ""),
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
        editor_cfg = ensure_editor_sheet(service, editor_name)
        append_row(service, editor_name, row)
        append_row(service, SUMMARY_SHEET, row)
        color_editor_cell_in_summary(service, editor_name, editor_cfg["color"])

        sheet_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
        await processing_msg.edit(
            content=(
                f"✅ **{editor_name}** のデータを記録しました！\n"
                f"📊 動画視聴数: **{stats.get('動画視聴数')}**　"
                f"フォロワー増: **{stats.get('新規フォロワー数')}**　"
                f"視聴維持率: **{retention.get('視聴維持率')}**（{retention.get('時間')}地点）\n"
                f"📝 記録先: `{editor_name}` シート ＋ `{SUMMARY_SHEET}` シート\n"
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
