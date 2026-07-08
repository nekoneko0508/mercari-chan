import os
import json
import uuid
import base64
import io
import re
from datetime import date, datetime
import requests

import streamlit as st
import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from openai import OpenAI
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

load_dotenv()

SPREADSHEET_ID = "1q08itPY88CzG0yrQschTAMNRcVqy5-RwEASl4GkZor0"
WORKSHEET_NAME = "メルカリちゃん在庫"
DESCRIPTION_SHEET_NAME = "商品説明ログ"
PURCHASE_SHEET_NAME = "買付登録"
SALES_SHEET_NAME = "売上管理"
LINE_JUDGE_SHEET_NAME = "LINE買付ジャッジ状態"
RENDER_SERVICE_ACCOUNT_FILE = "/etc/secrets/service_account.json"
SERVICE_ACCOUNT_FILE = (
    RENDER_SERVICE_ACCOUNT_FILE
    if os.path.exists(RENDER_SERVICE_ACCOUNT_FILE)
    else "service_account.json"
)
DRIVE_FOLDER_ID = "1gNzzHYcjQcO7emNLWAQph9c8hIw-aXXG"
APPS_SCRIPT_PHOTO_URL = "https://script.google.com/macros/s/AKfycbzCY9tsIZuSqbVMyCuTZvhwYrRLXwvEzVYqcA2kiCtQZ_aqHskn9OfaylmGsgmbtNpT/exec"
PHOTO_SAVE_TOKEN = "mercari-chan-photo-save"

# ローカルでは .env、公開版では Streamlit Secrets からAPIキーを読む
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
try:
    if not OPENAI_API_KEY:
        OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", None)
except Exception:
    pass

client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

HEADERS = [
    "商品ID",
    "登録日時",
    "仕入日",
    "商品名",
    "カテゴリ",
    "仕入価格",
    "販売予定価格",
    "販売価格",
    "メルカリ手数料",
    "送料",
    "梱包資材費",
    "利益",
    "利益率",
    "在庫状況",
    "販売日",
    "メモ",
    "AI生成 商品説明文",
    "商品写真フォルダ",
    "早く売る価格",
    "標準価格",
    "高めに売る価格",
    "おすすめ販売価格",
    "価格理由",
    "",
    "販売先",
    "メルカリ商品ID",
    "メルカリ内プール売上金",
    "銀行入金状況",
    "freee登録状況",
    "仕入原価紐づけ状況",
    "発送状況",
    "在庫数"
]

DESCRIPTION_HEADERS = [
    "作成日時",
    "商品名",
    "カテゴリ",
    "ブランド",
    "カラー",
    "サイズ",
    "状態",
    "AI生成 タイトル",
    "AI生成 商品説明",
    "おすすめポイント",
    "ハッシュタグ",
    "ターゲット",
    "販売予定価格",
    "写真から見える特徴",
    "補足メモ",
    "早く売る価格",
    "標準価格",
    "高めに売る価格",
    "おすすめ販売価格",
    "価格理由"
]

PURCHASE_HEADERS = [
    "登録日時",
    "登録者名",
    "商品名",
    "購入価格",
    "色",
    "サイズ",
    "数量",
    "購入場所",
    "カテゴリ",
    "メモ",
    "商品写真ファイル名一覧",
    "商品写真フォルダ",
    "買付ステータス"
]

SALES_HEADERS = [
    "登録日時",
    "商品ID",
    "商品名",
    "販売日",
    "販売価格",
    "仕入価格",
    "メルカリ手数料",
    "送料",
    "梱包資材費",
    "利益",
    "利益率",
    "販売先",
    "在庫更新",
    "メモ"
]

LINE_JUDGE_HEADERS = [
    "LINEユーザーID",
    "状態",
    "画像URL",
    "商品特徴",
    "仕入価格",
    "判定結果",
    "更新日時"
]

def get_credentials():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    # ローカルでは service_account.json、RenderではSecret File、Streamlit CloudではSecretsから読む
    try:
        if "gcp_service_account" in st.secrets:
            return Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=scopes
            )
        else:
            return Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE,
                scopes=scopes
            )
    except Exception:
        return Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=scopes
        )

def get_spreadsheet():
    credentials = get_credentials()
    gc = gspread.authorize(credentials)
    return gc.open_by_key(SPREADSHEET_ID)

def column_number_to_letter(column_number):
    letters = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

def get_or_create_worksheet(sheet_name, headers):
    spreadsheet = get_spreadsheet()

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=sheet_name,
            rows=1000,
            cols=len(headers)
        )

    if worksheet.col_count < len(headers):
        worksheet.resize(cols=len(headers))

    values = worksheet.get_all_values()
    if not values:
        worksheet.append_row(headers)
    else:
        end_column = column_number_to_letter(len(headers))
        worksheet.update(f"A1:{end_column}1", [headers])

    return worksheet

def connect_sheet():
    return get_or_create_worksheet(WORKSHEET_NAME, HEADERS)

def connect_description_sheet():
    return get_or_create_worksheet(DESCRIPTION_SHEET_NAME, DESCRIPTION_HEADERS)

def connect_purchase_sheet():
    return get_or_create_worksheet(PURCHASE_SHEET_NAME, PURCHASE_HEADERS)

def connect_sales_sheet():
    return get_or_create_worksheet(SALES_SHEET_NAME, SALES_HEADERS)

def connect_line_judge_sheet():
    return get_or_create_worksheet(LINE_JUDGE_SHEET_NAME, LINE_JUDGE_HEADERS)

def generate_product_id():
    return "MC-" + uuid.uuid4().hex[:8].upper()

def reset_inventory_registration_form():
    st.session_state["reg_product_name"] = ""
    st.session_state["reg_category"] = ""
    st.session_state["reg_purchase_price"] = 0
    st.session_state["reg_purchase_date"] = date.today()
    st.session_state["reg_planned_price"] = 0
    st.session_state["reg_shipping_fee"] = 0
    st.session_state["reg_packing_fee"] = 0
    st.session_state["reg_memo"] = ""
    st.session_state["reg_stock_quantity"] = 1

def clean_json_text(text):
    text = text.strip()
    text = text.replace("```json", "").replace("```", "").strip()
    return text

def limit_title(text):
    text = str(text or "").strip()
    return text[:40]

def image_to_base64(uploaded_file):
    image_bytes = uploaded_file.getvalue()
    return base64.b64encode(image_bytes).decode("utf-8")

def save_uploaded_purchase_photo(uploaded_file):
    if uploaded_file is None:
        return "", ""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = uploaded_file.name.replace(" ", "_")
    file_name = f"purchase_{timestamp}_{safe_name}"

    base64_data = base64.b64encode(uploaded_file.getvalue()).decode("utf-8")

    payload = {
        "token": PHOTO_SAVE_TOKEN,
        "fileName": file_name,
        "mimeType": uploaded_file.type,
        "base64Data": base64_data
    }

    response = requests.post(APPS_SCRIPT_PHOTO_URL, json=payload, timeout=60)

    # Apps Scriptから返ってきた内容を確認しやすくする
    if response.status_code != 200:
        raise Exception(f"Apps Scriptエラー status={response.status_code}: {response.text[:500]}")

    try:
        result = response.json()
    except Exception:
        raise Exception(f"Apps ScriptからJSON以外が返ってきました: {response.text[:500]}")

    if not result.get("success"):
        raise Exception(result.get("error", "写真保存に失敗しました"))

    file_url = result.get("fileUrl", "")
    return file_name, file_url

def analyze_photo(uploaded_file, product_name, category):
    image_base64 = image_to_base64(uploaded_file)
    mime_type = uploaded_file.type

    prompt = f"""
あなたはメルカリ出品写真を見るプロです。

アップロードされた商品写真を見て、
商品説明に使える「写真から見える特徴」を日本語でまとめてください。

商品名: {product_name}
カテゴリ: {category}

見てほしいポイント:
- 色
- 形
- 素材感
- 付属品
- ロゴや柄の有無
- 傷、汚れ、箱つぶれなど気になる点
- 新品っぽいか、中古感があるか
- 購入者に伝えた方がよい注意点

注意:
- 写真から分からないことは断定しない
- 実物以上に良く見せる表現はしない
- メルカリの商品説明にそのまま活かせる言葉にする

出力は箇条書きで5〜8個程度にしてください。
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {
                "role": "system",
                "content": "あなたはフリマアプリの商品写真を見て、販売説明に使える特徴を整理する専門家です。"
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{mime_type};base64,{image_base64}"
                        }
                    }
                ]
            }
        ],
        temperature=0.4
    )

    return response.choices[0].message.content

def analyze_photos(uploaded_files, product_name, category):
    prompt = f"""
あなたはメルカリ出品写真を見るプロです。

アップロードされた複数の商品写真を見比べて、
商品説明に使える「写真から見える特徴」を日本語でまとめてください。

商品名: {product_name}
カテゴリ: {category}

写真ごとに見てほしいポイント:
1枚目: 商品全体写真
2枚目: ブランド名・商品名・タグが分かる写真
3枚目: サイズ・裏面・成分・状態が分かる写真
4〜5枚目: 傷・汚れ・付属品・パッケージ・裏面

総合的に見てほしいポイント:
- 色
- 形
- 素材感
- 付属品
- ロゴや柄の有無
- 傷、汚れ、箱つぶれなど気になる点
- 新品っぽいか、中古感があるか
- 購入者に伝えた方がよい注意点

注意:
- 写真から分からないことは断定しない
- 実物以上に良く見せる表現はしない
- メルカリの商品説明にそのまま活かせる言葉にする

出力は箇条書きで5〜10個程度にしてください。
"""

    content = [{"type": "text", "text": prompt}]
    for uploaded_file in uploaded_files:
        image_base64 = image_to_base64(uploaded_file)
        mime_type = uploaded_file.type
        content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{image_base64}"
                }
            }
        )

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {
                "role": "system",
                "content": "あなたは複数の商品写真を見比べて、フリマ出品に必要な特徴を整理する専門家です。"
            },
            {
                "role": "user",
                "content": content
            }
        ],
        temperature=0.4
    )

    return response.choices[0].message.content

def generate_description(
    product_name,
    category,
    brand,
    color,
    size,
    condition,
    features,
    photo_features,
    target,
    memo,
    planned_price,
    purchase_price=""
):
    prompt = f"""
あなたはメルカリで「5分で売れるタイトル・商品説明」と
現実的な販売価格を提案する専門家です。

以下の情報をもとに、メルカリ出品用のタイトル、商品説明、販売価格提案を作成してください。
必ずJSON形式のみで出力してください。

キーは必ず以下にしてください。
title
description
bullet_points
hashtags
quick_sale_price
standard_price
premium_price
recommended_price
price_reason

【商品情報】
商品名: {product_name}
カテゴリ: {category}
ブランド: {brand}
カラー: {color}
サイズ: {size}
状態: {condition}
商品の特徴: {features}
写真から見える特徴: {photo_features}
ターゲット: {target}
販売予定価格: {planned_price}
購入価格: {purchase_price}
補足メモ: {memo}

【タイトル条件】
- 必ず40文字以内
- メルカリ検索で見つかりやすい言葉を入れる
- 商品名、色、状態、用途、ターゲットのうち重要なものを入れる
- 「商品名＋魅力＋使う場面」の順番を基本にする
- 読んだ人が商品説明を開きたくなるタイトルにする
- ただし、釣りタイトルや誇大表現は禁止
- 「激安」「早い者勝ち」「絶対おすすめ」「新品同様」は使わない
- 文字を詰め込みすぎず、スマホで読みやすくする

【タイトル例】
ストロー付きタンブラー 持ち歩きに便利なブルー
外出用タンブラー ストロー付きで仕事中にも便利
ブルータンブラー ストロー付きでデスク使いにも
韓国風ポーチ 小物整理に便利な淡色デザイン
旅行用ミニバッグ 軽くて持ち歩きやすい黒

【商品説明条件】
- コンセプトは「5分で売れる商品説明」
- 冒頭2行で「欲しい」と思える魅力を伝える
- 何の商品か、どんな人におすすめかを分かりやすく書く
- 使う場面や使用シーンを具体的に入れる
- 購入後のイメージが浮かぶ文章にする
- 写真から見える特徴を自然に反映する
- 状態、サイズ、色、発送予定、注意点を整理する
- 状態の安心感を入れつつ、傷や汚れなどの注意点は正直に書く
- 購入前に確認してほしい点も自然に入れる
- そのままメルカリにコピペできる完成文にする
- 長すぎず、スマホで読みやすい文章にする
- 「絶対売れる」「新品同様」「完璧」などの断定・誇大表現は禁止
- 実物と違う印象を与えない
- 中古品や自宅保管品の場合は、自然な注意書きを入れる

【文章の型】
1. 冒頭：商品の魅力を2行で伝える
2. 商品の特徴
3. おすすめの使い方・おすすめの人
4. 状態・サイズ・発送について
5. 注意書き
6. 購入を後押しする一言

【bullet_points】
- 3〜5個
- 商品の魅力を短く箇条書き

【hashtags】
- 5〜8個
- メルカリ検索に合いそうな日本語ハッシュタグ

【価格提案条件】
- 類似商品を想定して、現実的な販売価格を提案する
- 写真から分かる状態、ブランド、カテゴリ、サイズ、付属品、注意点を考慮する
- 購入価格と販売予定価格がある場合は参考にする
- quick_sale_price は早く売りたい場合の価格
- standard_price は標準的に売れやすい価格
- premium_price は少し高めに狙う価格
- recommended_price は総合的におすすめする販売価格
- price_reason には、なぜその価格にしたかを2〜4文で書く
- 価格は日本円で「¥2,980」のように分かりやすく書く
- 相場を断定せず、「想定」「目安」として説明する
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {
                "role": "system",
                "content": "あなたはメルカリ出品文の専門家です。実物に忠実で、購入されやすい文章を作ります。"
            },
            {"role": "user", "content": prompt}
        ],
        temperature=0.7
    )

    text = response.choices[0].message.content
    text = clean_json_text(text)
    return json.loads(text)

def save_description_to_sheet(data):
    worksheet = connect_description_sheet()

    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        data.get("product_name", ""),
        data.get("category", ""),
        data.get("brand", ""),
        data.get("color", ""),
        data.get("size", ""),
        data.get("condition", ""),
        data.get("title", ""),
        data.get("description", ""),
        "\n".join(data.get("bullet_points", [])),
        " ".join(data.get("hashtags", [])),
        data.get("target", ""),
        data.get("planned_price", ""),
        data.get("photo_features", ""),
        data.get("memo", ""),
        data.get("quick_sale_price", ""),
        data.get("standard_price", ""),
        data.get("premium_price", ""),
        data.get("recommended_price", ""),
        data.get("price_reason", "")
    ]

    worksheet.append_row(row)

def make_safe_file_name(text):
    safe_text = text.strip()
    for character in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        safe_text = safe_text.replace(character, "_")
    safe_text = safe_text.replace(" ", "_")
    return safe_text or "商品名未入力"

def save_description_to_text_file(data):
    os.makedirs("outputs", exist_ok=True)

    product_name = make_safe_file_name(data.get("product_name", "商品名未入力"))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"{product_name}_{timestamp}.txt"
    file_path = os.path.join("outputs", file_name)

    content = [
        "タイトル",
        data.get("title", ""),
        "",
        "商品説明",
        data.get("description", ""),
        "",
        "おすすめポイント",
        "\n".join(f"- {item}" for item in data.get("bullet_points", [])),
        "",
        "ハッシュタグ",
        " ".join(data.get("hashtags", [])),
        "",
        "価格提案",
        f"早く売る価格: {data.get('quick_sale_price', '')}",
        f"標準価格: {data.get('standard_price', '')}",
        f"高めに売る価格: {data.get('premium_price', '')}",
        f"おすすめ販売価格: {data.get('recommended_price', '')}",
        "",
        "価格理由",
        data.get("price_reason", "")
    ]

    with open(file_path, "w", encoding="utf-8") as f:
        f.write("\n".join(content))

    return file_path

def save_uploaded_purchase_photos(product_id, product_name, uploaded_files):
    """複数写真も買付登録と同じApps Script経由でGoogle Driveへ保存する。"""
    if not uploaded_files:
        return [], ""

    saved_file_names = []
    saved_file_urls = []

    for index, uploaded_file in enumerate(uploaded_files, start=1):
        try:
            file_name, file_url = save_uploaded_purchase_photo(uploaded_file)
            if file_name:
                saved_file_names.append(file_name)
            if file_url:
                saved_file_urls.append(file_url)
        except Exception as e:
            original_name = getattr(uploaded_file, "name", f"{index}枚目の写真")
            raise Exception(f"{original_name} のGoogle Drive保存に失敗しました: {e}")

    return saved_file_names, ", ".join(saved_file_urls)

def make_google_sheets_hyperlink(photo_urls_text):
    """Google Sheetsでクリックできる写真リンク数式を作る。複数枚は1枚目を代表リンクにする。"""
    if not photo_urls_text:
        return ""

    if str(photo_urls_text).strip().startswith("=HYPERLINK"):
        return photo_urls_text

    photo_urls = [
        url.strip()
        for url in str(photo_urls_text).replace("\n", ",").split(",")
        if url.strip()
    ]
    if not photo_urls:
        return ""

    first_url = photo_urls[0].replace('"', '""')
    label = "写真1" if len(photo_urls) > 1 else "写真を開く"
    return f'=HYPERLINK("{first_url}","{label}")'

def save_purchase_to_sheet(data):
    worksheet = connect_purchase_sheet()
    photo_link_formula = make_google_sheets_hyperlink(data.get("photo_file_path", ""))

    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        data.get("buyer_name", ""),
        data.get("product_name", ""),
        data.get("purchase_price", ""),
        data.get("color", ""),
        data.get("size", ""),
        data.get("quantity", ""),
        data.get("place", ""),
        data.get("category", ""),
        data.get("memo", ""),
        data.get("photo_file_name", ""),
        photo_link_formula,
        "買付済み"
    ]

    # 必ずA列からM列に保存する
    all_values = worksheet.get_all_values()
    next_row = len(all_values) + 1

    # 1行目が空の場合は、A1からヘッダーを入れる
    if next_row == 1:
        worksheet.update(range_name="A1:M1", values=[PURCHASE_HEADERS])
        next_row = 2

    worksheet.update(
        range_name=f"A{next_row}:M{next_row}",
        values=[row],
        value_input_option="USER_ENTERED"
    )

def save_inventory_to_sheet(data):
    worksheet = connect_sheet()

    row = [
        data.get("product_id", ""),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        data.get("purchase_date", ""),
        data.get("product_name", ""),
        data.get("category", ""),
        data.get("purchase_price", ""),
        data.get("planned_price", ""),
        "",
        "",
        data.get("shipping_fee", ""),
        data.get("packing_fee", ""),
        "",
        "",
        "在庫あり",
        "",
        data.get("memo", ""),
        data.get("description", ""),
        data.get("photo_folder_path", ""),
        data.get("quick_sale_price", ""),
        data.get("standard_price", ""),
        data.get("premium_price", ""),
        data.get("recommended_price", ""),
        data.get("price_reason", ""),
        "",
        data.get("sales_channel", ""),
        data.get("mercari_item_id", ""),
        data.get("mercari_pool_sales", ""),
        data.get("bank_deposit_status", ""),
        data.get("freee_registration_status", ""),
        data.get("purchase_cost_link_status", ""),
        data.get("shipping_status", ""),
        data.get("stock_quantity", "")
    ]

    all_values = worksheet.get_all_values()
    next_row = len(all_values) + 1
    end_column = column_number_to_letter(len(HEADERS))
    worksheet.update(f"A{next_row}:{end_column}{next_row}", [row])
    return next_row

def update_inventory_description(row_number, description):
    worksheet = connect_sheet()
    worksheet.update(f"Q{row_number}", [[description]])

def update_inventory_price_suggestions(row_number, data):
    worksheet = connect_sheet()
    row = [
        data.get("quick_sale_price", ""),
        data.get("standard_price", ""),
        data.get("premium_price", ""),
        data.get("recommended_price", ""),
        data.get("price_reason", "")
    ]
    worksheet.update(f"S{row_number}:W{row_number}", [row])

def parse_yen_value(value):
    if value in (None, ""):
        return 0

    text = str(value).replace("¥", "").replace(",", "").strip()
    try:
        return int(float(text))
    except Exception:
        return 0

def format_yen(value):
    return f"¥{int(value):,}"

def format_profit_rate(profit_rate):
    return f"{profit_rate:.1%}" if profit_rate else "0.0%"

def calculate_sales_values(sale_price, purchase_price, shipping_fee, packing_fee):
    mercari_fee = int(sale_price * 0.10)
    profit = sale_price - purchase_price - mercari_fee - shipping_fee - packing_fee
    profit_rate = profit / sale_price if sale_price else 0
    return mercari_fee, profit, profit_rate

HIGH_RISK_PURCHASE_KEYWORDS = [
    "食品",
    "飲料",
    "健康食品",
    "サプリ",
    "サプリメント",
    "ハーブ",
    "パウダー",
    "モリンガ",
    "スーパーフード",
    "栄養",
    "栄養補助",
    "ビタミン",
    "ミネラル",
    "ダイエット",
    "デトックス",
    "美容",
    "健康",
    "改善",
    "治療",
    "予防",
    "効果",
    "効能",
    "免疫",
    "疲労回復",
    "医薬品",
    "医薬部外品",
    "薬",
    "クリーム",
    "化粧品",
    "コスメ",
    "小分け",
    "成分不明",
    "賞味期限",
    "消費期限",
    "海外製",
]

HIGH_RISK_PURCHASE_CATEGORIES = [
    "食品",
    "飲料",
    "健康食品",
    "サプリメント",
    "ハーブパウダー",
    "モリンガ",
    "スーパーフード",
    "栄養補助系商品",
    "医薬品",
    "医薬部外品",
    "医薬品と誤認される商品",
    "化粧品",
    "小分け化粧品",
    "海外製で成分・表示・賞味期限が確認しづらい商品",
    "健康、美容、治療、改善、効能効果をうたいやすい商品",
    "メルカリで削除・利用制限につながる可能性が高い商品",
]

FOOD_HEALTH_KEYWORDS = [
    "食品",
    "飲料",
    "健康食品",
    "サプリ",
    "ハーブ",
    "パウダー",
    "モリンガ",
    "スーパーフード",
    "栄養",
    "ビタミン",
    "ミネラル",
]

MEDICAL_COSMETIC_KEYWORDS = [
    "医薬品",
    "医薬部外品",
    "薬",
    "クリーム",
    "化粧品",
    "コスメ",
    "小分け",
]

EFFICACY_CLAIM_KEYWORDS = [
    "ダイエット",
    "デトックス",
    "美容",
    "健康",
    "改善",
    "治療",
    "予防",
    "効果",
    "効能",
    "免疫",
    "疲労回復",
]

IMPORT_LABEL_RISK_KEYWORDS = [
    "海外製",
    "輸入",
    "成分不明",
    "表示",
    "賞味期限",
    "消費期限",
]

def find_matched_keywords(text, keywords):
    normalized_text = str(text or "").lower()
    return [
        keyword
        for keyword in keywords
        if keyword.lower() in normalized_text
    ]

def judge_purchase_policy_risk(product_name="", category="", memo="", description=""):
    target_text = "\n".join(
        str(value or "")
        for value in [product_name, category, memo, description]
    )
    high_risk_matches = find_matched_keywords(target_text, HIGH_RISK_PURCHASE_KEYWORDS)
    food_health_matches = find_matched_keywords(target_text, FOOD_HEALTH_KEYWORDS)
    medical_cosmetic_matches = find_matched_keywords(target_text, MEDICAL_COSMETIC_KEYWORDS)
    efficacy_matches = find_matched_keywords(target_text, EFFICACY_CLAIM_KEYWORDS)
    import_label_matches = find_matched_keywords(target_text, IMPORT_LABEL_RISK_KEYWORDS)

    is_high_risk = bool(high_risk_matches)
    return {
        "sale_permission": "販売不可" if is_high_risk else "販売可",
        "policy_risk": "高" if is_high_risk else "低",
        "high_risk_category": is_high_risk,
        "food_health": bool(food_health_matches),
        "medical_cosmetic": bool(medical_cosmetic_matches),
        "efficacy_claim": bool(efficacy_matches),
        "import_label": bool(import_label_matches),
        "matched_keywords": high_risk_matches,
        "food_health_matches": food_health_matches,
        "medical_cosmetic_matches": medical_cosmetic_matches,
        "efficacy_matches": efficacy_matches,
        "import_label_matches": import_label_matches,
    }

def format_policy_risk_block(policy_risk):
    matched_keywords = "、".join(policy_risk.get("matched_keywords") or ["該当なし"])
    return "\n".join([
        f"販売可否確認: {policy_risk['sale_permission']}",
        f"規約リスク: {policy_risk['policy_risk']}",
        f"削除リスク高カテゴリ該当: {'はい' if policy_risk['high_risk_category'] else 'いいえ'}",
        f"食品・健康系該当: {'はい' if policy_risk['food_health'] else 'いいえ'}",
        f"医薬品・医薬部外品・化粧品該当: {'はい' if policy_risk['medical_cosmetic'] else 'いいえ'}",
        f"効能効果表現リスク: {'はい' if policy_risk['efficacy_claim'] else 'いいえ'}",
        f"海外輸入時の表示リスク: {'はい' if policy_risk['import_label'] else 'いいえ'}",
        f"該当キーワード: {matched_keywords}",
    ])

def build_policy_rejected_purchase_message(policy_risk):
    if policy_risk.get("food_health"):
        risk_reason = "食品・健康食品・ハーブパウダー系に該当する可能性があり、メルカリで削除対象になるリスクが高いためです。"
    elif policy_risk.get("medical_cosmetic"):
        risk_reason = "医薬品・医薬部外品・化粧品系に該当する可能性があり、メルカリで削除対象になるリスクが高いためです。"
    elif policy_risk.get("efficacy_claim"):
        risk_reason = "健康、美容、治療、改善などの効能効果表現につながる可能性があり、メルカリで削除対象になるリスクが高いためです。"
    elif policy_risk.get("import_label"):
        risk_reason = "海外輸入時の成分・表示・賞味期限確認が必要になり、メルカリで削除対象になるリスクが高いためです。"
    else:
        risk_reason = "メルカリで削除対象になるリスクが高いカテゴリに該当する可能性があるためです。"

    matched_keywords = policy_risk.get("matched_keywords") or []
    product_candidate = "対象商品"
    if "モリンガ" in matched_keywords:
        product_candidate = "モリンガパウダー"

    return "\n".join([
        "買付判定：買付不可",
        f"商品候補：{product_candidate}",
        "",
        "理由：",
        risk_reason,
        "利益が出そうでも、販売不可・削除・損失につながるため買付しないでください。",
        "",
        "この商品は登録せず、次の商品を確認してください。",
    ])

def parse_mercari_purchase_email(email_text):
    text = str(email_text or "")

    buyer_match = re.search(r"下記の商品を\s*(.+?さん)が購入しました", text)
    item_id_match = re.search(r"商品ID\s*[:：]\s*([A-Za-z0-9_-]+)", text)
    item_name_match = re.search(r"商品名\s*[:：]\s*(.+)", text)
    price_match = re.search(r"商品価格\s*[:：]\s*([\d,]+)\s*円?", text)

    sale_price = 0
    if price_match:
        sale_price = parse_yen_value(price_match.group(1))

    return {
        "buyer_name": buyer_match.group(1).strip() if buyer_match else "",
        "mercari_item_id": item_id_match.group(1).strip() if item_id_match else "",
        "item_name": item_name_match.group(1).strip() if item_name_match else "",
        "sale_price": sale_price
    }

def is_bundle_item_name(item_name):
    bundle_keywords = ["まとめ商品", "リクエスト", "2点", "3点", "複数"]
    return any(keyword in str(item_name or "") for keyword in bundle_keywords)

def get_available_inventory_items():
    worksheet = connect_sheet()
    rows = worksheet.get_all_values()
    items = []

    for row_number, row in enumerate(rows[1:], start=2):
        values = row + [""] * max(0, len(HEADERS) - len(row))
        if values[13] != "在庫あり":
            continue

        items.append(
            {
                "row_number": row_number,
                "product_id": values[0],
                "product_name": values[3],
                "purchase_price": parse_yen_value(values[5]),
                "planned_price": parse_yen_value(values[6])
            }
        )

    return items

def save_sales_registration(item, sale_date, sale_price, shipping_fee, packing_fee, sales_channel, memo):
    sales_sheet = connect_sales_sheet()
    inventory_sheet = connect_sheet()
    mercari_fee, profit, profit_rate = calculate_sales_values(
        sale_price,
        item["purchase_price"],
        shipping_fee,
        packing_fee
    )
    profit_rate_text = format_profit_rate(profit_rate)

    sales_row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        item["product_id"],
        item["product_name"],
        str(sale_date),
        sale_price,
        item["purchase_price"],
        mercari_fee,
        shipping_fee,
        packing_fee,
        profit,
        profit_rate_text,
        sales_channel,
        "売約済み更新済み",
        memo
    ]

    sales_values = sales_sheet.get_all_values()
    sales_next_row = len(sales_values) + 1
    if sales_next_row == 1:
        sales_sheet.update(range_name="A1:N1", values=[SALES_HEADERS])
        sales_next_row = 2

    sales_sheet.update(
        range_name=f"A{sales_next_row}:N{sales_next_row}",
        values=[sales_row],
        value_input_option="USER_ENTERED"
    )

    inventory_update_row = [
        sale_price,
        mercari_fee,
        shipping_fee,
        packing_fee,
        profit,
        profit_rate_text,
        "売約済み",
        str(sale_date),
        memo
    ]
    inventory_sheet.update(
        range_name=f"H{item['row_number']}:P{item['row_number']}",
        values=[inventory_update_row],
        value_input_option="USER_ENTERED"
    )

    return {
        "mercari_fee": mercari_fee,
        "profit": profit,
        "profit_rate": profit_rate_text,
        "sales_row": sales_next_row,
        "inventory_row": item["row_number"]
    }

def save_mercari_email_sales_registration(
    selected_items,
    parsed_email,
    sale_date,
    sale_price,
    shipping_fee,
    packing_fee,
    user_memo
):
    if not selected_items:
        raise ValueError("在庫商品が選択されていません。")

    sales_sheet = connect_sales_sheet()
    inventory_sheet = connect_sheet()

    purchase_price_total = sum(item["purchase_price"] for item in selected_items)
    mercari_fee, profit, profit_rate = calculate_sales_values(
        sale_price,
        purchase_price_total,
        shipping_fee,
        packing_fee
    )
    profit_rate_text = format_profit_rate(profit_rate)
    product_ids = ", ".join(item["product_id"] for item in selected_items)
    memo = (
        f"メルカリ商品ID：{parsed_email.get('mercari_item_id', '')}"
        f"／購入者：{parsed_email.get('buyer_name', '')}"
        f"／メール商品名：{parsed_email.get('item_name', '')}"
    )
    if user_memo:
        memo += f"／売上登録メモ：{user_memo}"

    sales_row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        product_ids,
        parsed_email.get("item_name", ""),
        str(sale_date),
        sale_price,
        purchase_price_total,
        mercari_fee,
        shipping_fee,
        packing_fee,
        profit,
        profit_rate_text,
        "メルカリ",
        "売約済み更新済み",
        memo
    ]

    sales_values = sales_sheet.get_all_values()
    sales_next_row = len(sales_values) + 1
    if sales_next_row == 1:
        sales_sheet.update(range_name="A1:N1", values=[SALES_HEADERS])
        sales_next_row = 2

    sales_sheet.update(
        range_name=f"A{sales_next_row}:N{sales_next_row}",
        values=[sales_row],
        value_input_option="USER_ENTERED"
    )

    if len(selected_items) == 1:
        item = selected_items[0]
        inventory_update_row = [
            sale_price,
            mercari_fee,
            shipping_fee,
            packing_fee,
            profit,
            profit_rate_text,
            "売約済み",
            str(sale_date),
            memo
        ]
        inventory_sheet.update(
            range_name=f"H{item['row_number']}:P{item['row_number']}",
            values=[inventory_update_row],
            value_input_option="USER_ENTERED"
        )
    else:
        for item in selected_items:
            inventory_sheet.update(
                range_name=f"N{item['row_number']}:P{item['row_number']}",
                values=[["売約済み", str(sale_date), memo]],
                value_input_option="USER_ENTERED"
            )

    return {
        "mercari_fee": mercari_fee,
        "profit": profit,
        "profit_rate": profit_rate_text,
        "sales_row": sales_next_row,
        "inventory_rows": [item["row_number"] for item in selected_items],
        "purchase_price_total": purchase_price_total
    }

def get_line_channel_access_token():
    token = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
    try:
        if not token:
            token = st.secrets.get("LINE_CHANNEL_ACCESS_TOKEN", None)
    except Exception:
        pass
    return token

def get_line_judge_state(line_user_id):
    worksheet = connect_line_judge_sheet()
    rows = worksheet.get_all_values()
    for row_number, row in enumerate(rows[1:], start=2):
        values = row + [""] * max(0, len(LINE_JUDGE_HEADERS) - len(row))
        if values[0] == line_user_id:
            return {
                "row_number": row_number,
                "line_user_id": values[0],
                "status": values[1],
                "image_url": values[2],
                "features": values[3],
                "purchase_price": values[4],
                "judge_result": values[5],
                "updated_at": values[6]
            }
    return None

def upsert_line_judge_state(line_user_id, status, image_url="", features="", purchase_price="", judge_result=""):
    worksheet = connect_line_judge_sheet()
    row = [
        line_user_id,
        status,
        image_url,
        features,
        purchase_price,
        judge_result,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ]
    state = get_line_judge_state(line_user_id)
    if state:
        worksheet.update(
            range_name=f"A{state['row_number']}:G{state['row_number']}",
            values=[row],
            value_input_option="USER_ENTERED"
        )
    else:
        all_values = worksheet.get_all_values()
        next_row = len(all_values) + 1
        worksheet.update(
            range_name=f"A{next_row}:G{next_row}",
            values=[row],
            value_input_option="USER_ENTERED"
        )
    return row

def fetch_line_message_content(message_id):
    token = get_line_channel_access_token()
    if not token:
        raise Exception("LINE_CHANNEL_ACCESS_TOKEN が設定されていません。")

    response = requests.get(
        f"https://api-data.line.me/v2/bot/message/{message_id}/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=60
    )
    if response.status_code != 200:
        raise Exception(f"LINE画像取得エラー status={response.status_code}: {response.text[:500]}")
    return response.content, response.headers.get("Content-Type", "image/jpeg")

def save_line_judge_image_to_drive(image_bytes, mime_type, line_user_id):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    extension = "jpg"
    if "png" in mime_type:
        extension = "png"
    elif "heic" in mime_type or "heif" in mime_type:
        extension = "heic"

    file_name = f"line_judge_{line_user_id}_{timestamp}.{extension}"
    payload = {
        "token": PHOTO_SAVE_TOKEN,
        "fileName": file_name,
        "mimeType": mime_type,
        "base64Data": base64.b64encode(image_bytes).decode("utf-8")
    }

    response = requests.post(APPS_SCRIPT_PHOTO_URL, json=payload, timeout=60)
    if response.status_code != 200:
        raise Exception(f"Apps Scriptエラー status={response.status_code}: {response.text[:500]}")

    try:
        result = response.json()
    except Exception:
        raise Exception(f"Apps ScriptからJSON以外が返ってきました: {response.text[:500]}")

    if not result.get("success"):
        raise Exception(result.get("error", "写真保存に失敗しました"))

    return file_name, result.get("fileUrl", "")

def analyze_line_judge_photo(image_bytes, mime_type):
    if not client:
        raise Exception("OPENAI_API_KEY が設定されていません。")

    image_base64 = base64.b64encode(image_bytes).decode("utf-8")
    prompt = """
あなたは海外買付の商品を、メルカリ販売目線で見るAIバイヤーです。
写真の商品について、以下を日本語で短く整理してください。

- 商品ジャンル
- 見た目の特徴
- 日本未発売感
- 写真映え
- 発送しやすさ
- メルカリ向けキーワード
- 注意点
- 追加で確認すべきこと

まだ仕入価格は分からないため、価格や利益は断定しないでください。
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": "あなたは海外買付とメルカリ販売の実務に詳しいAI社員です。"},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime_type};base64,{image_base64}"}
                    }
                ]
            }
        ],
        temperature=0.4
    )
    return response.choices[0].message.content

def parse_purchase_price_text(price_text):
    text = str(price_text or "").strip().lower()
    amount_match = re.search(r"([\d,]+(?:\.\d+)?)", text)
    if not amount_match:
        return {
            "original_text": price_text,
            "amount": 0,
            "currency": "",
            "yen": 0
        }

    amount = float(amount_match.group(1).replace(",", ""))
    is_thb = any(keyword in text for keyword in ["バーツ", "baht", "thb", "฿"])
    currency = "THB" if is_thb else "JPY"
    yen = int(amount * 4.5) if is_thb else int(amount)
    return {
        "original_text": price_text,
        "amount": amount,
        "currency": currency,
        "yen": yen
    }

LINE_PURCHASE_JUDGE_START_TEXTS = [
    "買付ジャッジ",
    "買付",
    "仕入れ判断",
    "仕入判断",
    "ジャッジ",
]

LINE_PURCHASE_JUDGE_START_REPLY = """買付ジャッジを開始します。

気になる商品の写真、または
商品名と仕入価格を送ってください。

先にメルカリ販売可否と規約リスクを確認します。
食品・健康食品・サプリ・ハーブパウダー・医薬品・医薬部外品・化粧品など、削除リスクが高い商品は利益に関係なく買付不可になります。"""

def normalize_line_text(text):
    return re.sub(r"\s+", " ", str(text or "").replace("　", " ")).strip()

def is_line_purchase_judge_start_text(text):
    normalized_text = normalize_line_text(text)
    return (
        normalized_text in LINE_PURCHASE_JUDGE_START_TEXTS
        or "買付ジャッジ" in normalized_text
    )

def sanitize_line_purchase_reply(reply_text):
    hidden_prefixes = [
        "販売可否確認",
        "規約リスク",
        "スコア",
        "削除リスク高カテゴリ該当",
        "削除リスクカテゴリ",
        "食品・健康系該当",
        "医薬品・医薬部外品・化粧品該当",
        "効能効果表現リスク",
        "海外輸入時の表示リスク",
        "利益率判定",
        "スコア加点",
    ]
    visible_lines = []
    for line in str(reply_text or "").splitlines():
        stripped_line = line.strip()
        if any(stripped_line.startswith(prefix) for prefix in hidden_prefixes):
            continue
        visible_lines.append(line)
    return "\n".join(visible_lines).strip()

def judge_line_purchase(features, price_info):
    policy_risk = judge_purchase_policy_risk(description=features)
    if policy_risk["policy_risk"] == "高":
        return build_policy_rejected_purchase_message(policy_risk)

    if not client:
        raise Exception("OPENAI_API_KEY が設定されていません。")

    policy_risk_block = format_policy_risk_block(policy_risk)
    prompt = f"""
以下の商品特徴と仕入価格をもとに、海外買付すべきか判定してください。
LINEで読むため、以前の買付アドバイス型の自然で短い買付メモとして返してください。

内部では必ず次の順番で判定してください。
1. 食品・健康系に該当しないか
2. 規約リスクが高くないか
3. 許可カテゴリに該当するか
4. 利益率25％以上か
5. 想定利益500円以上か
6. 最終買付判定

規約リスクが高い商品、削除リスク高カテゴリの商品、食品、飲料、健康食品、サプリメント、
ハーブパウダー、医薬品誤認商品、医薬部外品、化粧品は、利益が出そうでも買付不可にしてください。

規約リスクが低い商品のみ、以下の利益基準で判定してください。
- 利益率25％以上
- 想定利益500円以上

上記2つを満たす場合だけ「買付OK」候補です。
利益500円未満でも、発信ネタ、写真映え、売れ筋検証、少額テストとして意味がある場合は「要確認」または「検証目的なら可」としてください。
ただし削除リスク高カテゴリは例外不可です。

商品名と価格だけで写真情報がない場合でも、衣類・雑貨など通常カテゴリなら仮の想定販売価格と想定利益を必ず出してください。
例: 猫柄タイパンツ 300円なら、想定販売価格は1,200〜1,800円、想定利益は500〜900円前後のように幅で返してください。
写真なしの場合は、買付判定を「要確認」または「条件付きで買い」にしてください。

【商品特徴】
{features}

【仕入価格】
入力: {price_info.get('original_text')}
概算円: {price_info.get('yen')}円

【規約リスクの事前判定】
{policy_risk_block}

100点満点で内部採点して構いません。ただし、スコアやスコア加点の内訳は表示しないでください。

判定:
- 買う: 規約リスクが低く、利益率25％以上、想定利益500円以上を満たす
- 保留: 写真、サイズ、素材、状態、需要、利益の確認余地がある
- 買付不可: 規約リスクが高い、削除リスクが高い、または利益基準を満たさない

LINE返信は必ず以下の形式だけにしてください。

買付判定：買う / 保留 / 買付不可 のどれか
- 商品候補
- 理由
- 想定販売価格
- 想定利益
- 注意点

通常商品のLINE返信には、以下の内部項目を絶対に表示しないでください。
- 販売可否確認
- 規約リスク
- スコア
- 削除リスク高カテゴリ該当
- 削除リスクカテゴリ
- 食品・健康系該当
- 医薬品・医薬部外品・化粧品該当
- 効能効果表現リスク
- 海外輸入時の表示リスク
- スコア加点の内訳

通常商品の返信形式:
買付判定：買う
商品候補：カラフルリゾートパンツ（海外限定デザイン）

理由：
写真映え抜群で日本未発売感が強く、夏のカジュアルファッションとして需要あり。
軽量で発送も楽。
仕入価格1500円は利益確保しやすい水準。

想定販売価格：3500〜4000円
想定利益：約2000円前後（手数料・送料差引後）

注意点：
サイズ・素材・タグの有無を必ず確認。
色落ちや生地の状態もチェック必須。
伸縮性や透け感も重要。

買う場合は「登録」と送ってください。
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": "あなたは海外買付の商品をメルカリで売れるか判定するAIバイヤーです。"},
            {"role": "user", "content": prompt}
        ],
        temperature=0.5
    )
    return sanitize_line_purchase_reply(response.choices[0].message.content)

def reply_line_text(reply_token, text):
    token = get_line_channel_access_token()
    if not token:
        raise Exception("LINE_CHANNEL_ACCESS_TOKEN が設定されていません。")

    response = requests.post(
        "https://api.line.me/v2/bot/message/reply",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json={
            "replyToken": reply_token,
            "messages": [{"type": "text", "text": text[:4900]}]
        },
        timeout=30
    )
    if response.status_code not in (200, 201):
        raise Exception(f"LINE返信エラー status={response.status_code}: {response.text[:500]}")

def handle_line_purchase_judge_event(event, should_reply=True):
    source = event.get("source", {})
    message = event.get("message", {})
    line_user_id = source.get("userId", "")
    reply_token = event.get("replyToken", "")
    message_type = message.get("type")
    reply_text = ""

    if not line_user_id:
        return "LINEユーザーIDが取得できませんでした。"

    if message_type == "text":
        raw_text = message.get("text", "")
        text = normalize_line_text(raw_text)
        print(f"[LINE_TEXT] received_text={text}", flush=True)

        if is_line_purchase_judge_start_text(text):
            print("[LINE_ROUTE] purchase_judge_start matched", flush=True)
            upsert_line_judge_state(line_user_id, "写真待ち")
            reply_text = LINE_PURCHASE_JUDGE_START_REPLY
        else:
            state = get_line_judge_state(line_user_id)
            print(f"[LINE_ROUTE] current_status={state.get('status') if state else 'none'}", flush=True)

            if state and state.get("status") == "写真待ち":
                print("[LINE_ROUTE] purchase_judge_photo_wait_text", flush=True)
                price_info = parse_purchase_price_text(text)
                if price_info["yen"] <= 0:
                    reply_text = "気になる商品の写真を送ってください。商品名と価格をテキストで送る場合は、例：猫柄タイパンツ 300円 のように送ってください。"
                else:
                    judge_result = judge_line_purchase(text, price_info)
                    upsert_line_judge_state(
                        line_user_id,
                        "登録確認待ち",
                        image_url=state.get("image_url", ""),
                        features=text,
                        purchase_price=f"{price_info['original_text']}（概算 {price_info['yen']}円）",
                        judge_result=judge_result
                    )
                    reply_text = judge_result
            elif state and state.get("status") == "価格待ち":
                print("[LINE_ROUTE] purchase_judge_price_wait_text", flush=True)
                price_info = parse_purchase_price_text(text)
                if price_info["yen"] <= 0:
                    reply_text = "仕入価格を読み取れませんでした。例：250バーツ、1200円 のように送ってください。"
                else:
                    judge_result = judge_line_purchase(state.get("features", ""), price_info)
                    upsert_line_judge_state(
                        line_user_id,
                        "登録確認待ち",
                        image_url=state.get("image_url", ""),
                        features=state.get("features", ""),
                        purchase_price=f"{price_info['original_text']}（概算 {price_info['yen']}円）",
                        judge_result=judge_result
                    )
                    reply_text = judge_result
            else:
                print("[LINE_ROUTE] purchase_judge_start_not_matched", flush=True)
                reply_text = "買付ジャッジを始める場合は「買付ジャッジ」と送ってください。"

    elif message_type == "image":
        state = get_line_judge_state(line_user_id)
        if not state or state.get("status") != "写真待ち":
            reply_text = "先に「買付ジャッジ」と送ってから、商品の写真を送ってください。"
        else:
            image_bytes, mime_type = fetch_line_message_content(message.get("id", ""))
            _, image_url = save_line_judge_image_to_drive(image_bytes, mime_type, line_user_id)
            features = analyze_line_judge_photo(image_bytes, mime_type)
            upsert_line_judge_state(
                line_user_id,
                "価格待ち",
                image_url=image_url,
                features=features
            )
            reply_text = (
                "写真を確認しました。\n\n"
                f"{features}\n\n"
                "仕入価格を教えてください。例：250バーツ、1200円"
            )
    else:
        reply_text = "買付ジャッジでは、テキストまたは写真を送ってください。"

    if should_reply and reply_token:
        reply_line_text(reply_token, reply_text)

    return reply_text

def process_line_purchase_judge_webhook(payload, should_reply=True):
    """LINE webhook payload内の買付ジャッジ会話を処理する。外部Webhookから呼び出す想定。"""
    replies = []
    for event in payload.get("events", []):
        replies.append(handle_line_purchase_judge_event(event, should_reply=should_reply))
    return replies

def get_recent_purchase_rows(limit=3):
    """買付登録シートの直近データを表示用に整える。失敗時は画面を止めず仮データにする。"""
    fallback_rows = [
        {
            "登録日時": "2024/05/28 14:23",
            "商品名": "ハワイアンパンツ",
            "購入価格": "¥1,500",
            "色": "ホワイト",
            "サイズ": "フリー",
            "数量": "1",
            "写真": "登録済み",
            "保存先": "Google Drive"
        }
    ]

    try:
        worksheet = connect_purchase_sheet()
        records = worksheet.get_all_records()
        if not records:
            return fallback_rows

        rows = []
        for record in records[-limit:][::-1]:
            purchase_price = record.get("購入価格", "")
            if isinstance(purchase_price, (int, float)):
                purchase_price = f"¥{purchase_price:,.0f}"

            folder_url = str(record.get("商品写真フォルダ", "") or "")
            rows.append(
                {
                    "登録日時": record.get("登録日時", ""),
                    "商品名": record.get("商品名", ""),
                    "購入価格": purchase_price,
                    "色": record.get("色", ""),
                    "サイズ": record.get("サイズ", ""),
                    "数量": record.get("数量", ""),
                    "写真": "登録済み" if record.get("商品写真ファイル名一覧", "") else "なし",
                    "保存先": folder_url if folder_url else "未登録"
                }
            )
        return rows
    except Exception:
        return fallback_rows

PURCHASE_FORM_FIELD_KEYS = [
    "buyer_name",
    "product_name",
    "brand",
    "price",
    "quantity",
    "color",
    "size",
    "place",
    "date",
    "category",
    "memo"
]

def get_purchase_form_key(field_name):
    version = st.session_state.get("purchase_form_version", 0)
    return f"purchase_form_{version}_{field_name}"

def cleanup_old_purchase_form_keys():
    """リセット後に、古いバージョンの入力値と写真アップロード値を消す。"""
    current_form_prefix = f"purchase_form_{st.session_state.get('purchase_form_version', 0)}_"
    current_photo_key = f"purchase_photos_{st.session_state.get('purchase_photo_uploader_version', 0)}"

    for key in list(st.session_state.keys()):
        key_text = str(key)
        if key_text.startswith("purchase_form_") and key_text != "purchase_form_version":
            if not key_text.startswith(current_form_prefix):
                del st.session_state[key]
        elif key_text.startswith("purchase_photos_") and key_text != current_photo_key:
            del st.session_state[key]

def reset_purchase_form(clear_save_result=True):
    """買付登録フォームと写真アップロード欄を次の空フォームに切り替える。"""
    st.session_state["purchase_form_version"] = (
        st.session_state.get("purchase_form_version", 0) + 1
    )
    st.session_state["purchase_photo_uploader_version"] = (
        st.session_state.get("purchase_photo_uploader_version", 0) + 1
    )
    st.session_state["purchase_form_needs_cleanup"] = True

    if clear_save_result and "purchase_save_result" in st.session_state:
        del st.session_state["purchase_save_result"]

def render_purchase_table(rows):
    table_rows = []
    for row in rows:
        save_to = row.get("保存先", "")
        if str(save_to).startswith("http"):
            save_to_html = f'<a href="{save_to}" target="_blank">Google Drive</a>'
        else:
            save_to_html = save_to

        table_rows.append(
            "<tr>"
            f"<td>{row.get('登録日時', '')}</td>"
            f"<td>{row.get('商品名', '')}</td>"
            f"<td>{row.get('購入価格', '')}</td>"
            f"<td>{row.get('色', '')}</td>"
            f"<td>{row.get('サイズ', '')}</td>"
            f"<td>{row.get('数量', '')}</td>"
            f"<td>{row.get('写真', '')}</td>"
            f"<td>{save_to_html}</td>"
            "</tr>"
        )

    st.markdown(
        """
        <div class="recent-table-card">
            <div class="section-title">📋 最近の買付登録</div>
            <div class="table-scroll">
                <table class="recent-table">
                    <thead>
                        <tr>
                            <th>登録日時</th>
                            <th>商品名</th>
                            <th>購入価格</th>
                            <th>色</th>
                            <th>サイズ</th>
                            <th>数量</th>
                            <th>写真</th>
                            <th>保存先</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows}
                    </tbody>
                </table>
            </div>
        </div>
        """.format(rows="".join(table_rows)),
        unsafe_allow_html=True
    )

if os.getenv("MERCARI_WEBHOOK_IMPORT") != "1":
    st.set_page_config(
        page_title="メルカリちゃん",
        page_icon="🛍️",
        layout="wide"
    )

    st.markdown(
        """
        <style>
        :root {
            --mercari-red: #f5222d;
            --mercari-pink: #fff1f1;
            --mercari-border: #e8edf3;
            --mercari-text: #202734;
            --mercari-muted: #667085;
            --mercari-green: #12a673;
            --mercari-yellow: #f0aa00;
        }

        .block-container {
            padding-top: 1.1rem;
            padding-bottom: 2.2rem;
            max-width: 1540px;
        }

        [data-testid="stApp"],
        [data-testid="stAppViewContainer"] {
            background: #ffffff;
            color: var(--mercari-text);
        }

        [data-testid="stMarkdownContainer"],
        [data-testid="stWidgetLabel"],
        [data-testid="stWidgetLabel"] label,
        label,
        p {
            color: var(--mercari-text);
        }

        [data-testid="stSidebar"] {
            background: #fbfdff;
            border-right: 1px solid var(--mercari-border);
        }

        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
            margin: 0;
        }

        .brand-title {
            font-size: 2.2rem;
            line-height: 1.1;
            font-weight: 900;
            color: var(--mercari-text);
        }

        .brand-title span {
            color: var(--mercari-red);
        }

        .brand-copy {
            margin-top: 8px;
            color: #344054;
            font-weight: 600;
        }

        .sidebar-menu {
            padding-top: 12px;
        }

        .sidebar-item {
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 12px 14px;
            margin: 6px 0;
            border-radius: 10px;
            color: #344054;
            font-weight: 700;
        }

        .sidebar-item.active {
            background: linear-gradient(90deg, #ffe6e6, #fff4f4);
            color: var(--mercari-red);
        }

        .line-card {
            margin-top: 26px;
            padding: 14px;
            border: 1px solid #ccebd9;
            background: #f3fff8;
            border-radius: 10px;
            color: #079455;
            font-weight: 800;
            line-height: 1.5;
        }

        .page-title-row {
            display: flex;
            justify-content: space-between;
            gap: 16px;
            align-items: center;
            margin-bottom: 14px;
        }

        .page-title {
            color: var(--mercari-red);
            font-size: 2rem;
            font-weight: 900;
            line-height: 1.2;
        }

        .page-description {
            color: #344054;
            font-weight: 600;
            margin-left: 10px;
        }

        .section-title {
            font-size: 1.2rem;
            font-weight: 900;
            color: #202734;
            margin-bottom: 12px;
        }

        .pill {
            display: inline-block;
            background: #fff1f1;
            color: var(--mercari-red);
            padding: 4px 10px;
            border-radius: 999px;
            font-weight: 800;
            font-size: .82rem;
            margin-left: 6px;
        }

        .drive-note {
            border: 1px solid #c9defd;
            background: #f5f9ff;
            border-radius: 10px;
            padding: 13px 14px;
            color: #174ea6;
            font-weight: 800;
            margin-top: 12px;
        }

        .drive-note small {
            color: #344054;
            font-weight: 600;
        }

        .recent-table-card {
            border: 1px solid var(--mercari-border);
            border-radius: 12px;
            padding: 16px;
            margin-top: 16px;
            box-shadow: 0 8px 22px rgba(16, 24, 40, .04);
        }

        .table-scroll {
            overflow-x: auto;
        }

        .recent-table {
            width: 100%;
            border-collapse: collapse;
            min-width: 820px;
            font-size: .94rem;
        }

        .recent-table th,
        .recent-table td {
            border-top: 1px solid var(--mercari-border);
            padding: 12px 10px;
            text-align: left;
            white-space: nowrap;
        }

        .recent-table th {
            color: #202734;
            background: #fafafa;
            font-weight: 900;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 13px;
            border-color: var(--mercari-border);
            box-shadow: 0 8px 22px rgba(16, 24, 40, .04);
        }

        .stTextInput input,
        .stNumberInput input,
        .stDateInput input,
        .stTextArea textarea {
            background: #ffffff !important;
            border: 1px solid #d0d5dd !important;
            border-radius: 9px;
            color: var(--mercari-text) !important;
            min-height: 44px;
            font-size: 1rem;
        }

        .stTextInput input::placeholder,
        .stTextArea textarea::placeholder {
            color: #98a2b3 !important;
            opacity: 1 !important;
        }

        .stButton button {
            border-radius: 9px;
            min-height: 46px;
            font-weight: 900;
        }

        .stButton button[kind="primary"] {
            background: var(--mercari-red) !important;
            border-color: var(--mercari-red) !important;
            color: #ffffff !important;
        }

        [data-testid="stFileUploaderDropzone"] {
            border: 2px dashed #c9cdd4;
            border-radius: 13px;
            padding: 24px 14px;
            background: #ffffff !important;
            color: #344054 !important;
        }

        [data-testid="stFileUploaderDropzone"] * {
            color: #344054 !important;
        }

        [data-testid="stFileUploaderDropzone"] button {
            border-color: var(--mercari-red);
            color: var(--mercari-red) !important;
            font-weight: 900;
        }

        @media (max-width: 820px) {
            .block-container {
                padding-left: 1rem;
                padding-right: 1rem;
            }

            .page-title-row {
                align-items: flex-start;
                flex-direction: column;
            }

            .brand-title {
                font-size: 1.8rem;
            }

            .page-title {
                font-size: 1.7rem;
            }

            .page-description {
                display: block;
                margin-left: 0;
                margin-top: 6px;
            }
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.sidebar.markdown(
        """
        <div class="sidebar-menu">
            <div class="sidebar-item">🏠 ホーム</div>
            <div class="sidebar-item active">🛒 買付登録</div>
            <div class="sidebar-item">✨ AI出品登録</div>
            <div class="sidebar-item">🧰 在庫一覧</div>
            <div class="sidebar-item">🏷️ 売上登録・在庫更新</div>
            <div class="sidebar-item">📈 売上管理</div>
            <div class="sidebar-item">🖼️ 写真加工</div>
            <div class="sidebar-item">📋 データ一覧</div>
            <div class="sidebar-item">⚙️ 設定</div>
            <div class="sidebar-item">📖 使い方ガイド</div>
            <div class="line-card">LINE公式アカウントからも<br>登録できます！</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="app-header-text">
            <div class="brand-title">🛍️ <span>メルカリ</span>ちゃん</div>
            <div class="brand-copy">買付から出品・売上管理まで、あなたの物販をサポート！</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    tab4, tab1, tab6, tab5, tab3, tab2 = st.tabs([
        "買付・AI出品登録",
        "在庫登録",
        "売上登録",
        "使い方ガイド",
        "買付登録",
        "AI出品サポート"
    ])

    with tab1:
        st.subheader("商品登録")

        if st.session_state.pop("reset_inventory_registration_form", False):
            reset_inventory_registration_form()

        registration_success_message = st.session_state.pop("inventory_registration_success_message", "")
        if registration_success_message:
            st.success(registration_success_message)
            st.info(f"保存先：{WORKSHEET_NAME}")

        product_name = st.text_input("商品名", key="reg_product_name")
        category = st.text_input("カテゴリ", key="reg_category")
        purchase_price = st.number_input("仕入れ価格", min_value=0, step=100, key="reg_purchase_price")
        purchase_date = st.date_input("仕入日", value=date.today(), key="reg_purchase_date")
        planned_price = st.number_input("予定販売価格", min_value=0, step=100, key="reg_planned_price")
        stock_quantity = st.number_input(
            "在庫数",
            min_value=1,
            step=1,
            value=1,
            key="reg_stock_quantity"
        )
        shipping_fee = st.number_input("送料", min_value=0, step=100, key="reg_shipping_fee")
        packing_fee = st.number_input("梱包資材費", min_value=0, step=10, key="reg_packing_fee")
        memo = st.text_area("メモ", key="reg_memo")

        if st.button("登録する"):
            if product_name == "":
                st.error("商品名を入力してください。")
            else:
                try:
                    worksheet = connect_sheet()

                    product_id = generate_product_id()
                    registered_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    row = [
                        product_id,
                        registered_date,
                        str(purchase_date),
                        product_name,
                        category,
                        purchase_price,
                        planned_price,
                        "",
                        "",
                        shipping_fee,
                        packing_fee,
                        "",
                        "",
                        "在庫あり",
                        "",
                        memo,
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        stock_quantity
                    ]

                    worksheet.append_row(row)

                    st.session_state["inventory_registration_success_message"] = (
                        f"スプレッドシートに商品を登録しました！ 在庫数：{stock_quantity}"
                    )
                    st.session_state["reset_inventory_registration_form"] = True
                    st.rerun()

                except FileNotFoundError:
                    st.error(f"Google認証ファイルが見つかりません。参照先：{SERVICE_ACCOUNT_FILE}")
                except Exception as e:
                    st.error("エラーが発生しました。")
                    st.write(e)

    with tab2:
        st.subheader("AI出品サポート")
        st.write("写真を見ながら、40文字以内タイトルと5分で売れる商品説明文を作ります。")

        if not OPENAI_API_KEY:
            st.warning(".env に OPENAI_API_KEY が設定されていません。")
        else:
            uploaded_photo = st.file_uploader(
                "商品写真をアップロード",
                type=["jpg", "jpeg", "png"],
                key="ai_uploaded_photo"
            )

            if uploaded_photo is not None:
                st.image(uploaded_photo, caption="アップロードされた商品写真", width=350)

            ai_product_name = st.text_input("商品名", key="ai_product_name")
            ai_category = st.text_input("カテゴリ", key="ai_category")
            ai_brand = st.text_input("ブランド", key="ai_brand")
            ai_color = st.text_input("カラー", key="ai_color")
            ai_size = st.text_input("サイズ", key="ai_size")
            ai_condition = st.selectbox(
                "商品の状態",
                ["新品・未使用", "未使用に近い", "目立った傷や汚れなし", "やや傷や汚れあり", "傷や汚れあり"],
                key="ai_condition"
            )
            ai_features = st.text_area("商品の特徴", key="ai_features")
            ai_target = st.text_input("ターゲット（例：40代女性、学生、旅行好き など）", key="ai_target")
            ai_planned_price = st.number_input("販売予定価格", min_value=0, step=100, key="ai_planned_price")
            ai_memo = st.text_area("補足メモ", key="ai_memo")

            if st.button("写真から特徴を読み取る"):
                if uploaded_photo is None:
                    st.error("先に商品写真をアップロードしてください。")
                else:
                    try:
                        photo_text = analyze_photo(
                            uploaded_photo,
                            ai_product_name,
                            ai_category
                        )
                        st.session_state["photo_analysis"] = photo_text
                        st.success("写真から特徴を読み取りました！")
                    except Exception as e:
                        st.error("写真の読み取りでエラーが出ました。")
                        st.write(e)

            ai_photo_features = st.text_area(
                "写真から見える特徴",
                value=st.session_state.get("photo_analysis", ""),
                placeholder="写真を読み取ると、ここに特徴が入ります。自分で追記・修正もできます。",
                height=180,
                key="ai_photo_features"
            )

            if st.button("商品説明を作る"):
                if ai_product_name == "":
                    st.error("商品名を入力してください。")
                else:
                    try:
                        result = generate_description(
                            ai_product_name,
                            ai_category,
                            ai_brand,
                            ai_color,
                            ai_size,
                            ai_condition,
                            ai_features,
                            ai_photo_features,
                            ai_target,
                            ai_memo,
                            ai_planned_price
                        )

                        description_result = {
                            "product_name": ai_product_name,
                            "category": ai_category,
                            "brand": ai_brand,
                            "color": ai_color,
                            "size": ai_size,
                            "condition": ai_condition,
                            "features": ai_features,
                            "photo_features": ai_photo_features,
                            "target": ai_target,
                            "planned_price": ai_planned_price,
                            "memo": ai_memo,
                            "title": limit_title(result.get("title", "")),
                            "description": result.get("description", ""),
                            "bullet_points": result.get("bullet_points", []),
                            "hashtags": result.get("hashtags", []),
                            "quick_sale_price": result.get("quick_sale_price", ""),
                            "standard_price": result.get("standard_price", ""),
                            "premium_price": result.get("premium_price", ""),
                            "recommended_price": result.get("recommended_price", ""),
                            "price_reason": result.get("price_reason", "")
                        }

                        saved_file_path = save_description_to_text_file(description_result)
                        description_result["text_file_path"] = saved_file_path
                        st.session_state["description_result"] = description_result

                        st.success("商品説明文を作成しました！")
                        st.info(f"txtファイルにも自動保存しました：{saved_file_path}")

                    except Exception as e:
                        st.error("商品説明の生成でエラーが出ました。")
                        st.write(e)

            if "description_result" in st.session_state:
                saved = st.session_state["description_result"]

                st.write("### タイトル")
                st.write(saved.get("title", ""))
                st.caption(f"文字数：{len(saved.get('title', ''))}文字 / 40文字以内推奨")

                st.write("### 商品説明")
                st.text_area(
                    "コピペ用 商品説明",
                    value=saved.get("description", ""),
                    height=300
                )

                st.write("### おすすめポイント")
                for item in saved.get("bullet_points", []):
                    st.write(f"- {item}")

                st.write("### ハッシュタグ")
                st.text_area(
                    "コピペ用 ハッシュタグ",
                    value=" ".join(saved.get("hashtags", [])),
                    height=80
                )

                st.write("### 価格提案")
                st.write("早く売る価格：", saved.get("quick_sale_price", ""))
                st.write("標準価格：", saved.get("standard_price", ""))
                st.write("高めに売る価格：", saved.get("premium_price", ""))
                st.write("おすすめ販売価格：", saved.get("recommended_price", ""))
                st.write("価格理由：", saved.get("price_reason", ""))

                if st.button("商品説明をスプレッドシートに保存する"):
                    try:
                        save_description_to_sheet(saved)
                        st.success("商品説明ログに保存しました！")
                        st.info(f"保存先：{DESCRIPTION_SHEET_NAME}")
                    except Exception as e:
                        st.error("保存でエラーが出ました。")
                        st.write(e)

    with tab3:
        st.markdown(
            """
            <div class="page-title-row">
                <div>
                    <span class="page-title">🛒 買付登録</span>
                    <span class="page-description">仕入れた商品の情報と写真を登録します</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        purchase_save_result = st.session_state.pop("purchase_save_result", None)
        if purchase_save_result:
            st.success("買付情報をスプレッドシートに保存しました！")
            st.info(f"保存先：{purchase_save_result.get('sheet_name', PURCHASE_SHEET_NAME)}")
            saved_photo_urls = purchase_save_result.get("photo_urls", [])
            if saved_photo_urls:
                st.markdown("写真保存先： " + " / ".join(f"[Google Drive]({url})" for url in saved_photo_urls))
            else:
                st.write("写真保存先：写真なし")

        if "purchase_form_version" not in st.session_state:
            st.session_state["purchase_form_version"] = 0
        if "purchase_photo_uploader_version" not in st.session_state:
            st.session_state["purchase_photo_uploader_version"] = 0
        if st.session_state.pop("purchase_form_needs_cleanup", False):
            cleanup_old_purchase_form_keys()

        purchase_buyer_name_key = get_purchase_form_key("buyer_name")
        purchase_product_name_key = get_purchase_form_key("product_name")
        purchase_brand_key = get_purchase_form_key("brand")
        purchase_price_key = get_purchase_form_key("price")
        purchase_quantity_key = get_purchase_form_key("quantity")
        purchase_color_key = get_purchase_form_key("color")
        purchase_size_key = get_purchase_form_key("size")
        purchase_place_key = get_purchase_form_key("place")
        purchase_date_key = get_purchase_form_key("date")
        purchase_category_key = get_purchase_form_key("category")
        purchase_memo_key = get_purchase_form_key("memo")

        form_column, photo_column = st.columns([1, 1], gap="large")

        with form_column:
            with st.container(border=True):
                st.markdown('<div class="section-title">🧾 商品情報を入力</div>', unsafe_allow_html=True)

                buyer_name = st.text_input(
                    "登録者名",
                    placeholder="例）山田",
                    key=purchase_buyer_name_key
                )

                first_row_left, first_row_right = st.columns(2)
                with first_row_left:
                    purchase_product_name = st.text_input(
                        "商品名（必須）",
                        placeholder="例）花柄ワンピース",
                        key=purchase_product_name_key
                    )
                with first_row_right:
                    purchase_brand = st.text_input(
                        "ブランド",
                        placeholder="例）ZARA",
                        key=purchase_brand_key
                    )

                second_row_left, second_row_right = st.columns(2)
                with second_row_left:
                    purchase_price = st.number_input(
                        "購入価格（必須）",
                        min_value=0,
                        step=100,
                        format="%d",
                        key=purchase_price_key,
                        help="円単位で入力してください"
                    )
                with second_row_right:
                    purchase_quantity = st.number_input(
                        "数量（必須）",
                        min_value=1,
                        step=1,
                        value=1,
                        key=purchase_quantity_key
                    )

                third_row_left, third_row_right = st.columns(2)
                with third_row_left:
                    purchase_color = st.text_input(
                        "色",
                        placeholder="例）ネイビー",
                        key=purchase_color_key
                    )
                with third_row_right:
                    purchase_size = st.text_input(
                        "サイズ",
                        placeholder="例）M",
                        key=purchase_size_key
                    )

                fourth_row_left, fourth_row_right = st.columns(2)
                with fourth_row_left:
                    purchase_place = st.text_input(
                        "仕入れ先",
                        placeholder="例）ハワイ・アラモアナセンター",
                        key=purchase_place_key
                    )
                with fourth_row_right:
                    purchase_date = st.date_input(
                        "仕入れ日",
                        value=date.today(),
                        key=purchase_date_key
                    )

                purchase_category = st.text_input(
                    "カテゴリ",
                    placeholder="例）レディース ワンピース",
                    key=purchase_category_key
                )
                purchase_memo = st.text_area(
                    "メモ（任意）",
                    placeholder="商品の状態・特徴・気づいたことなどをメモできます",
                    key=purchase_memo_key,
                    height=110
                )

                reset_column, save_column = st.columns([1, 2])
                with reset_column:
                    st.button("リセット", use_container_width=True, on_click=reset_purchase_form)
                with save_column:
                    save_purchase_button = st.button(
                        "買付情報を保存する",
                        type="primary",
                        use_container_width=True
                    )

        with photo_column:
            with st.container(border=True):
                st.markdown(
                    '<div class="section-title">🖼️ 写真を登録 <span class="pill">最大10枚まで</span></div>',
                    unsafe_allow_html=True
                )
                purchase_photo_uploader_key = (
                    f"purchase_photos_{st.session_state['purchase_photo_uploader_version']}"
                )
                purchase_photos = st.file_uploader(
                    "ここに写真をドラッグ＆ドロップ、または写真を選択",
                    type=["jpg", "jpeg", "png", "heic"],
                    accept_multiple_files=True,
                    key=purchase_photo_uploader_key,
                    help="JPG / PNG / HEIC に対応しています"
                )

                photo_count = len(purchase_photos) if purchase_photos else 0
                st.caption("対応形式：JPG / PNG / HEIC")

                if photo_count > 10:
                    st.error("写真は最大10枚まで登録できます。10枚以下にしてください。")
                elif photo_count > 0:
                    st.caption(f"登録中の写真（{photo_count}枚）")
                    preview_columns = st.columns(min(photo_count, 5))
                    for index, uploaded_photo in enumerate(purchase_photos[:10]):
                        with preview_columns[index % len(preview_columns)]:
                            if uploaded_photo.type in ["image/jpeg", "image/png"]:
                                st.image(uploaded_photo, caption=f"{index + 1}枚目", use_container_width=True)
                            else:
                                st.caption(f"{index + 1}枚目：{uploaded_photo.name}")
                else:
                    st.info("スマホでは写真を選択して、そのまま買付情報と一緒に保存できます。")

                st.markdown(
                    """
                    <div class="drive-note">
                        写真は自動でGoogle Driveに保存されます<br>
                        <small>保存先フォルダ：メルカリちゃん買付写真</small>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

        if save_purchase_button:
            photo_count = len(purchase_photos) if purchase_photos else 0

            if purchase_product_name == "":
                st.error("商品名を入力してください。")
            elif buyer_name == "":
                st.error("登録者名を入力してください。")
            elif purchase_price <= 0:
                st.error("購入価格を入力してください。")
            elif photo_count > 10:
                st.error("写真は最大10枚まで登録できます。")
            else:
                try:
                    saved_photo_names = []
                    saved_photo_urls = []

                    if purchase_photos:
                        with st.spinner("写真をGoogle Driveへ保存しています。"):
                            for uploaded_photo in purchase_photos:
                                photo_file_name, photo_file_path = save_uploaded_purchase_photo(uploaded_photo)
                                if photo_file_name:
                                    saved_photo_names.append(photo_file_name)
                                if photo_file_path:
                                    saved_photo_urls.append(photo_file_path)

                    memo_lines = []
                    if purchase_brand:
                        memo_lines.append(f"ブランド: {purchase_brand}")
                    memo_lines.append(f"仕入れ日: {purchase_date}")
                    if purchase_memo:
                        memo_lines.append(purchase_memo)

                    data = {
                        "buyer_name": buyer_name,
                        "product_name": purchase_product_name,
                        "purchase_price": purchase_price,
                        "color": purchase_color,
                        "size": purchase_size,
                        "quantity": purchase_quantity,
                        "place": purchase_place,
                        "category": purchase_category,
                        "memo": "\n".join(memo_lines),
                        "photo_file_name": ", ".join(saved_photo_names),
                        "photo_file_path": ", ".join(saved_photo_urls)
                    }

                    with st.spinner("買付情報をスプレッドシートへ保存しています。"):
                        save_purchase_to_sheet(data)

                    st.session_state["purchase_save_result"] = {
                        "sheet_name": PURCHASE_SHEET_NAME,
                        "photo_urls": saved_photo_urls
                    }
                    reset_purchase_form(clear_save_result=False)
                    st.rerun()

                except Exception as e:
                    st.error("買付情報の保存でエラーが出ました。")
                    st.write(e)

        render_purchase_table(get_recent_purchase_rows(limit=3))

    with tab4:
        st.subheader("買付・AI出品登録")
        st.write("商品情報を1回入力して、買付登録・在庫保存・AI商品説明作成までまとめて行います。")

        v3_photos = st.file_uploader(
            "商品写真 3枚以上",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="v3_photos"
        )

        photo_count = len(v3_photos) if v3_photos else 0
        if photo_count > 0:
            st.caption(f"アップロード済み：{photo_count}枚")
            preview_columns = st.columns(min(photo_count, 5))
            for index, uploaded_photo in enumerate(v3_photos[:5]):
                with preview_columns[index % len(preview_columns)]:
                    st.image(uploaded_photo, caption=f"{index + 1}枚目", width=160)

            if photo_count < 3:
                st.error("写真を3枚以上アップロードしてください")
            elif photo_count < 5:
                st.warning("5枚あるとAIがより正確に判断できます")

        st.write("### 基本情報")
        col1, col2 = st.columns(2)
        with col1:
            v3_buyer_name = st.text_input("登録者名", key="v3_buyer_name")
            v3_product_name = st.text_input("商品名", key="v3_product_name")
            v3_category = st.text_input("カテゴリ", key="v3_category")
        with col2:
            v3_brand = st.text_input("ブランド", key="v3_brand")
            v3_color = st.text_input("色 / カラー", key="v3_color")
            v3_size = st.text_input("サイズ", key="v3_size")

        st.write("### 買付情報")
        col1, col2, col3 = st.columns(3)
        with col1:
            v3_purchase_price = st.number_input("購入価格", min_value=0, step=100, key="v3_purchase_price")
        with col2:
            v3_quantity = st.number_input("数量", min_value=1, step=1, value=1, key="v3_quantity")
        with col3:
            v3_purchase_place = st.text_input("購入場所", key="v3_purchase_place")

        st.write("### 販売情報")
        col1, col2, col3 = st.columns(3)
        with col1:
            v3_planned_price = st.number_input("予定販売価格", min_value=0, step=100, key="v3_planned_price")
        with col2:
            v3_shipping_fee = st.number_input("送料", min_value=0, step=100, key="v3_shipping_fee")
        with col3:
            v3_packing_fee = st.number_input("梱包資材費", min_value=0, step=10, key="v3_packing_fee")

        st.write("### 出品情報")
        v3_condition = st.selectbox(
            "商品の状態",
            ["新品・未使用", "未使用に近い", "目立った傷や汚れなし", "やや傷や汚れあり", "傷や汚れあり"],
            key="v3_condition"
        )
        v3_features = st.text_area("商品の特徴", key="v3_features")
        v3_target = st.text_input("ターゲット（例：40代女性、学生、旅行好き など）", key="v3_target")
        v3_memo = st.text_area("メモ / 補足メモ", key="v3_memo")

        if st.button("在庫保存してAI商品説明を作成する"):
            if v3_product_name == "":
                st.error("商品名を入力してください。")
            elif v3_buyer_name == "":
                st.error("登録者名を入力してください。")
            elif photo_count < 3:
                st.error("写真を3枚以上アップロードしてください")
            elif not OPENAI_API_KEY:
                st.error(".env に OPENAI_API_KEY が設定されていません。")
            else:
                if photo_count < 5:
                    st.warning("5枚あるとAIがより正確に判断できます")

                try:
                    with st.spinner("写真をGoogle Driveへ保存しています。"):
                        product_id = generate_product_id()
                        saved_file_names, photo_drive_urls = save_uploaded_purchase_photos(
                            product_id,
                            v3_product_name,
                            v3_photos
                        )

                    purchase_data = {
                        "buyer_name": v3_buyer_name,
                        "product_name": v3_product_name,
                        "purchase_price": v3_purchase_price,
                        "color": v3_color,
                        "size": v3_size,
                        "quantity": v3_quantity,
                        "place": v3_purchase_place,
                        "category": v3_category,
                        "memo": v3_memo,
                        "photo_file_name": ", ".join(saved_file_names),
                        "photo_file_path": photo_drive_urls
                    }

                    with st.spinner("買付登録シートへ保存しています。"):
                        save_purchase_to_sheet(purchase_data)

                    inventory_data = {
                        "product_id": product_id,
                        "purchase_date": str(date.today()),
                        "product_name": v3_product_name,
                        "category": v3_category,
                        "purchase_price": v3_purchase_price,
                        "planned_price": v3_planned_price,
                        "shipping_fee": v3_shipping_fee,
                        "packing_fee": v3_packing_fee,
                        "memo": v3_memo,
                        "description": "",
                        "photo_folder_path": photo_drive_urls
                    }

                    with st.spinner("メルカリちゃん在庫へ保存しています。"):
                        inventory_row_number = save_inventory_to_sheet(inventory_data)

                    with st.spinner("AIが複数写真を分析しています。"):
                        photo_features = analyze_photos(
                            v3_photos,
                            v3_product_name,
                            v3_category
                        )

                    with st.spinner("AI商品説明を生成しています。"):
                        result = generate_description(
                            v3_product_name,
                            v3_category,
                            v3_brand,
                            v3_color,
                            v3_size,
                            v3_condition,
                            v3_features,
                            photo_features,
                            v3_target,
                            v3_memo,
                            v3_planned_price,
                            v3_purchase_price
                        )

                    description_result = {
                        "product_name": v3_product_name,
                        "category": v3_category,
                        "brand": v3_brand,
                        "color": v3_color,
                        "size": v3_size,
                        "condition": v3_condition,
                        "features": v3_features,
                        "photo_features": photo_features,
                        "target": v3_target,
                        "planned_price": v3_planned_price,
                        "memo": v3_memo,
                        "title": limit_title(result.get("title", "")),
                        "description": result.get("description", ""),
                        "bullet_points": result.get("bullet_points", []),
                        "hashtags": result.get("hashtags", []),
                        "quick_sale_price": result.get("quick_sale_price", ""),
                        "standard_price": result.get("standard_price", ""),
                        "premium_price": result.get("premium_price", ""),
                        "recommended_price": result.get("recommended_price", ""),
                        "price_reason": result.get("price_reason", "")
                    }

                    with st.spinner("商品説明ログへ保存しています。"):
                        save_description_to_sheet(description_result)

                    with st.spinner("outputsフォルダへtxt保存しています。"):
                        saved_file_path = save_description_to_text_file(description_result)
                        description_result["text_file_path"] = saved_file_path
                        update_inventory_description(
                            inventory_row_number,
                            description_result.get("description", "")
                        )
                        update_inventory_price_suggestions(
                            inventory_row_number,
                            description_result
                        )

                    st.session_state["v3_description_result"] = description_result
                    st.session_state["v3_saved_product_id"] = product_id
                    st.session_state["v3_photo_drive_urls"] = photo_drive_urls

                    st.success("買付登録・在庫保存・AI商品説明作成が完了しました！")
                    st.info(f"買付登録保存先：{PURCHASE_SHEET_NAME}")
                    st.info(f"在庫保存先：{WORKSHEET_NAME}")
                    st.info(f"商品説明ログ保存先：{DESCRIPTION_SHEET_NAME}")
                    st.info(f"txt保存先：{saved_file_path}")
                    if photo_drive_urls:
                        drive_links = [
                            f"[Google Drive {index + 1}]({url.strip()})"
                            for index, url in enumerate(photo_drive_urls.split(","))
                            if url.strip()
                        ]
                        st.markdown("写真保存先： " + " / ".join(drive_links))

                except Exception as e:
                    st.error("買付・AI出品登録でエラーが出ました。")
                    st.write(e)

        if "v3_description_result" in st.session_state:
            saved = st.session_state["v3_description_result"]

            st.write("### AI生成結果")
            st.write("商品ID：", st.session_state.get("v3_saved_product_id", ""))
            photo_drive_urls = st.session_state.get("v3_photo_drive_urls", "")
            if photo_drive_urls:
                drive_links = [
                    f"[Google Drive {index + 1}]({url.strip()})"
                    for index, url in enumerate(photo_drive_urls.split(","))
                    if url.strip()
                ]
                st.markdown("写真保存先： " + " / ".join(drive_links))
            else:
                st.write("写真保存先：", "")

            st.write("### タイトル")
            st.write(saved.get("title", ""))
            st.caption(f"文字数：{len(saved.get('title', ''))}文字 / 40文字以内推奨")

            st.write("### 商品説明")
            st.text_area(
                "コピペ用 商品説明",
                value=saved.get("description", ""),
                height=300,
                key="v3_description_output"
            )

            st.write("### おすすめポイント")
            for item in saved.get("bullet_points", []):
                st.write(f"- {item}")

            st.write("### ハッシュタグ")
            st.text_area(
                "コピペ用 ハッシュタグ",
                value=" ".join(saved.get("hashtags", [])),
                height=80,
                key="v3_hashtags_output"
            )

            st.write("### 価格提案")
            st.write("早く売る価格：", saved.get("quick_sale_price", ""))
            st.write("標準価格：", saved.get("standard_price", ""))
            st.write("高めに売る価格：", saved.get("premium_price", ""))
            st.write("おすすめ販売価格：", saved.get("recommended_price", ""))
            st.write("価格理由：", saved.get("price_reason", ""))

    with tab6:
        st.subheader("売上登録")
        st.write("在庫ありの商品を選び、売上登録と在庫更新をまとめて行います。")

        try:
            available_items = get_available_inventory_items()
        except FileNotFoundError:
            available_items = []
            st.error(f"Google認証ファイルが見つかりません。参照先：{SERVICE_ACCOUNT_FILE}")
        except Exception as e:
            available_items = []
            st.error("在庫商品の取得でエラーが出ました。")
            st.write(e)

        with st.expander("メルカリ購入メールから登録", expanded=True):
            mercari_email_text = st.text_area(
                "メルカリ購入通知メール本文を貼り付け",
                height=220,
                key="mercari_purchase_email_text"
            )

            if st.button("メール内容を読み取る", key="parse_mercari_purchase_email_button"):
                if not mercari_email_text.strip():
                    if "parsed_mercari_purchase_email" in st.session_state:
                        del st.session_state["parsed_mercari_purchase_email"]
                    st.error("メール本文を貼り付けてください。")
                else:
                    parsed_email = parse_mercari_purchase_email(mercari_email_text)
                    st.session_state["parsed_mercari_purchase_email"] = parsed_email

            parsed_email = st.session_state.get("parsed_mercari_purchase_email")
            if parsed_email:
                missing_items = []
                if not parsed_email.get("buyer_name"):
                    missing_items.append("購入者名")
                if not parsed_email.get("mercari_item_id"):
                    missing_items.append("メルカリ商品ID")
                if not parsed_email.get("item_name"):
                    missing_items.append("商品名")
                if not parsed_email.get("sale_price"):
                    missing_items.append("販売価格")

                st.write("### 抽出結果")
                result_col1, result_col2 = st.columns(2)
                with result_col1:
                    st.write("購入者名：", parsed_email.get("buyer_name") or "未取得")
                    st.write("メルカリ商品ID：", parsed_email.get("mercari_item_id") or "未取得")
                with result_col2:
                    st.write("商品名：", parsed_email.get("item_name") or "未取得")
                    st.write("販売価格：", format_yen(parsed_email.get("sale_price", 0)))

                if missing_items:
                    st.error("抽出できない項目があります：" + "、".join(missing_items))

                is_bundle = is_bundle_item_name(parsed_email.get("item_name", ""))
                if is_bundle:
                    st.warning("まとめ商品の可能性があります。売約済みにする在庫商品を複数選択してください。")

                if not available_items:
                    st.info("登録できる在庫がありません。")
                else:
                    item_label = lambda item: (
                        f"{item['product_id']}｜{item['product_name']}｜"
                        f"仕入 {format_yen(item['purchase_price'])}｜予定 {format_yen(item['planned_price'])}"
                    )

                    if is_bundle:
                        selected_email_items = st.multiselect(
                            "売約済みにする在庫商品",
                            options=available_items,
                            format_func=item_label,
                            key="email_sales_selected_items"
                        )
                    else:
                        selected_email_item = st.selectbox(
                            "売約済みにする在庫商品",
                            options=available_items,
                            format_func=item_label,
                            key="email_sales_selected_item"
                        )
                        selected_email_items = [selected_email_item] if selected_email_item else []

                    email_sale_col1, email_sale_col2, email_sale_col3 = st.columns(3)
                    with email_sale_col1:
                        email_sale_date = st.date_input(
                            "販売日",
                            value=date.today(),
                            key="email_sales_sale_date"
                        )
                    with email_sale_col2:
                        email_sale_price = st.number_input(
                            "販売価格",
                            min_value=0,
                            step=100,
                            value=int(parsed_email.get("sale_price", 0)),
                            key=f"email_sales_sale_price_{parsed_email.get('mercari_item_id', 'unknown')}"
                        )
                    with email_sale_col3:
                        st.text_input("販売先", value="メルカリ", disabled=True, key="email_sales_channel_display")

                    email_cost_col1, email_cost_col2 = st.columns(2)
                    with email_cost_col1:
                        email_shipping_fee = st.number_input(
                            "送料",
                            min_value=0,
                            step=100,
                            key="email_sales_shipping_fee"
                        )
                    with email_cost_col2:
                        email_packing_fee = st.number_input(
                            "梱包資材費",
                            min_value=0,
                            step=10,
                            key="email_sales_packing_fee"
                        )

                    email_sales_memo = st.text_area("メモ", key="email_sales_memo")
                    purchase_price_total = sum(item["purchase_price"] for item in selected_email_items)
                    email_mercari_fee, email_profit, email_profit_rate = calculate_sales_values(
                        email_sale_price,
                        purchase_price_total,
                        email_shipping_fee,
                        email_packing_fee
                    )

                    email_calc_col1, email_calc_col2, email_calc_col3, email_calc_col4 = st.columns(4)
                    with email_calc_col1:
                        st.metric("仕入価格合計", format_yen(purchase_price_total))
                    with email_calc_col2:
                        st.metric("メルカリ手数料", format_yen(email_mercari_fee))
                    with email_calc_col3:
                        st.metric("利益", format_yen(email_profit))
                    with email_calc_col4:
                        st.metric("利益率", format_profit_rate(email_profit_rate))

                    if st.button(
                        "メール内容から売上登録して在庫を売約済みにする",
                        type="primary",
                        use_container_width=True,
                        key="save_email_sales_button"
                    ):
                        if missing_items:
                            st.error("抽出できない項目があるため保存できません。")
                        elif not selected_email_items:
                            st.error("在庫商品を選択してください。")
                        elif email_sale_price <= 0:
                            st.error("販売価格を入力してください。")
                        else:
                            try:
                                result = save_mercari_email_sales_registration(
                                    selected_email_items,
                                    parsed_email,
                                    email_sale_date,
                                    email_sale_price,
                                    email_shipping_fee,
                                    email_packing_fee,
                                    email_sales_memo
                                )
                                st.success("メール内容から売上登録し、在庫を売約済みに更新しました！")
                                st.info(f"売上管理シート 行番号：{result['sales_row']}")
                                st.info(
                                    "在庫シート 行番号："
                                    + ", ".join(str(row_number) for row_number in result["inventory_rows"])
                                    + " を売約済みに更新しました。"
                                )
                                st.write("利益：", format_yen(result["profit"]))
                                st.write("利益率：", result["profit_rate"])
                            except Exception as e:
                                st.error("メール内容からの売上登録でエラーが出ました。")
                                st.write(e)

        if not available_items:
            st.info("登録できる在庫がありません。")
        else:
            selected_item = st.selectbox(
                "売上登録する商品",
                options=available_items,
                format_func=lambda item: (
                    f"{item['product_id']}｜{item['product_name']}｜"
                    f"仕入 {format_yen(item['purchase_price'])}｜予定 {format_yen(item['planned_price'])}"
                )
            )

            st.write("### 売上情報")
            sales_col1, sales_col2, sales_col3 = st.columns(3)
            with sales_col1:
                sale_date = st.date_input("販売日", value=date.today(), key="sales_sale_date")
            with sales_col2:
                sale_price = st.number_input(
                    "販売価格",
                    min_value=0,
                    step=100,
                    value=selected_item["planned_price"],
                    key=f"sales_sale_price_{selected_item['row_number']}"
                )
            with sales_col3:
                sales_channel = st.text_input("販売先", value="メルカリ", key="sales_channel")

            cost_col1, cost_col2 = st.columns(2)
            with cost_col1:
                shipping_fee = st.number_input("送料", min_value=0, step=100, key="sales_shipping_fee")
            with cost_col2:
                packing_fee = st.number_input("梱包資材費", min_value=0, step=10, key="sales_packing_fee")

            sales_memo = st.text_area("メモ", key="sales_memo")

            mercari_fee, profit, profit_rate = calculate_sales_values(
                sale_price,
                selected_item["purchase_price"],
                shipping_fee,
                packing_fee
            )

            result_col1, result_col2, result_col3 = st.columns(3)
            with result_col1:
                st.metric("メルカリ手数料", format_yen(mercari_fee))
            with result_col2:
                st.metric("利益", format_yen(profit))
            with result_col3:
                st.metric("利益率", format_profit_rate(profit_rate))

            if st.button("売上を保存して在庫を更新する", type="primary", use_container_width=True):
                if sale_price <= 0:
                    st.error("販売価格を入力してください。")
                else:
                    try:
                        result = save_sales_registration(
                            selected_item,
                            sale_date,
                            sale_price,
                            shipping_fee,
                            packing_fee,
                            sales_channel,
                            sales_memo
                        )
                        st.success("売上管理への保存と在庫更新が完了しました！")
                        st.info(f"売上管理シート 行番号：{result['sales_row']}")
                        st.info(f"在庫シート 行番号：{result['inventory_row']} を売約済みに更新しました。")
                        st.write("利益：", format_yen(result["profit"]))
                        st.write("利益率：", result["profit_rate"])
                    except Exception as e:
                        st.error("売上登録または在庫更新でエラーが出ました。")
                        st.write(e)

    with tab5:
        st.subheader("使い方ガイド")
        st.write("メルカリちゃんの基本的な使い方を確認できます。")

        guide_col1, guide_col2 = st.columns(2)

        with guide_col1:
            with st.container(border=True):
                st.write("### 1. 買付登録")
                st.write("仕入れた商品の写真、商品名、購入価格、色、サイズ、数量などを登録します。")
                st.write("写真はGoogle Driveの「メルカリちゃん買付写真」フォルダに保存され、写真URLが買付登録シートに入ります。")

            with st.container(border=True):
                st.write("### 2. 買付・AI出品登録")
                st.write("買付情報、在庫登録、AI商品説明作成までまとめて行います。")
                st.write("写真は買付登録と同じ保存処理でGoogle Driveに保存されます。")

        with guide_col2:
            with st.container(border=True):
                st.write("### 3. AI出品サポート")
                st.write("写真や商品情報から、メルカリ用のタイトル・商品説明・価格提案を作成します。")

            with st.container(border=True):
                st.write("### 困ったとき")
                st.write("- 写真保存でエラーが出る場合は、Apps Script URLと保存トークンを確認してください。")
                st.write("- スプレッドシート保存でエラーが出る場合は、サービスアカウントの権限を確認してください。")
                st.write("- LINEから開いた場合も、まず買付登録タブから登録できます。")

        st.write("### 2026/06/25 更新：売上登録・メルカリ購入メールから登録")

        update_col1, update_col2 = st.columns(2)

        with update_col1:
            with st.container(border=True):
                st.write("### 売上登録の使い方")
                st.write("- 売上登録タブを開きます。")
                st.write("- 売れた商品を在庫一覧から選びます。")
                st.write("- 販売日、販売価格、送料、梱包資材費を入力します。")
                st.write("- メルカリ手数料、利益、利益率を確認します。")
                st.write("- 保存ボタンを押します。")
                st.write("- 売上管理シートに保存されます。")
                st.write("- メルカリちゃん在庫シートの在庫状況が「売約済み」になります。")

            with st.container(border=True):
                st.write("### メルカリ購入メールから登録")
                st.write("- メルカリの購入通知メール本文をコピーします。")
                st.write("- 売上登録タブの「メルカリ購入メールから登録」に貼り付けます。")
                st.write("- 「メール内容を読み取る」を押します。")
                st.write("- 購入者名、メルカリ商品ID、商品名、販売価格が自動で入ります。")
                st.write("- 売約済みにする在庫商品を選びます。")
                st.write("- 送料、梱包資材費、販売日を入力します。")
                st.write("- 利益計算を確認して、保存ボタンを押します。")

        with update_col2:
            with st.container(border=True):
                st.write("### まとめ商品の場合")
                st.write("- 商品名に「まとめ商品」「リクエスト」「2点」「3点」「複数」が含まれる場合は注意表示が出ます。")
                st.write("- まとめ商品の場合は複数の在庫を選べます。")
                st.write("- 売上金額と利益は売上管理シートに1行で記録します。")
                st.write("- 選択した在庫はすべて「売約済み」になります。")

            with st.container(border=True):
                st.write("### 利益計算")
                st.write("- メルカリ手数料 = 販売価格 × 10%")
                st.write("- 利益 = 販売価格 - 仕入価格 - メルカリ手数料 - 送料 - 梱包資材費")
                st.write("- 利益率 = 利益 ÷ 販売価格")

            with st.container(border=True):
                st.write("### 現在できる流れ")
                st.write("買付登録")
                st.write("↓")
                st.write("写真保存")
                st.write("↓")
                st.write("AI出品文作成")
                st.write("↓")
                st.write("在庫登録")
                st.write("↓")
                st.write("売上登録")
                st.write("↓")
                st.write("メルカリ購入メールから売上登録")
                st.write("↓")
                st.write("利益計算")
                st.write("↓")
                st.write("在庫を売約済みに更新")
