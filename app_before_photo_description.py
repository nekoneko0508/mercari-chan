import os
import json
import uuid
from datetime import date, datetime

import streamlit as st
import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from openai import OpenAI

load_dotenv()

SPREADSHEET_ID = "1q08itPY88CzG0yrQschTAMNRcVqy5-RwEASl4GkZor0"
WORKSHEET_NAME = "メルカリちゃん在庫"
DESCRIPTION_SHEET_NAME = "商品説明ログ"
SERVICE_ACCOUNT_FILE = "service_account.json"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

HEADERS = [
    "商品ID",
    "登録日",
    "仕入日",
    "商品名",
    "カテゴリ",
    "仕入れ価格",
    "予定販売価格",
    "販売価格",
    "メルカリ手数料",
    "送料",
    "梱包資材費",
    "利益",
    "利益率",
    "在庫状況",
    "販売日",
    "メモ",
    "商品説明文",
    "写真フォルダURL"
]

DESCRIPTION_HEADERS = [
    "作成日時",
    "商品名",
    "カテゴリ",
    "ブランド",
    "カラー",
    "サイズ",
    "状態",
    "タイトル",
    "商品説明",
    "おすすめポイント",
    "ハッシュタグ",
    "ターゲット",
    "販売予定価格",
    "写真から見える特徴",
    "補足メモ"
]

def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scopes
    )

    gc = gspread.authorize(credentials)
    return gc.open_by_key(SPREADSHEET_ID)

def connect_sheet():
    spreadsheet = get_spreadsheet()

    try:
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=WORKSHEET_NAME,
            rows=1000,
            cols=len(HEADERS)
        )

    values = worksheet.get_all_values()
    if not values:
        worksheet.append_row(HEADERS)

    return worksheet

def connect_description_sheet():
    spreadsheet = get_spreadsheet()

    try:
        worksheet = spreadsheet.worksheet(DESCRIPTION_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=DESCRIPTION_SHEET_NAME,
            rows=1000,
            cols=len(DESCRIPTION_HEADERS)
        )

    values = worksheet.get_all_values()
    if not values:
        worksheet.append_row(DESCRIPTION_HEADERS)

    return worksheet

def generate_product_id():
    return "MC-" + uuid.uuid4().hex[:8].upper()

def clean_json_text(text):
    text = text.strip()
    text = text.replace("```json", "").replace("```", "").strip()
    return text

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
    planned_price
):
    prompt = f"""
あなたはメルカリで「思わず商品説明を読みたくなるタイトル」と
「購入判断がしやすい商品説明文」を作る専門家です。

以下の情報をもとに、メルカリ出品用のタイトルと商品説明を作成してください。
必ずJSON形式のみで出力してください。

キーは必ず以下にしてください。
title
description
bullet_points
hashtags

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

【商品説明条件】
- コンセプトは「5分で売れる商品説明」
- 冒頭2行で購入者の興味を引く
- 何の商品か、どんな人におすすめかを分かりやすく書く
- 写真から見える特徴を自然に反映する
- 状態、サイズ、色、発送予定、注意点を整理する
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
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {
                "role": "system",
                "content": "あなたはメルカリ出品文の専門家です。実物に忠実で、購入されやすい文章を作ります。"
            },
            {
                "role": "user",
                "content": prompt
            }
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
        data.get("memo", "")
    ]

    worksheet.append_row(row)

st.set_page_config(
    page_title="メルカリちゃん",
    page_icon="🛍️",
    layout="wide"
)

st.title("🛍️ メルカリちゃん")
st.write("仕入れた商品の在庫管理と、メルカリ出品準備をサポートするAI社員です。")

tab1, tab2 = st.tabs(["在庫登録", "AI出品サポート"])

with tab1:
    st.subheader("商品登録")

    product_name = st.text_input("商品名", key="reg_product_name")
    category = st.text_input("カテゴリ", key="reg_category")
    purchase_price = st.number_input("仕入れ価格", min_value=0, step=100, key="reg_purchase_price")
    purchase_date = st.date_input("仕入日", value=date.today(), key="reg_purchase_date")
    planned_price = st.number_input("予定販売価格", min_value=0, step=100, key="reg_planned_price")
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
                    ""
                ]

                worksheet.append_row(row)

                st.success("スプレッドシートに商品を登録しました！")
                st.info(f"保存先：{WORKSHEET_NAME}")

            except FileNotFoundError:
                st.error("service_account.json が見つかりません。mercari-chan フォルダに入れてください。")
            except Exception as e:
                st.error("エラーが発生しました。")
                st.write(e)

with tab2:
    st.subheader("AI出品サポート")
    st.write("メルカリ用の40文字以内タイトルと、5分で売れる商品説明文を作ります。")

    if not OPENAI_API_KEY:
        st.warning(".env に OPENAI_API_KEY が設定されていません。")
    else:
        ai_product_name = st.text_input("商品名", key="ai_product_name")
        ai_category = st.text_input("カテゴリ", key="ai_category")
        ai_brand = st.text_input("ブランド", key="ai_brand")
        ai_color = st.text_input("カラー", key="ai_color")
        ai_size = st.text_input("サイズ", key="ai_size")
        ai_condition = st.selectbox(
            "商品の状態",
            ["新品・未使用", "未使用に近い", "目立った傷や汚れなし", "やや傷や汚れあり", "傷や汚れあり"]
        )
        ai_features = st.text_area("商品の特徴")
        ai_photo_features = st.text_area(
            "写真から見える特徴",
            placeholder="例：箱に少しへこみあり／ロゴ入り／ストロー付き／色はベージュ寄り など"
        )
        ai_target = st.text_input("ターゲット（例：40代女性、学生、旅行好き など）")
        ai_planned_price = st.number_input("販売予定価格", min_value=0, step=100, key="ai_planned_price")
        ai_memo = st.text_area("補足メモ")

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

                    st.session_state["description_result"] = {
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
                        "title": result.get("title", ""),
                        "description": result.get("description", ""),
                        "bullet_points": result.get("bullet_points", []),
                        "hashtags": result.get("hashtags", [])
                    }

                    st.success("商品説明文を作成しました！")

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

            if st.button("商品説明をスプレッドシートに保存する"):
                try:
                    save_description_to_sheet(saved)
                    st.success("商品説明ログに保存しました！")
                    st.info(f"保存先：{DESCRIPTION_SHEET_NAME}")
                except Exception as e:
                    st.error("保存でエラーが出ました。")
                    st.write(e)
