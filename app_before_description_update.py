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

def connect_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scopes
    )

    gc = gspread.authorize(credentials)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

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

def generate_product_id():
    return "MC-" + uuid.uuid4().hex[:8].upper()

def clean_json_text(text):
    text = text.strip()
    text = text.replace("```json", "").replace("```", "").strip()
    return text

def generate_description(product_name, category, brand, color, size, condition, features, target, memo):
    prompt = f"""
あなたはメルカリで売れる商品説明文を作るプロです。
誇大表現を避け、実物に忠実で、読みやすく、購入につながる日本語で作成してください。

以下の情報をもとに、JSON形式のみで出力してください。
キーは必ず以下にしてください。

title
description
bullet_points
hashtags

商品名: {product_name}
カテゴリ: {category}
ブランド: {brand}
カラー: {color}
サイズ: {size}
状態: {condition}
特徴: {features}
ターゲット: {target}
メモ: {memo}

条件:
- title はメルカリ向けに分かりやすく
- description はそのままコピペできる完成形
- bullet_points は3〜5個
- hashtags は5〜8個
- 日本語で出力
- 中古品の場合は、状態確認を促す自然な文を入れる
- 「絶対売れる」「新品同様」など断定や誇大表現は避ける
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": "あなたはフリマアプリの商品説明作成の専門家です。"},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7
    )

    text = response.choices[0].message.content
    text = clean_json_text(text)
    return json.loads(text)

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

                st.write("商品ID：", product_id)
                st.write("商品名：", product_name)
                st.write("カテゴリ：", category)
                st.write("仕入れ価格：", purchase_price)
                st.write("仕入日：", purchase_date)
                st.write("予定販売価格：", planned_price)
                st.write("送料：", shipping_fee)
                st.write("梱包資材費：", packing_fee)
                st.write("在庫状況：在庫あり")
                st.write("メモ：", memo)

            except FileNotFoundError:
                st.error("service_account.json が見つかりません。mercari-chan フォルダに入れてください。")
            except Exception as e:
                st.error("エラーが発生しました。")
                st.write(e)

with tab2:
    st.subheader("AI出品サポート")
    st.write("まずは、メルカリ用の商品説明文を作ります。")

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
        ai_target = st.text_input("ターゲット（例：40代女性、学生、旅行好き など）")
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
                        ai_target,
                        ai_memo
                    )

                    st.success("商品説明文を作成しました！")

                    st.write("### タイトル")
                    st.write(result.get("title", ""))

                    st.write("### 商品説明")
                    st.text_area(
                        "コピペ用",
                        value=result.get("description", ""),
                        height=250
                    )

                    st.write("### おすすめポイント")
                    for item in result.get("bullet_points", []):
                        st.write(f"- {item}")

                    st.write("### ハッシュタグ")
                    st.write(" ".join(result.get("hashtags", [])))

                except Exception as e:
                    st.error("商品説明の生成でエラーが出ました。")
                    st.write(e)
