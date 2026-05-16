import os
import json
import uuid
import base64
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
PURCHASE_SHEET_NAME = "買付登録"
SERVICE_ACCOUNT_FILE = "service_account.json"

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
    "写真ファイル名",
    "写真保存先",
    "ステータス"
]

def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    # ローカルでは service_account.json、公開版では Streamlit Secrets から読む
    try:
        if "gcp_service_account" in st.secrets:
            credentials = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=scopes
            )
        else:
            credentials = Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE,
                scopes=scopes
            )
    except Exception:
        credentials = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=scopes
        )

    gc = gspread.authorize(credentials)
    return gc.open_by_key(SPREADSHEET_ID)

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

    values = worksheet.get_all_values()
    if not values:
        worksheet.append_row(headers)

    return worksheet

def connect_sheet():
    return get_or_create_worksheet(WORKSHEET_NAME, HEADERS)

def connect_description_sheet():
    return get_or_create_worksheet(DESCRIPTION_SHEET_NAME, DESCRIPTION_HEADERS)

def connect_purchase_sheet():
    return get_or_create_worksheet(PURCHASE_SHEET_NAME, PURCHASE_HEADERS)

def generate_product_id():
    return "MC-" + uuid.uuid4().hex[:8].upper()

def clean_json_text(text):
    text = text.strip()
    text = text.replace("```json", "").replace("```", "").strip()
    return text

def image_to_base64(uploaded_file):
    image_bytes = uploaded_file.getvalue()
    return base64.b64encode(image_bytes).decode("utf-8")

def save_uploaded_purchase_photo(uploaded_file):
    if uploaded_file is None:
        return "", ""

    os.makedirs("uploads", exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = uploaded_file.name.replace(" ", "_")
    file_name = f"purchase_{timestamp}_{safe_name}"
    file_path = os.path.join("uploads", file_name)

    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    return file_name, file_path

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

【タイトル例】
ストロー付きタンブラー 持ち歩きに便利なブルー
外出用タンブラー ストロー付きで仕事中にも便利
ブルータンブラー ストロー付きでデスク使いにも
韓国風ポーチ 小物整理に便利な淡色デザイン
旅行用ミニバッグ 軽くて持ち歩きやすい黒

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
        data.get("memo", "")
    ]

    worksheet.append_row(row)

def save_purchase_to_sheet(data):
    worksheet = connect_purchase_sheet()

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
        data.get("photo_file_path", ""),
        "買付済み"
    ]

    # 必ずA列からM列に保存する
    all_values = worksheet.get_all_values()
    next_row = len(all_values) + 1

    # 1行目が空の場合は、A1からヘッダーを入れる
    if next_row == 1:
        worksheet.update("A1:M1", [PURCHASE_HEADERS])
        next_row = 2

    worksheet.update(f"A{next_row}:M{next_row}", [row])

st.set_page_config(
    page_title="メルカリちゃん",
    page_icon="🛍️",
    layout="wide"
)

st.title("🛍️ メルカリちゃん")
st.write("仕入れた商品の在庫管理と、メルカリ出品準備をサポートするAI社員です。")

tab1, tab2, tab3 = st.tabs(["在庫登録", "AI出品サポート", "買付登録フォーム"])

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

with tab3:
    st.subheader("買付登録フォーム")
    st.write("買付先で写真・価格・色・サイズ・数量を記録します。")

    purchase_photo = st.file_uploader(
        "商品写真",
        type=["jpg", "jpeg", "png"],
        key="purchase_photo"
    )

    if purchase_photo is not None:
        st.image(purchase_photo, caption="登録する商品写真", width=350)

    buyer_name = st.text_input("登録者名")
    purchase_product_name = st.text_input("商品名", key="purchase_product_name")
    purchase_price = st.number_input("購入価格", min_value=0, step=100, key="purchase_price")
    purchase_color = st.text_input("色")
    purchase_size = st.text_input("サイズ")
    purchase_quantity = st.number_input("数量", min_value=1, step=1, value=1)
    purchase_place = st.text_input("購入場所")
    purchase_category = st.text_input("カテゴリ", key="purchase_category")
    purchase_memo = st.text_area("メモ", placeholder="素材感、傷、付属品、販売時の注意点など")

    if st.button("買付情報を保存する"):
        if purchase_product_name == "":
            st.error("商品名を入力してください。")
        elif buyer_name == "":
            st.error("登録者名を入力してください。")
        else:
            try:
                photo_file_name, photo_file_path = save_uploaded_purchase_photo(purchase_photo)

                data = {
                    "buyer_name": buyer_name,
                    "product_name": purchase_product_name,
                    "purchase_price": purchase_price,
                    "color": purchase_color,
                    "size": purchase_size,
                    "quantity": purchase_quantity,
                    "place": purchase_place,
                    "category": purchase_category,
                    "memo": purchase_memo,
                    "photo_file_name": photo_file_name,
                    "photo_file_path": photo_file_path
                }

                save_purchase_to_sheet(data)

                st.success("買付情報をスプレッドシートに保存しました！")
                st.info(f"保存先：{PURCHASE_SHEET_NAME}")

                st.write("商品名：", purchase_product_name)
                st.write("購入価格：", purchase_price)
                st.write("色：", purchase_color)
                st.write("サイズ：", purchase_size)
                st.write("数量：", purchase_quantity)
                st.write("写真保存先：", photo_file_path if photo_file_path else "写真なし")

            except Exception as e:
                st.error("買付情報の保存でエラーが出ました。")
                st.write(e)
