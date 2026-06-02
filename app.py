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
    "価格理由"
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

def generate_product_id():
    return "MC-" + uuid.uuid4().hex[:8].upper()

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
    if not uploaded_files:
        return [], ""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_product_name = make_safe_file_name(product_name)
    folder_name = f"{product_id}_{safe_product_name}_{timestamp}"
    folder_path = os.path.join("uploads", folder_name)
    os.makedirs(folder_path, exist_ok=True)

    saved_file_names = []
    for index, uploaded_file in enumerate(uploaded_files, start=1):
        safe_original_name = make_safe_file_name(uploaded_file.name)
        file_name = f"{index:02d}_{safe_original_name}"
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        saved_file_names.append(file_name)

    return saved_file_names, folder_path

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
        data.get("price_reason", "")
    ]

    all_values = worksheet.get_all_values()
    next_row = len(all_values) + 1
    worksheet.update(f"A{next_row}:W{next_row}", [row])
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

st.set_page_config(
    page_title="メルカリちゃん",
    page_icon="🛍️",
    layout="wide"
)

st.title("🛍️ メルカリちゃん")
st.write("仕入れた商品の在庫管理と、メルカリ出品準備をサポートするAI社員です。")

tab1, tab2, tab3, tab4 = st.tabs(["在庫登録", "AI出品サポート", "買付登録フォーム", "買付・AI出品登録"])

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
                with st.spinner("写真を保存しています。"):
                    product_id = generate_product_id()
                    saved_file_names, photo_folder_path = save_uploaded_purchase_photos(
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
                    "photo_file_path": photo_folder_path
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
                    "photo_folder_path": photo_folder_path
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
                st.session_state["v3_photo_folder_path"] = photo_folder_path

                st.success("買付登録・在庫保存・AI商品説明作成が完了しました！")
                st.info(f"買付登録保存先：{PURCHASE_SHEET_NAME}")
                st.info(f"在庫保存先：{WORKSHEET_NAME}")
                st.info(f"商品説明ログ保存先：{DESCRIPTION_SHEET_NAME}")
                st.info(f"txt保存先：{saved_file_path}")
                st.info(f"写真保存先：{photo_folder_path}")

            except Exception as e:
                st.error("買付・AI出品登録でエラーが出ました。")
                st.write(e)

    if "v3_description_result" in st.session_state:
        saved = st.session_state["v3_description_result"]

        st.write("### AI生成結果")
        st.write("商品ID：", st.session_state.get("v3_saved_product_id", ""))
        st.write("写真保存先：", st.session_state.get("v3_photo_folder_path", ""))

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
