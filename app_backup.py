import streamlit as st
from datetime import date, datetime
import uuid
import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "1q08itPY88CzG0yrQschTAMNRcVqy5-RwEASl4GkZor0"
WORKSHEET_GID = 2133283807
SERVICE_ACCOUNT_FILE = "service_account.json"

def connect_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scopes
    )

    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    worksheet = spreadsheet.get_worksheet_by_id(WORKSHEET_GID)

    return worksheet

def generate_product_id():
    return "MC-" + uuid.uuid4().hex[:8].upper()

st.set_page_config(
    page_title="メルカリちゃん",
    page_icon="🛍️",
    layout="centered"
)

st.title("🛍️ メルカリちゃん")
st.write("仕入れた商品の在庫管理とメルカリ出品をサポートするAI社員です。")

st.subheader("商品登録")

product_name = st.text_input("商品名")
category = st.text_input("カテゴリ")
purchase_price = st.number_input("仕入れ価格", min_value=0, step=100)
purchase_date = st.date_input("仕入日", value=date.today())
planned_price = st.number_input("予定販売価格", min_value=0, step=100)
shipping_fee = st.number_input("送料", min_value=0, step=100)
packing_fee = st.number_input("梱包資材費", min_value=0, step=10)
memo = st.text_area("メモ")

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
