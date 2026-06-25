import os
import traceback

from dotenv import load_dotenv
from fastapi import FastAPI, Request

load_dotenv()

# WebhookサーバーではStreamlit画面を起動せず、app.py内のLINE処理関数だけを使います。
os.environ["MERCARI_WEBHOOK_IMPORT"] = "1"

import app as mercari_app

app = FastAPI(title="メルカリちゃん LINE Webhook")


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/line/webhook")
async def line_webhook(request: Request):
    try:
        payload = await request.json()
        mercari_app.process_line_purchase_judge_webhook(payload)
        return {"status": "ok"}
    except Exception as error:
        print("LINE webhook error:", error)
        traceback.print_exc()
        return {"status": "error", "message": str(error)}
