import os
import traceback

from dotenv import load_dotenv
from fastapi import BackgroundTasks, FastAPI, Request

load_dotenv()

# WebhookサーバーではStreamlit画面を起動せず、app.py内のLINE処理関数だけを使います。
os.environ["MERCARI_WEBHOOK_IMPORT"] = "1"

import app as mercari_app

app = FastAPI(title="メルカリちゃん LINE Webhook")


def run_purchase_judge(payload):
    try:
        mercari_app.process_line_purchase_judge_webhook(payload)
    except Exception as error:
        print("LINE webhook background error:", error)
        traceback.print_exc()


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/line/webhook")
async def line_webhook(request: Request, background_tasks: BackgroundTasks):
    try:
        payload = await request.json()
        background_tasks.add_task(run_purchase_judge, payload)
        return {"status": "ok"}
    except Exception as error:
        print("LINE webhook error:", error)
        traceback.print_exc()
        return {"status": "error", "message": str(error)}
