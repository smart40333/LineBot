# -*- coding: utf-8 -*-
import os
import pandas as pd
from flask import Flask, request, abort, send_file
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

app = Flask(__name__)

# 環境變數：LINE 憑證
line_bot_api = LineBotApi(os.getenv('LINE_CHANNEL_ACCESS_TOKEN'))
handler = WebhookHandler(os.getenv('LINE_CHANNEL_SECRET'))

# 預設資料路徑
DATA_DIR = "Data"
MAPPING_PATH = os.path.join(DATA_DIR, "mapping.csv")
NONE_EXIST_PATH = os.path.join(DATA_DIR, "none_exist.csv")

# 確保 Data 資料夾存在
os.makedirs(DATA_DIR, exist_ok=True)

# 讀取 mapping.csv（啟動時一次）
df_con = pd.read_csv(MAPPING_PATH)


@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    uid = event.source.user_id
    try:
        gid = event.source.group_id
        print(f'來自群組: {gid}')
    except AttributeError:
        gid = None

    profile = line_bot_api.get_profile(uid)
    display_name = profile.display_name
    print(f'來自使用者: {display_name}，UID: {uid}')

    # 若 UID 不存在於 mapping.csv，記錄到 none_exist.csv
    if uid not in df_con['LineID'].astype(str).values:
        new_row = pd.DataFrame([[display_name, uid]], columns=["Name", "UID"])

        if os.path.exists(NONE_EXIST_PATH):
            existing_df = pd.read_csv(NONE_EXIST_PATH)
            if uid not in existing_df['UID'].astype(str).values:
                updated_df = pd.concat([existing_df, new_row], ignore_index=True)
                updated_df.to_csv(NONE_EXIST_PATH, index=False)
            else:
                updated_df = existing_df
        else:
            updated_df = new_row
            updated_df.to_csv(NONE_EXIST_PATH, index=False)

        # 印出目前 none_exist.csv 的所有 UID 和名稱
        print("❗ 未配對名單：")
        for _, row in updated_df.iterrows():
            print(f"  👤 {row['Name']}：{row['UID']}")



@app.route("/download/none_exist", methods=['GET'])
def download_none_exist():
    if os.path.exists(NONE_EXIST_PATH):
        return send_file(NONE_EXIST_PATH, as_attachment=True)
    else:
        return "尚未產生 none_exist.csv", 404


if __name__ == "__main__":
    app.run(port=8080)
