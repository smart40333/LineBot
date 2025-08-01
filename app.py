# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 14:32:43 2022

@author: WAYNE.HUANG

@Eviroment: 統一正式 API
"""

import pandas as pd
import shutil
import os
#import dataframe_image as dfi
#import openai
import pyodbc
import PW
from os import listdir
from os.path import isfile, join
from datetime import date, datetime, time
import re
#import schedule
#import redis
from openai import OpenAI
client = OpenAI(api_key='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')

from flask import Flask, request, abort, render_template

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage, TemplateSendMessage, PostbackEvent, 
    PostbackTemplateAction, ButtonsTemplate, LocationSendMessage, ImageSendMessage, 
    URITemplateAction, CarouselTemplate, CarouselColumn, MessageTemplateAction
)

"""
counter = pd.read_csv('//10.72.228.228/Users/wayne.huang/AppData/Local/CBAS/counter.csv', header = None).iloc[0,0]
conn = pyodbc.connect(driver = 'ODBC Driver 18 for SQL Server', server = '10.72.228.139', user ='sa', password = 'Self@pscnet', database = 'CBAS', TrustServerCertificate='yes')
admin = []
cur = conn.cursor()
cur.execute(f"SELECT LineID FROM LineID WHERE CUSID = '23218183' or CUSID = 'cbast1' or CUSID = 'cbast2' or CUSID = 'cbast3'")
fetchall = cur.fetchall()
for a in range(len(fetchall)):
    admin.append(fetchall[a][0])
conn.commit()
cur.close()
"""
app = Flask(__name__)

liffid = '1657505414-2KjwZBYy'
line_bot_api = LineBotApi(os.getenv('LINE_CHANNEL_ACCESS_TOKEN'))
handler = WebhookHandler(os.getenv('LINE_CHANNEL_SECRET'))
client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
@app.route('/page')
def page():
    return render_template('query.js', liffid = liffid)

@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    global counter
    cb_list = pd.read_excel(r'\\10.72.228.112\cbas業務公用區\官網規劃\FTP\Market_Info\market_info.xlsx', sheet_name='2已發行CB資料', usecols=['CB代號','CB名稱'])
    cb_dict = dict(zip(cb_list['CB名稱'], cb_list['CB代號']))
    uid = event.source.user_id
    try:
        gid = event.source.group_id
        print(f'Group Name: {gid}')
    except AttributeError:
        pass
    user_input = event.message.text
    user_input = user_input.replace('\n', ' ')
    print(user_input)
    profile = line_bot_api.get_profile(uid)
    display_name = profile.display_name
    print('客戶顯示名稱:'+ display_name + '      客戶Line uid:' + uid)
    
    response = client.responses.create(
    model="gpt-4o-mini",
    input=[
            {
            "role": "user",
            "content": [
                {
                    "type": "input_text",
                    "text": f"""你是一個專業的轉換及精簡語意人員，將我輸入給你的文字輸出成我要的格式。我所需要輸出的格示範利如雙括弧內之文字: "Relevant,99332,B,20,100.5"
                    ***格式說明***
                    1. Relevant: 首先需判斷我給你的文字是否為下單指示，如果是，則輸出Relevant,如果不是，則輸出irrelevant
                    2. 如果Relevant，則判斷標的代號，我有可能給你標的名字，也可能給你標的代號，請參閱以下標的代號表: {cb_dict}
                    3. 如果irrelevant，則僅輸出"irrelevant"這個英文字，其餘不輸出及判斷
                    4. 如果是買進，則顯示"B"，如果是賣出，則顯示"S"
                    5. 判斷張數，並排列在買賣別之後
                    6. 最後為價格，如為100.5，則顯示100.5
                    7. 簡單來說，就是輸出內容照以下排列，不須任何其他文字以及符號: "是否相關","標的代號","買賣別","張數","價格"
                    8. 如果我打出"出"或是"S"都是賣的意思
                    9. 舉例: 84222 S*1 120 就是 84222賣出一張@120的意思
                    10. 如果偵測到兩筆，則依一樣的格式產出兩行


                    以下是我輸入的文字: {user_input}
                    """
                },
                    ]
            }
        ]
    )
    res = response.output_text.strip()
    res = res.replace('"', '')
    print(f'AI: {res}')
    
    
    if 'Relevant' in res and 'irrelevant' not in res:
        #counter = redis_client.get('counter')
        #counter = int(counter) if counter else 0
        numofContracts = res.count('Relevant')
        current_time = datetime.now().strftime('%H%M%S')
        current_date = date.today().strftime('%Y%m%d')
        cur = conn.cursor()
        cur.execute(f"SELECT CUSID FROM LineID WHERE LineID = '{uid}'")
        cusid = cur.fetchone()[0]
        conn.commit()
        cur.close()
        if cusid == 'cbast1':
            cusid = '23218183'
        if numofContracts == 1:
            try:
                if cusid == 'E221651871' and gid == 'Cf7a46d1047921c440acc488da7e1bdf9':
                    res = res.replace('Relevant', '23657154')
                else:
                    res = res.replace('Relevant', cusid)
            except:
                res = res.replace('Relevant', cusid)
            counter += 1
            header = '身分證,證券代號,BS,張數,價格,流水號\n' + res + f',{current_date}{counter}' + '\n' + 'END'
            with open(f'//10.72.228.228/Users/wayne.huang/AppData/Local/CBAS/自動匯入/{cusid}_{current_time}.csv', 'w') as out:
                out.write(header)
            
            with open('//10.72.228.228/Users/wayne.huang/AppData/Local/CBAS/counter.csv', 'w') as out:
                out.write(str(counter))
            
            print(res)
            print(f'Current number is: {counter}')
        
        else:
            res = res.replace('Relevant', cusid)
            lines = res.splitlines()
            p = 0
            for x in range(counter+1, counter + numofContracts + 1):
                lines[p] = lines[p] + f',{current_date}{x}'
                p = p + 1 
                
            finalRes = '\n'.join(lines)
            header = '身分證,證券代號,BS,張數,價格,流水號\n' + finalRes + '\n' + 'END'
            with open(f'//10.72.228.228/Users/wayne.huang/AppData/Local/CBAS/自動匯入/{cusid}_{current_time}.csv', 'w') as out:
                out.write(header)
            
            counter = counter + numofContracts
            
            with open('//10.72.228.228/Users/wayne.huang/AppData/Local/CBAS/counter.csv', 'w') as out:
                out.write(str(counter))
            
            print(finalRes)
            print(f'Current number is: {counter}')
            
    
    if '@相關資訊' == user_input:
        line_bot_api.reply_message(event.reply_token, TemplateSendMessage(
            alt_text = 'Buttorm Template',
            template = ButtonsTemplate(
                title = '市場相關資訊',
                text = '請選擇所需功能',
                actions = [
                    PostbackTemplateAction(
                        label ='近一月預計發行及掛牌CB',
                        text = '近一月預計發行及掛牌CB',
                        data = '近一月預計發行及掛牌CB'
                    ),
                    PostbackTemplateAction(
                        label ='上週熱門CB標的',
                        text = '上週熱門CB標的',
                        data = '上週熱門CB標的'
                    ),
                            ]
                )
            )
        )
    
    elif '@聯絡我們' == user_input:
        line_bot_api.reply_message(event.reply_token, TemplateSendMessage(
            alt_text = 'Buttorm Template',
            template = ButtonsTemplate(
                title = '聯絡資訊',
                text = '請選擇所需功能',
                actions = [
                    URITemplateAction(
                        label = '分公司據點',
                        uri = 'https://www.pscnet.com.tw/pscnetAbout/branch/list.do'
                    ),
                    PostbackTemplateAction(
                        label ='CBAS交易相關',
                        text = 'CBAS交易相關',
                        data = 'CBAS交易相關'
                    ),
                    # PostbackTemplateAction(
                    #     label ='CBAS業務相關',
                    #     text = 'CBAS業務相關',
                    #     data = 'CBAS業務相關'
                    # ),
                    PostbackTemplateAction(
                        label ='帳戶狀態查詢',
                        text = '帳戶狀態查詢',
                        data = '帳戶狀態查詢'
                    ),
                    PostbackTemplateAction(
                        label ='測試區',
                        text = '測試區',
                        data = '測試區'
                    )
                            ]
                )
            )
        )
      
        
    
    #====================================

@handler.add(PostbackEvent)
def handle_postback_trading(event):
    
    sel = event.postback.data
    uid = event.source.user_id
    #user_ai_input = event.message.text

#=======================聯絡我們===========================

    if sel == '總公司':
        location_message = LocationSendMessage(
            title = '統一證券總公司',
            address = '台北市松山區東興路8號',
            latitude = 25.05056,
            longitude = 121.56495
            )
        line_bot_api.reply_message(event.reply_token, location_message)
        
    elif sel == 'CBAS交易相關':
        carouseltrader(event)
        
    elif sel == 'CBAS業務相關':
        carouselsales(event)
        
    elif sel == '帳戶狀態查詢':
        line_bot_api.reply_message(event.reply_token, TemplateSendMessage(
            alt_text = 'Buttorm Template',
            template = ButtonsTemplate(
                title = '帳戶相關',
                text = '請選擇所需功能',
                actions = [
                    PostbackTemplateAction(
                        label ='查詢開戶狀態',
                        text = '查詢開戶狀態',
                        data = '查詢開戶狀態'
                    ),
                    PostbackTemplateAction(
                        label ='查詢銀行授扣狀態',
                        text = '查詢銀行授扣狀態',
                        data = '查詢銀行授扣狀態'
                    )

                            ]
                )
            )
        )

    elif sel == '近一月預計發行及掛牌CB':
        cb_schedule(event)

    elif sel == '上週熱門CB標的':
        hot_cb_weekly(event)
    
    elif sel == '查詢開戶狀態':
        accStatus(event, uid)

    elif sel == '查詢銀行授扣狀態':
        bankAuthorization(event, uid)
        
    # elif sel == '測試區':
        #line_bot_api.push_message(uid, TextSendMessage(text='想問點什麼呢?'))    
        #cur = conn.cursor()
        # cur.execute(f"Update LineID SET OpenAI = 'On' WHERE LineID = '{uid}'")
        # conn.commit()
        # cur.close()
#=======================Def===========================
        
    
def carouseltrader(event): #CBAS交易相關
    
    Contacts = TemplateSendMessage(
        alt_text = '交易室聯絡資訊',
        template = CarouselTemplate(
            columns = [
                CarouselColumn(
                    #thumbnail_image_url = 'https://imgur.com/vS67C9D',
                    #title = '王慕約',
                    text = '統一證券計量交易部',
                    actions = [
                        MessageTemplateAction(
                            label = '電話',
                            text = '02-27463801'
                        ),
                        MessageTemplateAction(
                            label = 'e-mail',
                            text = 'PSC.CBAS@uni-psg.com'
                            ),
                        ]
                    ),
                # CarouselColumn(
                #     #thumbnail_image_url = 'https://imgur.com/uaPHDZH',
                #     #title = '周琬瑜',
                #     text = '統一證券計量交易部',
                #     actions = [
                #         MessageTemplateAction(
                #             label = '電話: 02-27463727',
                #             text = '02-27463727'
                #         ),
                #         MessageTemplateAction(
                #             label = 'e-mail: fish.chou@uni-psg.com',
                #             text = 'fish.chou@uni-psg.com'
                #             ),
                #         ]
                #     )
                ]
            )
        )
    line_bot_api.reply_message(event.reply_token, Contacts)

def carouselsales(event): #CBAS業務相關
    
    Contacts = TemplateSendMessage(
        alt_text = '業務相關聯絡資訊',
        template = CarouselTemplate(
            columns = [
                CarouselColumn(
                    thumbnail_image_url = 'https://i.imgur.com/cELIutA.jpg',
                    title = '黃暐庭',
                    text = '統一證券計量交易部',
                    actions = [
                        MessageTemplateAction(
                            label = '電話',
                            text = '02-27463670\n0926-968-011'
                        ),
                        MessageTemplateAction(
                            label = 'e-mail',
                            text = 'wayne.huang@uni-psg.com'
                            ),
                        ]
                    ),
                CarouselColumn(
                    thumbnail_image_url = 'https://imgur.com/gallery/BRsot2x',
                    title = '曲博麟',
                    text = '統一證券計量交易部',
                    actions = [
                        MessageTemplateAction(
                            label = '電話',
                            text = '02-2747-8266 #3801'
                        ),
                        MessageTemplateAction(
                            label = 'e-mail',
                            text = 'CHUPOLIN@uni-psg.com'
                            ),
                        ]
                    )
                ]
            )
        )
    line_bot_api.reply_message(event.reply_token, Contacts)
        
    
def accStatus(event, uid):
    #uid = 'U03ed999c65659344f17f36427b6975c5' #test
    cur = conn.cursor()
    cur.execute(f"SELECT CUSID FROM LineID WHERE LineID = '{uid}'")
    cusid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    df_acc = pd.read_excel(r'\\10.72.228.112\cbas開戶作業區\CBAS開戶扣款進度.xlsm', sheet_name='新開戶', header=[1], usecols=['號碼', '四_開戶進度'])
    try:
        status = df_acc.iloc[df_acc.loc[df_acc['號碼'] == cusid].index.item(), 1]
        if status == 1:
            reply = '您的開戶已完成，可進行交易，交易前請確認您的自動扣款授權狀態'
    
        elif status == 2:
            reply = '您的KYC(客戶投資能力及風險屬性問卷)已到期、或不符合承作本商品所需之等級。目前僅能履約、不得新作交易。您可以透過統一E指發APP重新填寫KYC，若您有更新KYC，可再通知我們做設定，謝謝您'
    
        else:
            reply = '您的開戶審核中或尚未開戶，待開戶完成後將以Email通知'
    except:
        reply = '您尚未開戶，如有開戶需求還請與我們聯絡'
        
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

def bankAuthorization(event, uid):
    #uid = 'U03ed999c65659344f17f36427b6975c5' #test
    cur = conn.cursor()
    cur.execute(f"SELECT CUSID FROM LineID WHERE LineID = '{uid}'")
    cusid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn_as400 = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
    #cusid = '23218183' #test
    df_bank = pd.read_sql(f"select a.CUSID, a.BNKFNM, a.BNKACTNO, b.ADMARK, b.RCODE FROM FSPFLIB.FSPCS0M a LEFT JOIN FSPFLIB.FSPEACH01 b ON a.CUSID = b.CUSID WHERE a.CUSID = '{cusid}'", conn_as400)
    bankName = df_bank.iloc[0,1]
    bankNumber = df_bank.iloc[0,2]
    aDMARK = df_bank.iloc[0,3]
    rCODE = df_bank.iloc[0,4]
    if len(df_bank) > 0:
        try:    
            if aDMARK == 'A' and rCODE == '0':
                reply = f'您的自動扣款授權已完成，採T+2日交割，早上8:30扣款。\n您的扣款帳號為：\n{bankName.strip()}(*****{bankNumber.strip()[-5:]})'
            elif aDMARK == 'A' and rCODE != 0:
                reply = '您的自動扣款授權申請中，完成後將以Email通知'
            elif aDMARK != 'A':
                reply = '您尚未申請自動扣款，如需申請請聯絡您的業務人員，配合銀行可參考以下網址：\nhttps://cbas16889.pscnet.com.tw/createAccount'
                print('S')
            else:
                reply = '您尚未開戶，如有開戶需求還請與我們聯絡'
        
        except:
            reply = '您尚未開戶，如有開戶需求還請與我們聯絡'
    
    elif len(df_bank) == 0:
        reply = '您尚未開戶，如有開戶需求還請與我們聯絡'
    
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))


if __name__ == "__main__":

    app.run(port=8080)