# -*- coding: utf-8 -*-
from skpy import Skype, SkypeChats
import requests
import json
import datetime
import os

next_date = datetime.datetime.today() - datetime.timedelta(days=1)
dt = next_date.strftime('%Y/%m/%d')
ft = next_date.strftime('%Y%m%d')
fn = ("檔案名稱前綴"+ ft +".xls")

my_params = {
    'token':'',
    'start_d':dt,
    'end_d':dt,
    'lang':'zh_cn'
    }

# 輸入skype帳號密碼登入
sk = Skype("帳號", "密碼") # connect to Skype

#下載檔案
r = requests.get('後台報表下載URL',params = my_params)
open(fn,'wb').write(r.content)

# 設定群組ID
ch = sk.chats["客戶頻道ID"]

# 傳送訊息與檔案
ch.sendMsg("您好，這是 " + next_date.strftime('%m') + "月" + next_date.strftime('%d') + "號的所有玩家投注額報表下載連結，提供給貴司。")
ch.sendFile(open(fn,"rb"), fn)

os.remove(fn)
