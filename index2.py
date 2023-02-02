import win32com.client
import pandas as pd

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
inbox = outlook.GetDefaultFolder(6)

# 受信トレイのフォルダ情報を取得
folders = inbox.Folders

# TestAフォルダのアイテムを取得
messages = folders('TestA').Items

# 空のデータフレームを作成
df_mail = pd.DataFrame()

# メッセージ内容を順次取得してデータフレームに追加
i = 0
for message in messages: 
    df_mail.loc[i,"receivedtime"] = pd.to_datetime(str(message.ReceivedTime)[:-6])
    df_mail.loc[i,"sender"] = str(message.Sender)
    df_mail.loc[i,"subject"] = str(message.Subject)
    df_mail.loc[i,"body"] = str(message.body)
    i +=1

df_mail.to_excel('pd_data.xlsx', sheet_name='new_sheet', header=False, index=False)