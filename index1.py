import win32com.client

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
inbox = outlook.GetDefaultFolder(6) #受信トレイは「6」を指定

#受信ボックスの情報を取得
print('Name: ' + inbox.name)
print('Count: ' + str(len(inbox.Items)))

#受信ボックスのフォルダを取得
folders = inbox.Folders
for folder in folders:
    print('Name: ' + folder.name)
    Folder = folders(folder.name)
    print('Message: ' + str(len(Folder.Items)))