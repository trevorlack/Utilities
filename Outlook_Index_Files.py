import win32com.client
import glob
import datetime as date
from MySQL_Date_Check import pull_hy_match_set

#def outlook_file_grab(ref_date):

current_hyhg_data = int(pull_hy_match_set())
print(current_hyhg_data)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders('Scripting')

messages = inbox.Items

message = messages.GetFirst ()

while message:
    print(message.Subject[0:15])
    if message.Subject[0:15] == 'Citi High Yield':
        File_Date = int(message.Subject[-16:-7])
        if File_Date > current_hyhg_data:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            print(attachment.FileName)
            attachment.SaveASFile('C:\\Users\\tlack\\Documents\\Python Scripts\\MySQL Pandas\\Loader\\' + attachment.FileName)
    message = messages.GetNext ()


#    if current_hyhg_data in message.Subject: