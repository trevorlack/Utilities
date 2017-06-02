import sys 
import win32com.client
import os
import glob
import pandas as pd
from FTP_Credentials import FTP_Auth
import xlsxwriter

Host, Login, Password = FTP_Auth()
os.chdir('R:/Fixed Income/IVY/Index Holdings')
DateList = []

for file in list(glob.glob("*SP5MAIG*.S*")):
    DateList.append(file[0:8])

Max_Date = int(max(DateList))

MySite = win32com.client.Dispatch('CuteFTPPro.TEConnection') 

MySite.Protocol = 'FTP' 
MySite.Host = Host
MySite.Login = Login
MySite.Password = Password
#MySite.UseProxy = 'BOTH'
MySite.Connect() 

if not MySite.IsConnected: 
    print('Could not connect to: %s Aborting!' % MySite.Host)
    sys.exit(1)
else: 
    print('You are now connected to: %s' % MySite.Host)

MySite.LocalFolder = 'R:/Fixed Income/IVY/Index Holdings'
MySite.RemoteFolder = '/Inbox'
MySite.RemoteFilterInclude = '*SP5MAIG*;'
#MySite.Download('*SP5MAIG*')
#Result = MySite.GetList (MySite.RemoteFolder, "C:/Users/tlack/Documents/Python Scripts/Yieldbook/ftplist.txt", "*SP5*")
Result = MySite.GetList("/Inbox","C:/temp_list.txt","%NAME")
FileLister = MySite.GetResult
FTP_list = pd.read_table('C:/temp_list.txt', header=None, names='p')
FTP_list = pd.DataFrame(FTP_list)
FTP_list['date'] = FTP_list['p'].str[0:8]
#FTP_list = FTP_list.rename(columns={'':'path'})
print(FTP_list)
counter = 0
prior = 0
CLS_Convert_List = []
for i in range(0, len(FTP_list)-1):
    checker = int(FTP_list['date'].iloc[i])
    file = FTP_list['p'].iloc[i]
    print(checker)
    print(file)
    if checker > Max_Date:
        MySite.Download(file)
        if checker != prior:
            counter = counter + 1
        CLS_Convert_List.append(str(checker))
        prior = checker
if counter == 0:
    print('No new Investment Grade Index Files')
else:
    for i in range(0,len(CLS_Convert_List)-1):
        index = pd.read_csv('R:/Fixed Income/IVY/Index Holdings/'+CLS_Convert_List[i]+'_SP5MAIG_CLS.SPFIC', sep='\t')
        writer = pd.ExcelWriter('R:/Fixed Income/IVY/Index Holdings/'+CLS_Convert_List[i]+'_SP5MAIG_CLS.xlsx')
        index.to_excel(writer, 'SP5MAIG_CLS', index=False)
        writer.save()
    print(counter, 'Day(s) of Index Files Saved Down')

MySite.Disconnect()
MySite.TECommand('exit')
print(MySite.Status)


