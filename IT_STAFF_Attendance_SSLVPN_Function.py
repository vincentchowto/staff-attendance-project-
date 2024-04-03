import pandas as pd
import os
import datetime as dt
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import DataFrame
from datetime import datetime

#import SSLVPN record
pre=os.path.dirname(os.path.realpath(_file_))
frame1='ITStaffList.xlsx')
path1=os.path.join(pre,frame1)
df1=pd.read_excel(path1)

frame2='domin_user.xlsx'
path2=os.path.join(pre,frame2)
df2=pd.read(_excel(path2)

df3=pd.DataFrame()
path3=os.listdir(pre)
for file in path3:
	if file.startswith('SSLVPN_Access_summary_report'):
		df3=df3.append(pd.read_csv(file,sep" ",skiprows=2,names=['Info','Device ID','Start Date','Start Time','Username','IP Address','Duration','a']))


#Delete unwanted columns and unwanted characters in strings
df3.drop(['Info','Device ID','IP Address','a'], inplace=True, axis=1)
df3['Duration']=df3['Duration'].str.lstrip('duration=')
df3['Start Date']=df3['Start Date'].str.lstrip('time="')
df3['Username']=df3['Username'].str.lower()
df3=df3.rename(columns={'Username':'STAFF_LOGIN'})
df3['Start Date']=df3['Start Date'].str.rstrip('""')

