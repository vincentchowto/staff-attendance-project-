import pandas as pd
import os 
import tkinter
import numpy as np
from tkinter import fieldialog
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import DataFrame
from collections import Counter
from functools import reduce
from decimal import Decimal
import datetime
from datetime import timedelta

from CLICK_ME_IT_STAFF_Attendance import dfinprocess

#Extracting the Start Time and End Time usable data
dfinprocess['Start Time w hour']=pd.to_datetime(dfinprocess['Start Time'],format='%Y%m%d %H:%M:%S').dt.strftime('%Y-%m-%d %H:%M:%S')
ddfinprocess['End ztime w hour']=pd.to_datetime(dfinprocess('End Time'],format='%Y%m%d %H:%M:%S').dt.strftime('%Y-%m-%d %H:%M:%S')

##Counting (Testing)
start_hours=pd.to_datetime(dfinprocess['Start Time w hour'])
end_hours=pd.to_datetime(dfinprocess['End Time w hour'])
dfinprocess=pd.Series(dt.date()for group in [pd.date_range, end, freq='H')for start, end in zip(start_hours,end_hours)] for dt in group).values_counts()

print(dfinprocess)

IT_STAFF_Attendance_Door_Access_Function.py
import pandas as pd
import os
import tkinter 
from tkinter import fiedialog
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import DataFrame
from collections import Counter
from functools import reduce
from decimal import Decimal
import datetime
from datetime import timedelta

#Reading different required files
pre=os.path.dirname(os.path.realpath(_file_))
frame1='ITStaffList.xlsx'
path1=os.path.join(pre,frame1)
df1=pd.read_excel(path1,'Sheet1')

frame2='domain_user.xlsx'
path2=os.path.join(pre.frame1)
df2=pd.read_excel(path2)

df3=pd.DataFrame()
path3=os.listdir(pre)
for file in path3:
	if file.startswith('Attendance'):
		df3=df3.append(pd.read_excel(file),ignore_index=True)
df3.head()

#Part for modifying the format of ITStaffList excel file
df1.drop('Rank',inplace=True,axis=1)
df1.drop('Type',inplace=True,axis=1)
df1=df1.rename(columns＝{'Staff Code''Staff ID'})
df1=df1[['Staff ID','User Name','IsPerm?','IT Team']]
df1['Staff ID
]=df1['Staff ID'].astype(str)

#Part for modifying the format of Domain User excel file
df2=df2.loc['DEPT_ID']=='IT']
df2=df2.rename(columns={'STAFF_ID'::'STAFF ID'})
df2.drop(['DOMAIN_ID','DISPLAY_NAME','EMIL',"ACTIVE''Action;],inplace=True, axis=1_

#Part for modifying the raw data of Door access record
#Change Value and title to make it standardize
df3['Channel']=df3['Entry Door']
df3.drop{['Entry Door','Exit Door','Cardholder Division','FirstName'], inplace=True, axis=1]
df3['Last Name']=df3['Last Name'].astype(str)

#Standardize the time format and calculate thr duration
ndf3['Entry Time']=pd.to_Datetime(df3['Entry Time'].dt.strftime('%H:%M:%S')
df3['Exit Time']=pd.to_datetime(df3['Exit Time'].dt.strtime('%H:%M:%S')

df3['Entry Date']=pd.to_datetime)df3['Entry Date'])dt.strftime('%d/%m/%Y')
df3['Exit Date']=pd.to_datetime)df3['Exit Date'])dt.strftime('%d/%m/%Y')

df3['Start Time']=df3['Entry Date']+' '+df3['Entry Time']
df3['End Time']=df3['Exit Date']+' '+df3[Exit Time']

df3['Start Time']=pd.to_datetime(df3['Start Time'],format='%d/%m/%Y %H:%M:%S')
df3['End Time']=pd.to_datetime(df3['End Time'],format='%d/%m/%Y %H:%M:%S')

#Converting time format form dd mm yyyy to yyyy mm dd
df3['Start Time']=pd.to_datetime(df3['Start Time'], format='%d/%m/%Y %H:%M:%S').dt.strftime('%Y/%m/%d %H:%M:%S')
df3['End Time']=pd.to_datetime(df3['End Time'],format='%d/%m/%Y %H:%M:%S').dt.strftime('%Y/%m/%d %H:%M:%S')

#change format make it possibly to show on the excel file
df3['Duration']=pd.to_datetime(df3['Exit Time'])-pd.to_datetime(df3['Entry Time'])
df3['Duration']=pd.to_timedelta(df3['Duration'])
df3['Duration']=df3['Duration']-pd.to_timedelta(df3['Duration'].dt.days, unit='d')
df3['Duration']=df3['Duration'].astype(str)
df3['Duration']=df3['Duration'].str.replace('0 days','').astype(str)
print(df1, df2, df3)

#Code for merging all required dataframe
#dfy
dfy=df1[['Staff ID','User Name','Is Perm?','IT Team']].merge(df2[['DEPT_ID','Staff ID','STAFF_LOGIN']],on='Staff ID',how='right')
print(dfy)
#dfda(Code for modifying the output of this part)
dfda=pd.merge(dfy,df3,left_on=['Staff ID'],right_on=['Last Name'],how='right')
dfda=dfda.dropna()
dfda.drop(['DEPT_ID','STAFF_LOGIN','Last Name','Entry Date','Exit Date','Entry Time','Exit Time','Channel'],inplace=True,axis=1)
dfda.insert(4,'Channel','Office')
dfda.insert(4,'Department','IT')
dfda=dfda[['Staff ID','User Name','IsPerm?','IT Team','Channel','Start Time','End Time','Duration','Department']]

#Testing
print(dfda)
#Print as excel
