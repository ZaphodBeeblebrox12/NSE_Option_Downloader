# -*- coding: utf-8 -*-
"""
Created on Wed Apr 15 14:36:43 2020

@author: inderjeet

Script downloads in pdf the options chain for near and next month
"""


import pdfkit
import datetime as date
import calendar
import os.path
import getpass
from pandas import read_excel as read_excel
from pandas import to_datetime
from dateutil.relativedelta import *

xl_file = read_excel(r'C:\Users\inderjeet\Calendar\nse.xlsx')


# xl_file.set_index("Date", inplace = True)
# xl_file.drop("Market Segment", inplace = True, axis =1)
# xl_file.index =  to_datetime(xl_file.index, format='%Y-%m-%d')

# if date.datetime.today().strftime('%Y-%m-%d') not in xl_file.index:
  
# if date.datetime.today().strftime('%Y-%m-%d') in dates.strftime('%Y-%m-%d') and date.datetime.today() == date.datetime.weekday:
todayyear= date.datetime.today().year
todaymonth = date.datetime.today().month
now = date.datetime.today()
now1 = date.datetime.today().strftime("%d%b%Y")
cc = calendar.TextCalendar(calendar.SUNDAY)  
monthly_calander = cc.formatmonth(todayyear,todaymonth)
datethurs = now + relativedelta(day=31, weekday=TH(-1))
expiry = datethurs.strftime("%d%b%Y").upper()

a = date.datetime.today().month
b = date.datetime.strptime(str(a) , '%m')
current_month_name = b.strftime('%B')

c = (date.datetime.today()) + relativedelta(months =1)

username = getpass.getuser()


#path = os.path.join("home","{username}","Data")
# in case of window you will need to add drive "c:" or "d:" before os.path.sep
path = os.path.join("Option Chain",  f'{current_month_name}', f"{expiry}")
relative_path = os.path.relpath(path, os.getcwd())
os.makedirs(relative_path, exist_ok=True)
#start = os.getcwd()
destination = os.path.join(relative_path, 'data.pdf')

today = date.date.today()
first_day_nextmonth = today.replace(day=1) + relativedelta(months = 1)
datenextthurs = first_day_nextmonth + relativedelta(day = 31, weekday = TH(-1))
nextexpiry = datenextthurs.strftime("%d%b%Y").upper()

path1 = os.path.join("Option Chain",  f'{current_month_name}', f"{nextexpiry}")
relative_path1 = os.path.relpath(path1, os.getcwd())
os.makedirs(relative_path1, exist_ok=True)

a = date.datetime.today().month
b = date.datetime.strptime(str(a) , '%m')
current_month_name = b.strftime('%B')
#next_month_name = b + relativedelta(months =1)
#next_month_folder = month_name.strftime('%B')
pdfkit.from_url(f'https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?segmentLink=17&instrument=OPTIDX&symbol=NIFTY&date={expiry}', f'{relative_path}/{now1}.pdf')
pdfkit.from_url(f'https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?segmentLink=17&instrument=OPTIDX&symbol=NIFTY&date={nextexpiry}', f'{relative_path1}/{now1}.pdf')
  
  
#else:
#    print(xl_file.loc[date.datetime.today().strftime('%Y-%m-%d')])

