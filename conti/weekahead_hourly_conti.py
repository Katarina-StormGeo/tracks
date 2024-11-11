import json
import pandas as pd
import xarray as xr
import yaml
import time
from itertools import dropwhile
import sys 
import os
import subprocess
import requests
import argparse
from io import StringIO
import logging
from logging.handlers import TimedRotatingFileHandler
import numpy as np
import traceback
import operator
import re
from dateutil.relativedelta import relativedelta
import glob
from datetime import datetime, timedelta, date
from calendar import monthrange, month_name

path = r'M:/EUROPA/Power Balance/Conti Model/History/Different runs within day'
files = os.listdir(path)


start_date = date(2023,1,1).strftime("%Y-%m-%d")
end_date = date(2023,12,31).strftime("%Y-%m-%d")

list_of_dates = pd.date_range(start_date, end_date)

last_file_in_date = []

for i in list_of_dates:
    date_files = []
    for j in files:
        if i.strftime("%Y-%m-%d") in j:
            date_files.append(j)
    if date_files:
        datoer = []
        for k in date_files:
            dato = datetime.strptime(k,  'result-ALL_%Y-%m-%d T%H%M.xls')
            datoer.append(dato)
        datoer = sorted(datoer)
        date_files_sorted = []
        for k in datoer:
            filename = k.strftime('result-ALL_%Y-%m-%d T%H%M.xls')
            date_files_sorted.append(filename)
        last_file_in_date.append(date_files_sorted[-1])


countries = ['PowerBalance-DE','PowerBalance-BE','PowerBalance-FR']
germany = []
belgium = []
france = []
lists = [germany, belgium, france]


for k in last_file_in_date:
    for i in range(0,3):

        df = pd.read_excel(path + '/' + k, index_col = 0, sheet_name=countries[i])

        day = (datetime.strptime(k[11:21], '%Y-%m-%d')).weekday()

        
        start_w0 = (datetime.strptime(k[11:21], '%Y-%m-%d')).strftime("%Y-%m-%d") + " 00:00:00"
        end_w0 = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=6 - day)).strftime("%Y-%m-%d") + " 23:00:00"
        w0 = df[start_w0:end_w0]['Price']

        start_w1 = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=7 - day)).strftime("%Y-%m-%d") + " 00:00:00"
        end_w1 = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=13 - day)).strftime("%Y-%m-%d") + " 23:00:00"
        w1 = df[start_w1:end_w1]['Price']

        start_w2 = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=14- day)).strftime("%Y-%m-%d") + " 00:00:00"
        end_w2 = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=20- day)).strftime("%Y-%m-%d") + " 23:00:00"
        w2 = df[start_w2:end_w2]['Price']

        start_w3 = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=21- day)).strftime("%Y-%m-%d") + " 00:00:00"
        end_w3 = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=27- day)).strftime("%Y-%m-%d") + " 23:00:00"
        w3 = df[start_w3:end_w3]['Price']

        new_df = pd.concat([w0,w1,w2,w3]).to_frame()
        new_df.columns= ['Price']
        new_df.assign(forecasted=datetime.strptime(k[11:21], '%Y-%m-%d'))
        new_df['forecasted'] = [datetime.strptime(k[11:21], '%Y-%m-%d')] * len(new_df['Price'])

        lists[i].append(new_df)


germany = pd.concat(germany)
belgium = pd.concat(belgium)
france = pd.concat(france)

with pd.ExcelWriter(r'weekahead_hourly_DE_BE_FR.xlsx') as writer:  
    germany.to_excel(writer, sheet_name='germany')
    belgium.to_excel(writer, sheet_name='belgium')
    france.to_excel(writer, sheet_name='france')
    
