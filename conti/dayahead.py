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

def parse_arguments():
    parser = argparse.ArgumentParser()
    
    parser.add_argument("-c","--countries",dest="countries",
            help="Countries", required=True)
    parser.add_argument("-s","--startday",dest="startday",
            help="Start day in the form YYYY-MM-DD", required=False)
    parser.add_argument("-e","--endday",dest="endday",
            help="End day in the form YYYY-MM-DD", required=False)
    parser.add_argument("-f","--filename",dest="filename",
        help="Choose between block, state, wigos and ship", required=True)
    args = parser.parse_args()
    
    if args.startday is None:
        pass
    else:
        try:
            datetime.strptime(args.startday,'%Y-%m-%d')
        except ValueError:
            raise ValueError
        
    if args.endday is None:
            pass
    else:
        try:
            datetime.strptime(args.endday,'%Y-%m-%d')
        except ValueError:
            raise ValueError
        
    if args.countries is None:
        parser.print_help()
        parser.exit()

    return args

get_args = parse_arguments()

print(get_args)

def nordic_or_conti():

    countries = get_args.countries.split(',')
    nordics = ['NO', 'NO1', 'NO2', 'NO3', 'NO4', 'NO5', 'SE', 'SE1', 'SE2', 'SE3', 'SE4', 'DK', 'DK1', 'DK2', 'FI']
    checker = 0
    for i in countries:
        if i in nordics:
            checker += 1

    if checker == 0:
        path = r'M:/EUROPA/Power Balance/Conti Model/History/Different runs within day'
    else:   
        path = r'M:/Database/NordicAreaPriceModel/prices_forecasts/napm_dayahead.csv'

    return path


#print(os.listdir(r'M:/EUROPA/Power Balance/Conti Model/History/Different runs within day'))

def conti(path, startdate, enddate):
    files = os.listdir(path)
    start_date = startdate
    end_date = enddate

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
            last_file_in_date.append(date_files_sorted[0])  # change here if you want to get the first or last file of date

    countries = get_args.countries.replace(" ", "").split(',')
    n = len(countries)
    lists = [[] for _ in range(len(countries))]



    for k in last_file_in_date:
        for i in range(0,n):

            df = pd.read_excel(path + '/' + k, index_col = 0, sheet_name=('PowerBalance-' + countries[i]))
            get_start = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=1)).strftime("%Y-%m-%d") + " 00:00:00"
            get_end = (datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=1)).strftime("%Y-%m-%d") + " 23:00:00"
            

            df_sliced = df[get_start:get_end]['Price'].to_frame()
            df_sliced['forecasted'] = datetime.strptime(k[11:27], '%Y-%m-%d T%H%M')
            #print(datetime.strptime(k[11:27], '%Y-%m-%d T%H%M'))
            df_mean = df_sliced['Price']

            #new_df = pd.DataFrame.from_dict({'Date': [(datetime.strptime(k[11:21], '%Y-%m-%d') + timedelta(days=1)).strftime("%Y-%m-%d")], 'Price': [df_sliced['Price']], 'Forecasted': [(datetime.strptime(k[11:21], '%Y-%m-%d')).strftime("%Y-%m-%d")]})
            #new_df = new_df.set_index('Date')
            #print(new_df)
            #sys.exit()
            lists[i].append(df_sliced)


    concated = []
    for i in lists:
        conc = pd.concat(i)
        concated.append(conc)


    return concated

def nordic(path, startdate, enddate):
    start_date = startdate
    end_date = enddate

    df = pd.read_csv(path, index_col = 0)
    df = df[startdate:enddate]

    countries = get_args.countries.replace(" ", "").split(',')
    countries.append('ForecastDate')
    if 'NO' not in countries or 'SE' not in countries or 'DK' not in countries:
        df = df[countries]
    else:
        print('You must specify area such as NO1, NO2, etc..')

    return df

#print(nordic(nordic_or_conti(), parse_arguments().startday, parse_arguments().endday))

def creating_file_conti(data):
    filename = get_args.filename
    countries = get_args.countries.replace(" ", "").split(',')

    with pd.ExcelWriter((filename + '.xlsx')) as writer:  
        for i in range(len(data)):
            data[i].to_excel(writer, sheet_name = ('DayAhead - ' + countries[i]))                


def creating_file_nordics(data):
    filename = get_args.filename
    countries = get_args.countries
    with pd.ExcelWriter((filename + '.xlsx')) as writer:
        data.to_excel(writer, sheet_name = countries)


data = conti(nordic_or_conti(), parse_arguments().startday, parse_arguments().endday)
creating_file_conti(data)

#data = nordic(nordic_or_conti(), parse_arguments().startday, parse_arguments().endday)
#creating_file_nordics(data)
