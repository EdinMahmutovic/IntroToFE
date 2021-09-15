# Intro to Financial Engineering
# Exercises week 3

import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
import yfinance as yf
import datetime
import openpyxl


print('\nIntro to Financial Engineering')
print('Exercises week 3\n')

today = "2021-09-12"
payday = "2021-11-15"
    
########### Exercise 1 ###########
print('-----------------------------------------------------------------------')
print('Creating Data sheet in Yields.xlsx')
print('-----------------------------------------------------------------------\n')
"""
Here i take the data from OblData.xlsx and finds the Dirty Prices, Maturities and Payoff matrisies
from there i do the work in Excel due to it being more standard in the real world, and it just being
easier. Thus this also shows how python can be a powerfull tool to use as assistance to excel
"""
OblData = pd.read_excel("C:\\Users\\Jonas\\OneDrive\\UNI\\IFE\\OblData.xlsx",engine="openpyxl", index_col=0, header=0)
OblData = OblData[["Kupon","Åbningskurs","Udløbsdato"]]
OblData["Maturity"] = (pd.DatetimeIndex(OblData["Udløbsdato"]) - pd.Timestamp(today))/np.timedelta64(1,"Y")#datetime.timedelta(days=365)
OblData[((pd.to_datetime(payday) - pd.Timestamp(today))/datetime.timedelta(days=365))] = OblData["Kupon"]#*(pd.to_datetime(payday) - pd.Timestamp(today))/datetime.timedelta(days=365)
OblData[((pd.to_datetime(payday) - pd.Timestamp(today))/datetime.timedelta(days=365))] += (pd.DatetimeIndex(OblData["Udløbsdato"]).year == 2021) * 100#OblData["Åbningskurs"]
OblData["Åbningskurs"] += (1-OblData["Maturity"]%1)*OblData["Kupon"]
OblData = OblData.rename(columns={"Åbningskurs": "Dirty Price"})
for yr in range(2022,2053):
  alive = pd.DatetimeIndex(OblData["Udløbsdato"]).year >= yr
  OblData[(yr-2021+(pd.to_datetime(payday) - pd.Timestamp(today))/datetime.timedelta(days=365))] = OblData["Kupon"]*alive
  OblData[(yr-2021+(pd.to_datetime(payday) - pd.Timestamp(today))/datetime.timedelta(days=365))] += (pd.DatetimeIndex(OblData["Udløbsdato"]).year == yr) *100 #OblData["Åbningskurs"]
OblData = OblData.sort_values(by=["Maturity"])



excelBook = openpyxl.load_workbook('C:\\Users\\Jonas\\OneDrive\\UNI\\IFE\\Yields.xlsx')
with pd.ExcelWriter('C:\\Users\\Jonas\\OneDrive\\UNI\\IFE\\Yields.xlsx', engine="openpyxl") as writer:
    writer.book = excelBook
    writer.sheets = dict((ws.title, ws) for ws in excelBook.worksheets)
    OblData.to_excel(writer, sheet_name="Data")
    writer.save()
