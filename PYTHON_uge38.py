# Intro to Financial Engineering
# Exercises week 3

import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
import yfinance as yf
import datetime
import openpyxl
from nelson_siegel_svensson.calibrate import calibrate_ns_ols



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
input(OblData)
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

YtM = OblData[["Udløbsdato"]]

dates = OblData.columns.values[4:]
## make YtM

from scipy.optimize import fsolve, minimize
def func(x):
    out =  sum( OblData[date]/(1+x)**date for date in dates )
    return out-OblData["Dirty Price"]
root = fsolve(func, [0]*len(OblData["Dirty Price"]))
YtM["YtM"] = root
today = pd.Timestamp(today)
YtM["T"] = (OblData["Udløbsdato"]-today)/datetime.timedelta(days=1)
curve, s = calibrate_ns_ols(np.array(YtM["T"]), np.array(YtM["YtM"]), tau0=1.0)
beta = [float(num) for num in str(curve).replace("=",",").replace(")",",").split(",") if num.replace("-","").replace(".","").isnumeric()]
YtM["NSS beta"] = beta+(len(YtM["T"])-len(beta))*["NaN"]
#YtM["beta"] = list(theta)+(len(YtM["YtM"])-len(theta))*["NaN"]


NS = YtM
NS = NS.drop(["NSS beta"],axis=1)
theta = [0,0,0,0]
#today = pd.Timestamp(today)
#NS["T"] = (OblData["Udløbsdato"]-today)/datetime.timedelta(days=1)
NS["theta"] = list(theta)+(len(NS["T"])-len(theta))*["NaN"]

def func(theta):
    NS["NS"] = theta[0] + (theta[1]+theta[2]/theta[3])*(1-np.exp(-theta[3]*NS["T"]))/theta[3]/NS["T"]-theta[2]/theta[3]*np.exp(-theta[3]*NS["T"])
    return sum((NS["NS"]*1000-NS["YtM"]*1000)**2)
theta = minimize(func, theta, method='nelder-mead').x
NS["theta"] = list(theta)+(len(NS["T"])-len(theta))*["NaN"]

Duration = OblData[["Kupon","Dirty Price","Maturity"]]
Duration["YtM"] = YtM["YtM"]
Duration["Macaulay"] = sum(date*OblData[date]/(1+YtM["YtM"])**date for date in dates)/OblData["Dirty Price"]
Duration["Modified"] = Duration["Macaulay"]/(1+Duration["YtM"])
def NSnow(t): t = t*365; return theta[0] + (theta[1]+theta[2]/theta[3])*(1-np.exp(-theta[3]*t))/theta[3]/t-theta[2]/theta[3]*np.exp(-theta[3]*t)
Duration["Fisher-Weil"] = sum(date*OblData[date]/(1+NSnow(date))**(date+1) for date in dates)/OblData["Dirty Price"]
Duration["MOD_DIFF"] = Duration["Modified"]-Duration["Macaulay"]
Duration["FW_DIFF"]  = Duration["Fisher-Weil"]-Duration["Macaulay"]
Duration["Convexity"]= sum(date*(date+1)*OblData[date]/(1+NSnow(date))**(date+2) for date in dates)/OblData["Dirty Price"]
Duration["Apprx. relative change In Price"] = Duration["Modified"]*(Duration["YtM"].diff())+0.5*Duration["Convexity"]*(Duration["YtM"].diff())**2





excelBook = openpyxl.load_workbook('C:\\Users\\Jonas\\OneDrive\\UNI\\IFE\\Yields.xlsx')
with pd.ExcelWriter('C:\\Users\\Jonas\\OneDrive\\UNI\\IFE\\Yields.xlsx', engine="openpyxl") as writer:
    writer.book = excelBook
    writer.sheets = dict((ws.title, ws) for ws in excelBook.worksheets)
    OblData.to_excel(writer, sheet_name="Data")
    YtM.to_excel(writer, sheet_name="YtM")
    NS.to_excel(writer,sheet_name="NS")
    Duration.to_excel(writer,sheet_name="Duration")
    writer.save()
