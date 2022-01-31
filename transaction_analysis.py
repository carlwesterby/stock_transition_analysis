import pandas as pd
%matplotlib notebook
import matplotlib.pyplot as plt
import numpy as np
import glob
import os.path


import yfinance as yf
import time
import datetime
from datetime import datetime

import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import string
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
import xlsxwriter
#####
#Note that the file must have column headings of: "Date",	"Action",	"Symbol",	"Quantity", and "Amount"

file=r"C:\YOURPATH\YOURFILE.CSV"
transactionsDF=pd.read_csv(file, index_col="Date")
#convert the strings to datetime
transactionsDF.index=pd.to_datetime(transactionsDF.index, format='%m/%d/%Y')
#convert the $ amounts from strings to floats
transactionsDF['Price']=transactionsDF['Price'].str.replace(r'$', '').str.replace(r',', '').astype(float)
transactionsDF['Amount']=transactionsDF['Amount'].str.replace(r'$', '').str.replace(r',', '').str.replace(r')', '').str.replace(r'(', '').astype(float)
#Get all the relavent stock symbols
temp_symbols=transactionsDF["Symbol"].values
tickers=[]
for tick in temp_symbols:
    if ( (tick not in tickers) and (not pd.isnull(tick))):
        tickers.append(tick)
portfolio_tickers=sorted(tickers)
ticker_columns = portfolio_tickers
#add the S&P500 so we can compare aginst it
ticker_columns.append('SPY')
stocks = yf.download(tickers = ticker_columns, start="2021-7-1", interval='1d', threads= False)
stock_data=stocks['Adj Close']
stock_data=stock_data.dropna()
#add in a column for tracking cash
tracker_columns=ticker_columns+['Cash']
#DF for tracking total # of share
Shares=pd.DataFrame(columns = tracker_columns, index=['Shares'])
#force to 0
Shares.loc['Shares'] = np.zeros(len(tracker_columns), dtype=int)
#DF for tracking cost basis
CostBasis=pd.DataFrame(columns = tracker_columns, index=['Cost'])
CostBasis.loc['Cost'] = np.zeros(len(tracker_columns))
TransTotalDF=pd.DataFrame(columns = ['Transactions'])

#data fram for tracking number of share over time
ShareDF= pd.DataFrame(columns = tracker_columns, index = stock_data.index)

#loop thorugh each date in the shares over time DF
for date in ShareDF.index:
    date_transactions=0
    #grab the transations on this date
    tempDF=transactionsDF[transactionsDF.index==date]
    #if there are none then keep the current share status
    if(len(tempDF)==0):
        ShareDF.loc[date]=Shares.loc['Shares']
        #print('0')
    #other wise update cost basis and share count
    else:
        #print('1')
        for index, transaction in tempDF.iterrows():
            #temp variables to simply code
            temp_tick=transaction['Symbol']
            temp_share=transaction['Quantity']
            temp_price=transaction['Price']
            
            if(transaction['Action']=='Buy'):
                #increase the # of shares of the correct company
                Shares.at['Shares',temp_tick]=Shares.at['Shares',temp_tick]+temp_share
                #increase cost basis based on price of the stock bought
                CostBasis.at['Cost',temp_tick]=CostBasis.at['Cost',temp_tick]+(temp_share*temp_price)
                #if we have more cash on hand than the cost of the purchase
                if(Shares.at['Shares','Cash']>=(temp_share*temp_price)):
                    #decrease cash position, so that sales of other stocks are reinvested
                    Shares.at['Shares','Cash']=Shares.at['Shares','Cash']-(temp_share*temp_price)
                    #no need to purchase more S&P500
                else:
                    #decrease cash position, so that sales of other stocks are reinvested
                    date_transactions=date_transactions+(temp_share*temp_price)-Shares.at['Shares','Cash']
                    #buy SP&P with the transation cost above cash on hand
                    Shares.at['Shares','SPY'] = Shares.at['Shares','SPY']+(((temp_share*temp_price)-Shares.at['Shares','Cash']) /stock_data.at[date,'SPY'])
                    #clear out cash position
                    Shares.at['Shares','Cash']=0
                #print(date, ",\t",temp_tick,': \t\t',stock_data.at[date,'SPY'], ',\t ',temp_share*temp_price,'\t',Shares.at['Shares','SPY'])
            elif(transaction['Action']=='Sell'):
                #decrease the number of shares of the ocmpany
                Shares.at['Shares',temp_tick]=Shares.at['Shares',temp_tick]-temp_share
                #update cost basis by subtracting the sale price
                CostBasis.at['Cost',temp_tick]=CostBasis.at['Cost',temp_tick]-(temp_share*temp_price)
                #update cash position based on sale price
                Shares.at['Shares','Cash']=Shares.at['Shares','Cash']+(temp_share*temp_price)
            elif("Div" in transaction['Action'] ):
                #add dividends into the cash position
                Shares.at['Shares','Cash']=Shares.at['Shares','Cash']+transaction['Amount']
        #push the share status to the current date
        ShareDF.loc[date]=Shares.loc['Shares']
    TransTotalDF.at[date,'Transactions']=date_transactions
#multiply shares times stock prices
ValueDF=stock_data.multiply(ShareDF)
ValueDF['Cash']=ShareDF['Cash']
#Get all the relavent stock symbols
temp_symbols=transactionsDF["Symbol"].values
tickers=[]
for tick in temp_symbols:
    if ( (tick not in tickers) and (not pd.isnull(tick))):
        tickers.append(tick)
portfolio_tickers=sorted(tickers)
print(portfolio_tickers)
#add together all positions in the portfolio
PortfolioDF=pd.DataFrame(columns=['Total'],index=stock_data.index)
PortfolioDF=PortfolioDF.replace(np.nan, 0)
temp_tickers=portfolio_tickers+['Cash']
for name in temp_tickers:
    PortfolioDF=PortfolioDF.add(ValueDF[name].values,axis='index')
plt.figure()
plt.plot(PortfolioDF.index,PortfolioDF['Total'])
plt.plot(ValueDF.index, ValueDF['SPY'])
plt.legend({"Your Portfolio",'S&P 500 Portfolio'})

current_day=PortfolioDF['Total'].values
temptransactions=TransTotalDF['Transactions'].values
#roll to create the previous day
prev_day=np.roll((current_day), 1)
#pct gains are today's value-stocks purchased outside cash on hand - yesterday / yesterday
pctvalues=(current_day-prev_day-temptransactions)/prev_day
#first value has no value to compare against
pctvalues[0]=0

current_SP=ValueDF['SPY'].values
temptransactions=TransTotalDF['Transactions'].values
#pct gains are today's value-stocks purchased outside cash on hand - yesterday / yesterday
prev_SP=np.roll((current_SP), 1)
pctSP=(current_SP-prev_SP-temptransactions)/prev_SP
pctSP[0]=0

plt.figure()
plt.plot(PortfolioDF.index,pctvalues)
plt.plot(ValueDF.index,pctSP)
plt.legend({"Your Portfolio",'S&P 500 Portfolio'})

#add 1 to covert from percent change, to todays multiple of yesterday
pctprev=pctvalues+1
portfolio_pct=[]
SP_portfolio_pct=[]
temp=1
#geometric summation of pct changes
for value in pctprev:
    portfolio_pct.append(value*temp)
    temp=temp*value

temp=1
for value in (pctSP+1):
    SP_portfolio_pct.append(value*temp)
    temp=temp*value
    
plt.figure()
plt.plot(ValueDF.index,portfolio_pct)
plt.plot(ValueDF.index, SP_portfolio_pct)
plt.legend({"Your Portfolio",'S&P 500 Portfolio'})
#push portfolio summary in time to a single DF
SummaryDF=PortfolioDF
SummaryDF['SPY']=ValueDF['SPY']
SummaryDF['Your Percent']=portfolio_pct
SummaryDF['SPY Percent']=SP_portfolio_pct

#combine current # of shares and cost basis
Status = pd.concat([Shares,CostBasis])

#push to excel document
with pd.ExcelWriter(r'C:\Users\cwesterb\Stocks\Transations\Portfolio_Summary.xlsx', engine='xlsxwriter') as writer:  
    stock_data.to_excel(writer, sheet_name='Stock_prices')
    transactionsDF.to_excel(writer, sheet_name='Transactions')
    ValueDF.to_excel(writer, sheet_name='Positions_in_time')
    ShareDF.to_excel(writer, sheet_name='Shares_in_time')
    SummaryDF.to_excel(writer, sheet_name='Summary_in_time')
    Status.to_excel(writer, sheet_name='Posrtfolio Sumary')
