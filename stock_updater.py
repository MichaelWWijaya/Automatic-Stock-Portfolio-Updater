import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import os
import string

#Section 0: Ask for information about the filename and cells
print("\nWelcome to Michael's Automatic Stock Portfolio Updater!\nPlease input your information below\n")
# fileN = input("Filename (without .xlsx): ")
# fileN += ".xlsx"
# month = input("Please input the month sheet you are working on (ex: Apr) : ")
# path_file = input(r"Please input your file path of the folder containing your portfolio file (C:\Users\etc): ")
# start_Cell = input("Please input the first cell of your Ticker from the Stock price data (ex: B84): ")
# end_Cell = input("Please input the last cell of your Ticker from the Stock price data (ex: B95): ")
# price_Cell = input("Please input the first cell of your Price column from the Stock price data (ex: C84): ")
close = input("\nPlease close your portfolio file before running this program (Click ENTER to start)")

# or you can save all the information without needing to input again 
# fileN = "2026 USA Portfolio Michael"
fileN = "2026 USA Portfolio Michael"
fileN += ".xlsx"
month = "Oct"
path_file = r"C:\Users\pingl\Documents\Michael\a PERSONAL PROJECT CS\Automatic Stock Updater"
# start_Cell = "B86"
# end_Cell = "B97"
# price_Cell = "C86"

# iterations = int(end_Cell[1:]) - int(start_Cell[1:]) + 1
# Ticker_cell_letter = start_Cell[:1]
# Price_cell_letter = price_Cell[:1]


#Section 1: Define Tickers and Time Range

## Define the list of tickers
from openpyxl import load_workbook

workbook = load_workbook(filename=fileN, data_only=True)
sheet = workbook[month]


# Find the cell that contains the String "Stock price data"
stock_price_data = "A0"
for i in range(1000):
    cell_ref = "A"+str(i+1)
    if(sheet[cell_ref].value == "Stock price data"):
        stock_price_data = cell_ref 
        # example stock_price_data = A2

# Useful variables to determine the raneg of data

start_letter = stock_price_data[:1]

index = 1
# ord('a') evaluates to 97, so 97 + 1 = 98
ticker_letter = str(chr(ord(start_letter) + index))


# print(letter_ticker)  # Output: B
ticker_number = int(stock_price_data[1:]) + 3 #adjusted by 3 rows

# print("the start letter is: "+ticker_letter)
# print("the cell number: "+str(ticker_number))

start_Cell = ticker_letter+str(ticker_number)
end_Cell = "B100"
price_Cell = str(chr(ord(start_letter) + 2)) + str(ticker_number)

# iterate throught tickers until no more stocks are left
for i in range(25):
    cell_ref = ticker_letter+str(ticker_number+i)
    print("the value is "+ str(sheet[cell_ref].value))
    if(str(sheet[cell_ref].value) == "None"):
        end_Cell = ticker_letter+str(ticker_number+i-1)
        break

print("\nthe start cell is "+start_Cell)
print("the end cell is "+end_Cell)
print("the price cell is "+price_Cell)
# print("the cell number: "+str(ticker_number))
# print("the cell number: "+str(ticker_number))


iterations = int(end_Cell[1:]) - int(start_Cell[1:]) + 1
Ticker_cell_letter = start_Cell[:1]
Price_cell_letter = price_Cell[:1]


tickers = []
for i in range(iterations):
    cell_ref = Ticker_cell_letter+str(int(start_Cell[1:])+i)
    tickers.append(sheet[cell_ref].value)

# tickers = ['AAPL','AMZN','MSFT','NVDA','NFLX']





## Set the end date to the last market day

# end_date = datetime.today()
# print("The date is "+str(end_date))


# Set the end date to the date in excel

date_cell = str(chr(ord(start_letter) + 3)) +str(int(stock_price_data[1:]) + 1)

print("\nThe date cell from excel is "+str(sheet[date_cell].value)+"\n")
end_date = sheet[date_cell].value + timedelta(days = 1)
# end_date = end_date + timedelta(days = 1)

print(end_date)

## Set the start date to 2 years ago
start_date = end_date - timedelta(days = 1)
print(start_date)



#Section 2: Download Close Prices

## Create an empty DataFrame to store the close prices
close_df = pd.DataFrame()

## Download the close prices for each ticker
for ticker in tickers:
    data = yf.download(ticker, start = start_date, end = end_date)
    close_df[ticker] = data['Close']

## Display the DataFrame
print(close_df)



# Section 3: Set Output Folder and Export Data

## Set the output folder path
output_folder = path_file
# output_folder = r"C:\Users\pingl\Documents\Michael\a PERSONAL PROJECT CS\Automatic Stock Updater"

## Export the DataFrame to Excel
output_file = os.path.join(output_folder, 'stock_prices.xlsx')
close_df.to_excel(output_file)



# now we update the closing price
 
from openpyxl import load_workbook
workbook = load_workbook(filename="stock_prices.xlsx")
sheet = workbook["Sheet1"]

# workbook2 = load_workbook(filename="2026 USA Portfolio Michael.xlsx")
# sheet2 = workbook2["Apr"]

workbook2 = load_workbook(filename=fileN)
sheet2 = workbook2[month]


for i in range(iterations):
    start_cell_num = int(start_Cell[1:])
    # ticker_cell_ref = "B"+str(start_cell_num+i)
    update_cell_ref = Price_cell_letter+str(start_cell_num+i)
    # tickers.append(sheet[cell_ref].value)
    
    letter = string.ascii_uppercase[1+i]
    reference_cell = letter+str(2)
    # sheet[ticker_cell_ref].value = tickers[i]
    # sheet[update_cell_ref].value = sheet[reference_cell].value
    sheet2[update_cell_ref].value = round(sheet[reference_cell].value,2)

workbook.save(filename="stock_prices.xlsx")
workbook2.save(filename=fileN)
# workbook2.save(filename="2026 USA Portfolio Michael.xlsx")
