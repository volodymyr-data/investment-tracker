import numpy as np
import pandas as pd
import yfinance as yf
from datetime import datetime
from typing import Union
import openpyxl
from pathlib import Path

#current problem: deleting transactions

def main():
    ticker, start_date, end_date, shares_number = user_choice()
    prices = load_prices(ticker, start_date, end_date)
    purchase_price = prices.loc[start_date].iloc[0]
    current_price = prices.loc["2025-12-12"].iloc[0] #simulating the last date
    data = format_for_excel(ticker, shares_number, purchase_price, current_price, calculations(purchase_price, current_price))
    filename = Path("stock_tracker.xlsx")
    if filename.is_file():
        print("entering the excel file...")
    else:
        create_excel()
    
    export_to_excel(data)

def calculations(purchase_price, current_price) -> float:
    """
    does calculations for percent change 
    """
    percent_change = ((current_price - purchase_price) / purchase_price) * 100
    return percent_change

def user_choice():
    date = input("enter the date of purchase (YYYY-MM-DD): ")
    start_date = pd.to_datetime(date).strftime("%Y-%m-%d")   
    end_date = datetime.now().strftime("%Y-%m-%d")
    
    ticker = input("Ticker name: ")

    shares_number = int(input("Number bought: "))

    return ticker, start_date, end_date, shares_number

def load_prices(ticker: str, 
                start: Union[str, datetime, pd.Timestamp],
                end:Union[str, datetime, pd.Timestamp]) -> pd.DataFrame:
    """
    loads prices from certain date till today and return only date and closing price
    """
    data = yf.download(
        tickers = ticker, 
        start = start, 
        end = end, 
        interval = "1d"
    )
    data.drop(["High", "Low", "Open", "Volume"], inplace= True, axis = 1)
    return data

def format_for_excel(ticker, shares_number, purchase_price, end_price, percent_change) -> pd.DataFrame:
    """
    format a datafame for excel exporting
    """
    form = {"Ticker": [ticker], "Number": [shares_number], "Purchase price": [purchase_price], 
            "End price": [end_price], "% change": [percent_change]}
    data = pd.DataFrame(data= form, index = [ticker])
    print(data)
    return data

def export_to_excel(data) -> None:
    "exports the changes to a .xlsx file"
    file_name = "stock_tracker.xlsx"
    reader = pd.read_excel(file_name)
    writer = pd.ExcelWriter(path = file_name, engine = "openpyxl", mode = "a", if_sheet_exists = "overlay")
    data.to_excel(writer, index = False, header = False, startrow = len(reader) + 1)
    writer.close()
    print("success!")

def create_excel() -> None:
    """creates excel file"""
    print("creating excel file...")
    file_name = "stock_tracker.xlsx"
    data = {"Ticker": [],
            "Number": [], 
            "Purchase price": [], 
            "End price": [],
            "% change": []
            }
    data = pd.DataFrame(data).set_index("Ticker")
    data.to_excel(file_name)
    print("success!")

if __name__ == "__main__":
    main()