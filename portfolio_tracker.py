import numpy as np
import pandas as pd
import yfinance as yf
from datetime import datetime
from typing import Union
import openpyxl
from pathlib import Path

#current problem: combining the tickers owned, deleting transactions
# now it just writes all the transactions

def main():
    filename = Path("stock_tracker.xlsx")
    if filename.is_file():
        print("entering the excel file...")
    else:
        create_excel()

    ticker, start_date, end_date, shares_number = user_choice()
    prices = load_prices(ticker, start_date, end_date)
    purchase_price = prices.loc[start_date].iloc[0]
    current_price = prices.loc["2025-12-11"].iloc[0] #simulating the last date
    if ticker_owned(filename, ticker):
        portfolio = pd.read_excel(filename, index_col = "Ticker")
        start_price, num_owned = weighted_average(filename, ticker, shares_number, purchase_price) 
        portfolio.at[ticker, "Number"] = num_owned
        portfolio.at[ticker, "Purchase price"] = start_price        
        # start price is the weighted average of prices and num_owned is number u owned + number bought
        ## data = format_for_excel(ticker, num_owned, start_price, current_price, percent_change(purchase_price, current_price))
        print(portfolio)
        update_to_excel(portfolio, filename)

    else:
        data = format_for_excel(ticker, shares_number, purchase_price, current_price, percent_change(purchase_price, current_price))
        export_to_excel(data, filename)

def ticker_owned(file: str, ticker:str) -> bool:
    """
    Docstring for ticker_owned
    
    :param file: Description
    :type file: str
    :param ticker: Description
    :type ticker: str
    :return: Description
    :rtype: bool
    """
    tickers = pd.read_excel(file)
    rows_list = tickers.values.tolist()
    for row in rows_list:
        if ticker in row:
            return True
    return False

def weighted_average(filename: str, ticker: str, shares_number: int, purchase_price: float) -> None:
    """
    Docstring for weighted_average
    
    :param filename: Description
    :param ticker: Description
    :param shares_number: Description
    :param purchase_price: Description
    :return: Description
    :rtype: float
    """
    # calculating new number owned and new weighted average
    data = pd.read_excel(filename, index_col= "Ticker")
    num_owned = data.loc[ticker]["Number"]
    start_price = data.loc[ticker]["Purchase price"]
    weighted_price = (num_owned * start_price + shares_number * purchase_price) / (num_owned + shares_number)
    # setting new start price and summing up the shares owned
    print("starting price at ", start_price)
    print("purchased at", purchase_price)
    start_price = weighted_price
    print("weighted price at ", start_price)
    num_owned += shares_number

    return start_price, num_owned

def percent_change(purchase_price:float, current_price: float) -> float:
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
    # print(data)
    return data

def update_to_excel(data: pd.DataFrame, filename: str) -> None:
    """
    Docstring for update_to_excel
    
    :param data: Description
    """
    writer = pd.ExcelWriter(path= filename, engine= "openpyxl", mode = "w")
    data.to_excel(writer, index = True)
    writer.close()
    print("the stocks owned was updated!")

def export_to_excel(data, filename: str) -> None:
    "exports the changes to a .xlsx file"
    reader = pd.read_excel(filename)
    writer = pd.ExcelWriter(path = filename, engine = "openpyxl", mode = "a", if_sheet_exists = "overlay")
    data.to_excel(writer, index = False, header = False, startrow = len(reader) + 1)
    writer.close()
    print("new stock has been added")

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
    print("new excel file has been created")

if __name__ == "__main__":
    main()