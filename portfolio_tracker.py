import numpy as np
import pandas as pd
import yfinance as yf
from datetime import datetime
from typing import Union
import openpyxl
from pathlib import Path

#current problem: deleting transactions

def main():
    filename = Path("stock_tracker.xlsx")
    if filename.is_file():
        print("entering the excel file...")
    else:
        create_excel()

    if take_action() == "buy":
        ticker, start_date, end_date, shares_number = add_ticker()
        prices = load_prices(ticker, start_date, end_date)
        purchase_price = prices.loc[start_date].iloc[0]
        current_price = prices.loc["2025-12-11"].iloc[0] #simulating the last date
        if ticker_owned(filename, ticker):
            portfolio = pd.read_excel(filename, index_col = "Ticker")
            start_price, num_owned = weighted_average(filename, ticker, shares_number, purchase_price) 
            portfolio.at[ticker, "Number"] = num_owned
            portfolio.at[ticker, "Purchase price"] = start_price    
            portfolio.at[ticker, "% change"] = percent_change(start_price, current_price)    
            # start price is the weighted average of prices and num_owned is number u owned + number bought
            ## data = format_for_excel(ticker, num_owned, start_price, current_price, percent_change(purchase_price, current_price))
            print(portfolio)
            update_to_excel(portfolio, filename)

        else:
            data = format_for_excel(ticker, shares_number, purchase_price, current_price, percent_change(purchase_price, current_price))
            export_to_excel(data, filename)

    elif take_action() == "sell":
        ticker, sale_date, shares_sold = delete_ticker()
        if ticker_owned(filename, ticker):
            portfolio = pd.read_excel(filename, index_col= "Ticker")
            num_owned = remove_shares(filename, ticker, shares_sold)
            portfolio.at[ticker, "Number"] = num_owned
            update_to_excel(portfolio, filename)
        else:
            print("you don't own that ticker and you cannot short")
        
def ticker_owned(file: str, ticker:str) -> bool:
    """
    returns true if the ticker entered is already owned

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

## CALCULATIONS
def remove_shares(filename: str, ticker: str, shares_sold: int) -> int:
    data = pd.read_excel(filename, index_col = "Ticker")
    num_owned = data.loc[ticker]["Number"]
    num_owned -= shares_sold
    return num_owned

def weighted_average(filename: str, ticker: str, shares_number: int, purchase_price: float) -> tuple[float, int]:
    """
    calculates the weighted average of shares bought at different prices
    
    :param filename: Description
    :type filename: str
    :param ticker: Description
    :type ticker: str
    :param shares_number: Description
    :type shares_number: int
    :param purchase_price: new "purchase price" - reweighted average of the stock price and the new shares number
    :type purchase_price: float
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
    calculated percent change from the price bought at (or weighted) till the most recent price

    :param purchase_price: Description
    :type purchase_price: float
    :param current_price: Description
    :type current_price: float
    :return: returns the percent change between the prices
    :rtype: float
    """
    percent_change = ((current_price - purchase_price) / purchase_price) * 100
    return percent_change

# def calculate_pl()-> float:

## USER CHOICES AND IMPLICATIONS
def take_action() -> str:
    """
    Docstring for take_action
    
    :return: Description
    :rtype: str
    """

    action = input("Do you want to buy or sell? ")
    return action

def delete_ticker() -> tuple[str, str, int]:
    """
    Docstring for delete_ticker
    
    :return: Description
    :rtype: tuple[str, str, int]
    """
    date = input("enter the date of sale (YYYY-MM-DD): ")
    sale_date = pd.to_datetime(date).strftime("%Y-%m-%d")   

    ticker = input("Ticker name: ")

    shares_number = int(input("Number bought: "))

    return ticker, sale_date, shares_number

def add_ticker() -> tuple[str, str, str, int]:
    """
    prompts a user to enter their transaction
    
    :return: Description
    :rtype: tuple[str, str, str, int]
    """
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
    loads the price of a ticker and deletes everything except for ticker and closing price for the day

    :param ticker: Description
    :type ticker: str
    :param start: Description
    :type start: Union[str, datetime, pd.Timestamp]
    :param end: Description
    :type end: Union[str, datetime, pd.Timestamp]
    :return: a dataframe with nothing but ticker name and closing price
    :rtype: DataFrame
    """
    data = yf.download(
        tickers = ticker, 
        start = start, 
        end = end, 
        interval = "1d"
    )
    data.drop(["High", "Low", "Open", "Volume"], inplace= True, axis = 1)
    return data

## EXCEL OPERATIONS

def format_for_excel(ticker: str, shares_number: int, purchase_price: float, end_price: float, percent_change: float) -> pd.DataFrame:
    """
    formats an entry for excel sheet

    :param ticker: Description
    :type ticker: str
    :param shares_number: Description
    :type shares_number: int
    :param purchase_price: Description
    :type purchase_price: float
    :param end_price: Description
    :type end_price: float
    :param percent_change: Description
    :type percent_change: float
    :return: a dataframe of a transaction with ticker, number of shares, buying price, recent price, and new percent change
    :rtype: DataFrame
    """
    form = {"Ticker": [ticker], "Number": [shares_number], "Purchase price": [purchase_price], 
            "End price": [end_price], "% change": [percent_change]}
    data = pd.DataFrame(data= form, index = [ticker])
    # print(data)
    return data

def update_remove(data: pd.DataFrame, filename: str) -> None:
    writer = pd.ExcelWriter(filename, engine = "openpyxl", mode = "w")
    data.to_excel(writer, index = True)
    writer.close()
    
    print("stock was sold")

def update_to_excel(data: pd.DataFrame, filename: str) -> None:
    """
    when the ticker entered is owned the function will write down the updated info
    
    :param data: Description
    :type data: pd.DataFrame
    :param filename: Description
    :type filename: str
    """
    writer = pd.ExcelWriter(path= filename, engine= "openpyxl", mode = "w")
    data.to_excel(writer, index = True)
    writer.close()
    print("the stocks owned was updated!")

def export_to_excel(data: pd.DataFrame, filename: str) -> None:
    """
    exports all the data to excel
        
    :param data: Description
    :type data: pd.DataFrame
    :param filename: Description
    :type filename: str
    """

    reader = pd.read_excel(filename)
    writer = pd.ExcelWriter(path = filename, engine = "openpyxl", mode = "a", if_sheet_exists = "overlay")
    data.to_excel(writer, index = False, header = False, startrow = len(reader) + 1)
    writer.close()
    print("new stock has been added")

def create_excel() -> None:
    """
    if there is no excel file for tracking in the current directory this will create one
    """

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