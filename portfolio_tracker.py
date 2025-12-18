import numpy as np
import pandas as pd
import yfinance as yf
from datetime import datetime
from typing import Union
import openpyxl
from pathlib import Path

#current problem: updating the most recent price

def main():
    FILENAME = Path("stock_tracker.xlsx")
    if FILENAME.is_file():
        print("entering the excel file...")
    else:
        create_excel()

    action = take_action()
    if action == "buy":
        ticker, start_date, end_date, shares_number = add_ticker()
        prices = load_prices(ticker, start_date, end_date)
        purchase_price = prices.loc[start_date].iloc[0]
        current_price = prices.loc["2025-12-11"].iloc[0] #simulating the last date
        if ticker_owned(FILENAME, ticker):
            portfolio = pd.read_excel(FILENAME, index_col = "Ticker")
            start_price, num_owned = weighted_average(FILENAME, ticker, shares_number, purchase_price) 
            portfolio.at[ticker, "Number"] = num_owned
            portfolio.at[ticker, "Purchase price"] = start_price    
            portfolio.at[ticker, "% change"] = percent_change(start_price, current_price)    
            # start price is the weighted average of prices and num_owned is number u owned + number bought
            ## data = format_for_excel(ticker, num_owned, start_price, current_price, percent_change(purchase_price, current_price))
            print(portfolio)
            update_to_excel(portfolio, FILENAME)

        else:
            data = format_for_excel(ticker, shares_number, purchase_price, current_price, percent_change(purchase_price, current_price))
            export_to_excel(data, FILENAME)

    elif action == "sell":
        ticker, sale_date, shares_sold = delete_ticker()
        if ticker_owned(FILENAME, ticker):
            portfolio = pd.read_excel(FILENAME, index_col= "Ticker")
            num_owned = remove_shares(FILENAME, ticker, shares_sold)
            portfolio.at[ticker, "Number"] = num_owned
            update_to_excel(portfolio, FILENAME)
        else:
            print("you don't own that ticker and you cannot short")
    
    elif action == "update":
        current_price = update_prices(FILENAME)


    num_holdings, total_sum, total_shares, average_price, overall_percent = prepare_summary(FILENAME)   
    summary_to_update = format_summary(num_holdings, total_sum, total_shares, average_price, overall_percent)
    update_summary(summary_to_update, FILENAME)

## CALCULATIONS
def prepare_summary(FILENAME: str) -> tuple[int, float, int, float, float]:
    """
    claculates the summary statistics of the portfolio
    
    :return: Description
    :rtype: tuple[int, float, int, float, float]
    """
    portfolio = pd.read_excel(FILENAME)
    num_holdings = len(portfolio.values.tolist())
    total_shares = portfolio["Number"].sum()
    total_sum = (portfolio["Number"] * portfolio["Purchase price"]).sum()
    average_price = total_sum / total_shares
    overall_percent = ((portfolio["End price"] - portfolio["Purchase price"]).sum() / portfolio["Purchase price"].sum()) * 100

    return num_holdings, total_sum, total_shares, average_price, overall_percent

def remove_shares(FILENAME: str, ticker: str, shares_sold: int) -> int:
    data = pd.read_excel(FILENAME, index_col = "Ticker")
    num_owned = data.loc[ticker]["Number"]
    num_owned -= shares_sold
    return num_owned

def weighted_average(FILENAME: str, ticker: str, shares_number: int, purchase_price: float) -> tuple[float, int]:
    """
    calculates the weighted average of shares bought at different prices
    
    :param FILENAME: Description
    :type FILENAME: str
    :param ticker: Description
    :type ticker: str
    :param shares_number: Description
    :type shares_number: int
    :param purchase_price: new "purchase price" - reweighted average of the stock price and the new shares number
    :type purchase_price: float
    """
    # calculating new number owned and new weighted average
    data = pd.read_excel(FILENAME, index_col= "Ticker")
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


## USER CHOICES AND IMPLICATIONS
def take_action() -> str:
    """
    Docstring for take_action
    
    :return: Description
    :rtype: str
    """

    action = input("Do you want to buy, sell, or update? ")
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

    shares_number = int(input("Number sold: "))

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

## EXCEL OPERATIONS

def update_summary(summary: pd.DataFrame, FILENAME: str) -> None:
    """
    updates the summary in excel sheet
    
    :param summary: Description
    :type summary: pd.DataFrame
    :param FILENAME: Description
    :type FILENAME: str
    """
    writer = pd.ExcelWriter(FILENAME, engine = "openpyxl", mode = "a", if_sheet_exists = "replace")
    summary.to_excel(writer, sheet_name = "Summary", index = True)
    writer.close()

    print("Your Summary has been updated")

def format_summary(num_holdings: int, total_sum: float, total_shares: int, average_price: float, overall_percent: float) -> None:
    """
    formats summary for excel sheet
    
    :param num_holdings: Description
    :type num_holdings: int
    :param total_sum: Description
    :type total_sum: float
    :param total_shares: Description
    :type total_shares: int
    :param average_price: Description
    :type average_price: float
    :param overall_percent: Description
    :type overall_percent: float
    """
    form = {
        "Holdings owned": [num_holdings], 
        "Total Capital invested": [total_sum], 
        "Total shares owned": [total_shares],
        "Average price of a share": average_price, 
        "Portfolio growth": [overall_percent]
    }
    data = pd.DataFrame(data = form)

    return data

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

def update_to_excel(data: pd.DataFrame, FILENAME: str) -> None:
    """
    when the ticker entered is owned the function will write down the updated info
    
    :param data: Description
    :type data: pd.DataFrame
    :param FILENAME: Description
    :type FILENAME: str
    """
    writer = pd.ExcelWriter(path= FILENAME, engine= "openpyxl", mode = "w")
    data.to_excel(writer, index = True, sheet_name = "MyInvestments")
    writer.close()
    print("the stocks owned was updated!")

def export_to_excel(data: pd.DataFrame, FILENAME: str) -> None:
    """
    exports all the new data to excel
        
    :param data: Description
    :type data: pd.DataFrame
    :param FILENAME: Description
    :type FILENAME: str
    """

    reader = pd.read_excel(FILENAME)
    writer = pd.ExcelWriter(path = FILENAME, engine = "openpyxl", mode = "a", if_sheet_exists = "overlay")
    data.to_excel(writer, index = False, header = False, startrow = len(reader) + 1, sheet_name = "MyInvestments")
    writer.close()
    print("new stock has been added")

def create_excel() -> None:
    """
    if there is no excel file for tracking in the current directory this will create one
    """

    print("creating excel file...")
    file_name = "stock_tracker.xlsx"
    investments = {"Ticker": [],
            "Number": [], 
            "Purchase price": [], 
            "End price": [],
            "% change": []
            }
    summary = {
        "Holdings owned": [], 
        "Total Capital invested": [], 
        "Total shares owned": [],
        "Average price of a share": [], 
        "Portfolio growth": []
    }
    writer = pd.ExcelWriter(file_name, engine="openpyxl")
    investments_df = pd.DataFrame(investments).set_index("Ticker")
    summary_df = pd.DataFrame(summary).set_index("Holdings owned")
    investments_df.to_excel(writer, sheet_name="MyInvestments")
    summary_df.to_excel(writer, sheet_name="Summary")
    writer.close()


    print("new excel file has been created")

## OTHER
def update_prices(filename: str) -> pd.DataFrame:
    """
    updates the prices on request and saves them in excel
    
    :param filename: Description
    :type filename: str
    :return: Description
    :rtype: DataFrame
    """
    df = pd.read_excel(filename)
    tickers = df["Ticker"].to_list()
    prices = yf.download(tickers, interval = "1d")["Close"]
    df["End price"] = df["Ticker"].apply(lambda x: prices[x.upper()].iloc[-2] if len(tickers) > 1 else prices.iloc[-2])
    # print(df["End price"])
    df.to_excel(filename, index=False)


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
if __name__ == "__main__":
    main()