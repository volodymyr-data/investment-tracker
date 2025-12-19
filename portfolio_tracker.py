import numpy as np
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from typing import Union
import openpyxl
from pathlib import Path

#current problem: updating the most recent price

def main():
    FILENAME = Path("investment-tracker/stock_tracker.xlsx")
    if FILENAME.exists():
        print("entering the excel file...")
    else:
        create_excel()

    action = take_action()
    if action == "buy":
        ticker, start_date, end_date, shares_number = add_ticker()
        prices = load_prices(ticker, start_date, end_date)
        purchase_price = prices.iloc[0, 0]
        current_price = prices.iloc[-1, 0] #simulating the last date
        if ticker_owned(FILENAME, ticker):
            portfolio = pd.read_excel(FILENAME, index_col = "Ticker")
            start_price, num_owned = weighted_average(FILENAME, ticker, shares_number, purchase_price) 
            portfolio.at[ticker, "Number"] = num_owned
            portfolio.at[ticker, "Purchase price"] = start_price    
            portfolio.at[ticker, "% change"] = percent_change(start_price, current_price)    
            # start price is the weighted average of prices and num_owned is number u owned + number bought
            print(portfolio)
            update_to_excel(portfolio, FILENAME)

        else:
            data = format_for_excel(ticker, shares_number, purchase_price, current_price, percent_change(purchase_price, current_price))
            export_to_excel(data, FILENAME)

    elif action == "sell":
        ticker, shares_sold = delete_ticker()
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
    calculates the summary statistics of the portfolio
    
    :return: returns the number of holdings owned, total sum invested, total shares owned, average price per share, overall percent growth
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
    """
    removes shares when sold, only counts the shares
    
    :param FILENAME: const. file name to be accessed
    :type FILENAME: str
    :param ticker: ticker code
    :type ticker: str
    :param shares_sold: number of shares sold
    :type shares_sold: int
    :return: new number of shares owned after the sale
    :rtype: int
    """
    data = pd.read_excel(FILENAME, index_col = "Ticker")
    num_owned = data.loc[ticker]["Number"]
    num_owned -= shares_sold
    return num_owned

def weighted_average(FILENAME: str, ticker: str, shares_number: int, purchase_price: float) -> tuple[float, int]:
    """
    calculates the weighted average of shares bought at different prices
    
    :param FILENAME: const. file name to be accessed
    :type FILENAME: str
    :param ticker: ticker code
    :type ticker: str
    :param shares_number: number of shares owned
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

    :param purchase_price: price at which the share was purchased
    :type purchase_price: float
    :param current_price: current price of a share
    :type current_price: float
    :return: returns the percent change between the prices
    :rtype: float
    """
    percent_change = ((current_price - purchase_price) / purchase_price) * 100
    return percent_change


## USER CHOICES AND IMPLICATIONS
def take_action() -> str:
    """
    choose from the actions buy/sell/update
    
    :return: action chosen by user
    :rtype: str
    """

    action = input("Do you want to buy, sell, or update? ")
    return action

def delete_ticker() -> tuple[str, str, int]:
    """
    removes the shares of the ticker
    
    :return: ticker from transaction and number sold
    :rtype: tuple[str, str, int]
    """
    ticker = input("Ticker name: ")

    shares_number = int(input("Number sold: "))

    return ticker, shares_number

def add_ticker() -> tuple[str, str, str, int]:
    """
    prompts a user to enter their transaction for buying shares
    
    :return: ticker, date when shares were bought, most recent date when market was operating, number of shares bought
    :rtype: tuple[str, str, str, int]
    """
    date = input("enter the date of purchase (YYYY-MM-DD): ")
    start_date = pd.to_datetime(date)   
    end_date = datetime.today() - timedelta(days=1)    
    ticker = input("Ticker name: ")

    shares_number = int(input("Number bought: "))

    return ticker, start_date, end_date, shares_number

## EXCEL OPERATIONS

def update_summary(summary: pd.DataFrame, FILENAME: str) -> None:
    """
    updates the summary in excel sheet
    
    :param summary: formatted consolidation of changes in investments 
    :type summary: pd.DataFrame
    :param FILENAME: const. file name to be accessed
    :type FILENAME: str
    """
    writer = pd.ExcelWriter(FILENAME, engine = "openpyxl", mode = "a", if_sheet_exists = "replace")
    summary.to_excel(writer, sheet_name = "Summary", index = True)
    writer.close()

    print("Your Summary has been updated")

def format_summary(num_holdings: int, total_sum: float, total_shares: int, average_price: float, overall_percent: float) -> None:
    """
    formats summary for excel sheet
    
    :param num_holdings: number of holdings owned
    :type num_holdings: int
    :param total_sum: total capital invested
    :type total_sum: float
    :param total_shares: total shares owned
    :type total_shares: int
    :param average_price: average price of share
    :type average_price: float
    :param overall_percent: overall percentage return on portfolio
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

    :param ticker: ticker of a company
    :type ticker: str
    :param shares_number: number of shares owned
    :type shares_number: int
    :param purchase_price: purchase price of a share
    :type purchase_price: float
    :param end_price: most recent price of a share
    :type end_price: float
    :param percent_change: percent change from the start_price till the end_price
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
    
    :param data: consolited info of a transaction
    :type data: pd.DataFrame
    :param FILENAME: const. file name to be accessed
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
    :param FILENAME: const. file name to be accessed
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
    
    :param filename: const. file name to be accessed
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

def ticker_owned(filename: str, ticker:str) -> bool:
    """
    returns true if the ticker entered is already owned

    :param filename: const. file name to be accessed
    :type filename: str
    :param ticker: ticker of the company
    :type ticker: str
    :return: returns True if ticker is in excel file, and False if not
    :rtype: bool
    """

    tickers = pd.read_excel(filename)
    rows_list = tickers.values.tolist()
    for row in rows_list:
        if ticker in row:
            return True
    return False


if __name__ == "__main__":
    main()