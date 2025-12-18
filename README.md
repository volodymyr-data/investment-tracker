# Project name: investment tracker
## Purpose:
    script for automatization of tracking of investments (shares) of publicly traded companies. It allows the use to simply choose action (buy/sell/update) and record transaction by entering a ticker, date of purchase, and number of shares purchased. Later the program retreives this information, formats it appropriately, and records it into the Excel-sheet, and updates the summary of your investments.
## Data: 
    data source is YahooFinance API and contains the historical Open, Close, Highest, Lowest, prices and Volume from the date requested by a user up to most recent price. It is then cleaned up to contain only a ticker, and a close price
## Tech stack: 
    Python (Numpy, Pandas), API(YFinance), Excel spreadsheets
## Work flow: 
    User choice (buy/sell/update) -> retreive of relevant prices -> Data cleansing to contain only the parts to be used -> formatting the relevant data into the appropriate table for the excel file -> update of summary information -> update of corresponding excel sheets in excel file
## Insights: 
    * File handling: I managed to realize the logic that check whether the file exists and creates one with the right name and headers if it doesn't
    * Data aggregation: I managed to summarize all the current investments on a separate sheet. Thus will show your total shares owned, total holdings owned, total capital invested, total portfolio growth
    * Automatization: instead of entering every single piece of data (it would be too hard to find a price of a share at some particular date) a user will enter ticker, date, and number of shares, and the program will fill the rest by itself
## Future roadmap: 
    i plan to add more insights on investments, visualization of portfolio returns over time with Matplotlib, P&L calculation, better UI/UX