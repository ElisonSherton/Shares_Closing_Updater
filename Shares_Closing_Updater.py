import pandas as pd
import numpy as np
import xlwings as xw
from datetime import datetime
from nsepy.history import get_price_list

# Read the daily sheet shares file
def main():
    wb = xw.Book.caller()
    DATE = wb.sheets[0].range("B1").value
    
    # Get the price list for defined Date above
    price_list = get_price_list(DATE)
    price_list = price_list[["SYMBOL", "CLOSE", "ISIN"]]

    # A one line function to return close price given ISIN
    get_close = lambda x: price_list[price_list.ISIN == x]["CLOSE"]
    
    Name = "XXX"
    share_row = 4

    while True:

        sheet_name = "dly"

        # Get the name of the stock
        Name = wb.sheets[sheet_name].range(f"B{share_row}").value
        
        # If the name is blank, then end the update
        if not Name:
            break
        
        # Get the ISIN Number of the stock
        ISIN = wb.sheets[sheet_name].range(f"A{share_row}").value
        ISIN = ISIN.strip()

        close_price = get_close(ISIN)

        # Fetch the current closing and correspondingly populate the closing column for the day
        if len(close_price) > 0:
            wb.sheets[sheet_name].range(f"C{share_row}").value = float(close_price)
        else:
            wb.sheets[sheet_name].range(f"C{share_row}").value = "#NA"

        share_row += 1