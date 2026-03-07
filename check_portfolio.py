import pandas as pd
from process_trades import calculate_portfolio, load_data
import os

df = load_data('Tradebook Template.xlsx')
if df is not None:
    portfolio_df = calculate_portfolio(df)
    print(portfolio_df[['Symbol', 'LTP', 'Current_Value']].head())
