import os

import pandas as pd

from src.data_io import load_data
from src.config import load_config
from src.calculations import calculate_portfolios, process_grouped_trades

df = load_data('Tradebook Template.xlsx')
if df is not None:
    config = load_config('input.cfg')
    grouped_df = process_grouped_trades(df, config)
    portfolio_df, overall_df = calculate_portfolios(df, grouped_df)
    print(portfolio_df[['Symbol', 'LTP', 'Current_Value']].head())
