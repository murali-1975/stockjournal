"""
Market Data API Module
======================

Fetches real-time and historical market data from Yahoo Finance using the
`yfinance` library. Provides the following indicators for each stock symbol:

    - **LTP** (Last Traded Price): The most recent closing price.
    - **EMA 9, 10, 11, 21**: Exponential Moving Averages computed over
      3 months of daily closing prices.

All symbols are assumed to be Indian (NSE) stocks; the `.NS` suffix is
appended automatically before querying Yahoo Finance.

Dependencies:
    - yfinance (pip install yfinance)
"""


def fetch_market_data_from_yahoo(symbols: list) -> dict:
    """
    Fetches LTP and EMA data from Yahoo Finance for a list of stock symbols.

    Downloads 3 months of historical daily closing prices and computes
    Exponential Moving Averages (EMAs) with spans of 9, 10, 11, and 21 days.

    Args:
        symbols: A list of NSE stock ticker symbols (without the '.NS' suffix).
                 Example: ['RELIANCE', 'TCS', 'INFY']

    Returns:
        A dictionary mapping each original symbol to its market data:
        {
            'RELIANCE': {
                'LTP': 2450.50,
                'EMA9': 2445.30,
                'EMA10': 2443.10,
                'EMA11': 2441.80,
                'EMA21': 2430.25
            },
            ...
        }

        Returns default values of 0.0 for all fields if the symbol data
        is unavailable or if yfinance is not installed.
    """
    import pandas as pd

    default_data = {'LTP': 0.0, 'EMA9': 0.0, 'EMA10': 0.0, 'EMA11': 0.0, 'EMA21': 0.0, 'Market_Cap': 0}

    try:
        import yfinance as yf
    except ImportError:
        print("yfinance library not found. Please install it using 'pip install yfinance'.")
        return {sym: default_data.copy() for sym in symbols}

    if not symbols:
        return {}

    print(f"Fetching Market Data (LTP, EMAs & Market Cap) from Yahoo Finance for {len(symbols)} symbols...")
    symbol_ns = [sym + '.NS' for sym in symbols]
    market_data = {sym: default_data.copy() for sym in symbols}

    try:
        # Download 3 months of data to ensure enough periods for a 21-day EMA
        data = yf.download(symbol_ns, period="3mo", progress=False)

        if 'Close' in data:
            close_data = data['Close']
            for sym, ns_sym in zip(symbols, symbol_ns):
                if len(symbols) == 1:
                    series = close_data
                else:
                    if ns_sym in close_data.columns:
                        series = close_data[ns_sym]
                    else:
                        continue

                # Drop NAs to compute valid EMAs
                valid_series = series.dropna()
                if not valid_series.empty:
                    market_data[sym]['LTP'] = round(float(valid_series.iloc[-1]), 2)
                    market_data[sym]['EMA9'] = round(float(valid_series.ewm(span=9, adjust=False).mean().iloc[-1]), 2)
                    market_data[sym]['EMA10'] = round(float(valid_series.ewm(span=10, adjust=False).mean().iloc[-1]), 2)
                    market_data[sym]['EMA11'] = round(float(valid_series.ewm(span=11, adjust=False).mean().iloc[-1]), 2)
                    market_data[sym]['EMA21'] = round(float(valid_series.ewm(span=21, adjust=False).mean().iloc[-1]), 2)

        # Fetch Market Cap and Splits for each symbol individually
        for sym, ns_sym in zip(symbols, symbol_ns):
            try:
                ticker = yf.Ticker(ns_sym)
                info = ticker.info
                market_data[sym]['Market_Cap'] = info.get('marketCap', 0) or 0

                # Fetch split/bonus history
                splits = ticker.splits
                if splits is not None and not splits.empty:
                    market_data[sym]['Splits'] = splits
            except Exception:
                pass

    except Exception as e:
        print(f"Error fetching data from Yahoo Finance: {e}")

    return market_data

