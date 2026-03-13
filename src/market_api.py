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
                    # Single symbol: close_data may be a DataFrame or Series;
                    # squeeze to ensure we have a 1-D Series.
                    if hasattr(close_data, 'squeeze'):
                        series = close_data.squeeze()
                    else:
                        series = close_data
                else:
                    if ns_sym in close_data.columns:
                        series = close_data[ns_sym]
                    else:
                        continue

                # Drop NAs to compute valid EMAs
                valid_series = series.dropna()
                if not valid_series.empty:
                    ltp_val = valid_series.iloc[-1]
                    if hasattr(ltp_val, 'item'):
                        ltp_val = ltp_val.item()
                    market_data[sym]['LTP'] = round(float(ltp_val), 2)

                    ema9 = valid_series.ewm(span=9, adjust=False).mean().iloc[-1]
                    ema10 = valid_series.ewm(span=10, adjust=False).mean().iloc[-1]
                    ema11 = valid_series.ewm(span=11, adjust=False).mean().iloc[-1]
                    ema21 = valid_series.ewm(span=21, adjust=False).mean().iloc[-1]
                    for ema_key, ema_val in [('EMA9', ema9), ('EMA10', ema10), ('EMA11', ema11), ('EMA21', ema21)]:
                        if hasattr(ema_val, 'item'):
                            ema_val = ema_val.item()
                        market_data[sym][ema_key] = round(float(ema_val), 2)

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


def fetch_benchmark_returns(start_date_str: str, custom_benchmarks: list = None) -> dict:
    """
    Fetches the historical start-date prices and current LTPs for Nifty benchmark
    ETFs to calculate portfolio-relative benchmark returns.
    
    Args:
        start_date_str: String date in 'DD-MM-YYYY' format.
        custom_benchmarks: Optional list of specific benchmark names from config.
        
    Returns:
        A dict with the performance metrics:
        {
            'Nifty 50': {'Start_Price': 250.0, 'LTP': 275.0, 'Return_Pct': 10.0},
            ...
        }
    """
    import pandas as pd
    
    # Map common index names to their Yahoo Finance ticker symbols
    KNOWN_BENCHMARKS = {
        'NIFTY_50': '^NSEI',
        'NIFTY_BANK': '^NSEBANK',
        'CNXMIDCAP': '^NSEMDCP50',
        'CNX100': '^CNX100',
        'NIFTY_MIDCAP_150': '^NSEMDCP150',
        'SENSEX': '^BSESN',
        # Fallbacks to ETFs
        'NIFTY_SMALLCAP_250': 'SMALLCAP.NS'
    }

    # Map friendly names to highly liquid tracking ETFs since direct indices
    # on Yahoo Finance (like ^CNXSC) are often broken/missing.
    benchmarks = {
        'Nifty 50': 'NIFTYBEES.NS',
        'Nifty Midcap 150': 'MID150BEES.NS',
        'Nifty Smallcap 250': 'SMALLCAP.NS'
    }

    if custom_benchmarks:
        benchmarks = {}
        for b_name in custom_benchmarks:
            upper_name = b_name.upper().replace(' ', '_')
            if upper_name in KNOWN_BENCHMARKS:
                benchmarks[b_name] = KNOWN_BENCHMARKS[upper_name]
            else:
                # If the user passes a direct YF ticker like ^NSEI, use it as-is
                benchmarks[b_name] = upper_name if '^' in upper_name else b_name
    
    results = {
        name: {'Start_Price': 0.0, 'LTP': 0.0, 'Return_Pct': 0.0} 
        for name in benchmarks.keys()
    }
    
    if not start_date_str:
        return results
        
    try:
        start_dt = pd.to_datetime(start_date_str, format='%d-%m-%Y', errors='coerce')
        if pd.isna(start_dt):
            return results
            
        import yfinance as yf
        tickers = list(benchmarks.values())
        print(f"\nFetching Historical Benchmark Returns since {start_dt.strftime('%d %b %Y')}...")
        
        # Start a bit earlier to ensure we catch the start date or closest trading day
        fetch_start = start_dt - pd.Timedelta(days=5)
        data = yf.download(tickers, start=fetch_start, progress=False)
        
        if 'Close' not in data:
            return results
            
        close_data = data['Close']
        
        for name, ticker in benchmarks.items():
            if ticker in close_data.columns:
                series = close_data[ticker].dropna()
                if not series.empty:
                    # Find the closest date on or after the requested start_date
                    future_dates = series[series.index.tz_localize(None) >= start_dt]
                    if not future_dates.empty:
                        start_price = float(future_dates.iloc[0])
                    else:
                        # Fallback to the last available if somehow there are no future dates
                        start_price = float(series.iloc[-1])
                        
                    current_price = float(series.iloc[-1])
                    
                    results[name]['Start_Price'] = round(start_price, 2)
                    results[name]['LTP'] = round(current_price, 2)
                    
                    if start_price > 0:
                        results[name]['Return_Pct'] = (current_price - start_price) / start_price
                        
    except Exception as e:
        print(f"Error fetching benchmark data: {e}")
        
    return results


