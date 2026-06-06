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
import os
import json
import logging
from datetime import datetime, timedelta
import pandas as pd

logger = logging.getLogger(__name__)

# Base directory paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_FILE = os.path.join(BASE_DIR, 'market_info_cache.json')

def load_market_cache() -> dict:
    """Loads the market metadata cache from market_info_cache.json."""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"Failed to load market cache: {e}")
    return {}

def save_market_cache(cache: dict) -> None:
    """Saves the market metadata cache to market_info_cache.json."""
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, indent=4)
    except Exception as e:
        logger.warning(f"Failed to save market cache: {e}")

def is_cache_fresh(symbol_data: dict, ttl_days: int = 7) -> bool:
    """Checks if the cached data for a symbol is within the TTL period."""
    if not symbol_data:
        return False
    if 'Company_Name' not in symbol_data or 'Market_Cap' not in symbol_data:
        return False
    last_updated_str = symbol_data.get('Last_Updated')
    if not last_updated_str:
        return False
    try:
        last_updated = datetime.strptime(last_updated_str, '%Y-%m-%d')
        if datetime.now() - last_updated < timedelta(days=ttl_days):
            return True
    except Exception:
        pass
    return False

def fetch_market_data_from_yahoo(symbols: list, classifications: dict = None, fetch_info: bool = True, refresh_cache: bool = False) -> dict:
    """
    Fetches LTP and EMA data from Yahoo Finance for a list of stock symbols.

    Downloads historical daily closing prices and computes Exponential Moving Averages
    (EMAs) with spans of 9, 10, 11, and 21 periods.
    - Core stocks: Daily closing EMAs (computed on daily historical data).
    - Satellite stocks: Weekly closing EMAs (resampled to weekly last prices).

    Args:
        symbols:         A list of NSE stock ticker symbols (without the '.NS' suffix).
                         Example: ['RELIANCE', 'TCS', 'INFY']
        classifications: Optional dict mapping Symbol to TF_Classification.
        fetch_info:      If True, fetches Market Cap and Company Name.
        refresh_cache:   If True, ignores cached values and forces fresh network fetches.

    Returns:
        A dictionary mapping each original symbol to its market data:
        {
            'RELIANCE': {
                'LTP': 2450.50,
                'EMA9': 2445.30,
                'EMA10': 2443.10,
                'EMA11': 2441.80,
                'EMA21': 2430.25,
                'Company_Name': 'Reliance Industries Limited',
                'Market_Cap': 18000000000000
            },
            ...
        }
    """
    default_data = {'LTP': 0.0, 'EMA9': 0.0, 'EMA10': 0.0, 'EMA11': 0.0, 'EMA21': 0.0, 'Market_Cap': 0, 'Prev_Day_Close': 0.0, 'Prev_Week_Close': 0.0, 'Prev_Month_Close': 0.0, 'Company_Name': ''}

    try:
        import yfinance as yf
    except ImportError:
        logger.error("yfinance library not found. Please install it using 'pip install yfinance'.")
        return {sym: default_data.copy() for sym in symbols}

    if not symbols:
        return {}

    logger.info(f"Downloading historical pricing from Yahoo Finance for {len(symbols)} symbols...")
    symbol_ns = [sym + '.NS' for sym in symbols]
    market_data = {sym: default_data.copy() for sym in symbols}

    try:
        # Download 2 years of data to ensure enough periods for a 21-week EMA (weekly calculations)
        # Setting threads=False to prevent yfinance from randomly deadlocking on bulk downloads
        data = yf.download(symbol_ns, period="2y", progress=False, threads=False)

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

                    # Calculate Previous Day Closing Price
                    prev_day_close = 0.0
                    if len(valid_series) >= 2:
                        prev_day_close = valid_series.iloc[-2]
                        if hasattr(prev_day_close, 'item'):
                            prev_day_close = prev_day_close.item()
                    market_data[sym]['Prev_Day_Close'] = round(float(prev_day_close), 2)

                    # Calculate Previous Week Closing Price
                    prev_week_close = 0.0
                    if isinstance(valid_series.index, pd.DatetimeIndex):
                        w_series = valid_series.resample('W').last().dropna()
                        if len(w_series) >= 2:
                            prev_week_close = w_series.iloc[-2]
                            if hasattr(prev_week_close, 'item'):
                                prev_week_close = prev_week_close.item()
                    market_data[sym]['Prev_Week_Close'] = round(float(prev_week_close), 2)

                    # Calculate Previous Month Closing Price
                    prev_month_close = 0.0
                    if isinstance(valid_series.index, pd.DatetimeIndex):
                        try:
                            m_series = valid_series.resample('ME').last().dropna()
                        except Exception:
                            m_series = valid_series.resample('M').last().dropna()
                        if len(m_series) >= 2:
                            prev_month_close = m_series.iloc[-2]
                            if hasattr(prev_month_close, 'item'):
                                prev_month_close = prev_month_close.item()
                    market_data[sym]['Prev_Month_Close'] = round(float(prev_month_close), 2)

                    # Determine if it's a Satellite stock
                    is_satellite = False
                    if classifications:
                        class_val = classifications.get(sym, '')
                        if isinstance(class_val, str) and 'satellite' in class_val.lower():
                            is_satellite = True

                    # Apply weekly closing logic for Satellite stocks if index is DatetimeIndex
                    if is_satellite and isinstance(valid_series.index, pd.DatetimeIndex):
                        ema_series = valid_series.resample('W').last().dropna()
                    else:
                        ema_series = valid_series

                    # Compute EMAs
                    if not ema_series.empty:
                        ema9 = ema_series.ewm(span=9, adjust=False).mean().iloc[-1]
                        ema10 = ema_series.ewm(span=10, adjust=False).mean().iloc[-1]
                        ema11 = ema_series.ewm(span=11, adjust=False).mean().iloc[-1]
                        ema21 = ema_series.ewm(span=21, adjust=False).mean().iloc[-1]
                    else:
                        # Fallback to daily if resampled is empty
                        ema9 = valid_series.ewm(span=9, adjust=False).mean().iloc[-1]
                        ema10 = valid_series.ewm(span=10, adjust=False).mean().iloc[-1]
                        ema11 = valid_series.ewm(span=11, adjust=False).mean().iloc[-1]
                        ema21 = valid_series.ewm(span=21, adjust=False).mean().iloc[-1]

                    for ema_key, ema_val in [('EMA9', ema9), ('EMA10', ema10), ('EMA11', ema11), ('EMA21', ema21)]:
                        if hasattr(ema_val, 'item'):
                            ema_val = ema_val.item()
                        market_data[sym][ema_key] = round(float(ema_val), 2)

        # Load metadata cache
        cache = load_market_cache()
        if 'symbols' not in cache:
            cache['symbols'] = {}

        # Handle Market Cap, Company Name, and Splits loading/fetching
        symbols_to_fetch_info = []
        if fetch_info:
            for sym in symbols:
                sym_cache = cache['symbols'].get(sym, {})
                if refresh_cache or not is_cache_fresh(sym_cache):
                    symbols_to_fetch_info.append(sym)
                else:
                    # Reuse cached data
                    market_data[sym]['Market_Cap'] = sym_cache['Market_Cap']
                    market_data[sym]['Company_Name'] = sym_cache['Company_Name']
                    splits_dict = sym_cache.get('Splits', {})
                    if splits_dict:
                        idx = pd.to_datetime(list(splits_dict.keys()))
                        market_data[sym]['Splits'] = pd.Series(list(splits_dict.values()), index=idx)

        # Fetch missing/expired symbol details
        if symbols_to_fetch_info:
            from src.utils import print_progress_bar
            logger.info(f"Fetching details (Market Cap, Name) for {len(symbols_to_fetch_info)} symbols from Yahoo Finance...")

            for idx, sym in enumerate(symbols_to_fetch_info):
                ns_sym = sym if sym.endswith('.NS') else f"{sym}.NS"
                print_progress_bar(idx, len(symbols_to_fetch_info), prefix='Fetching details:', suffix=f'({sym})', length=30)
                try:
                    ticker = yf.Ticker(ns_sym)
                    info = ticker.info
                    mcap = info.get('marketCap', 0) or 0
                    comp_name = info.get('longName') or info.get('shortName') or sym

                    market_data[sym]['Market_Cap'] = mcap
                    market_data[sym]['Company_Name'] = comp_name

                    # Fetch splits history
                    splits = ticker.splits
                    splits_dict = {}
                    if splits is not None and not splits.empty:
                        market_data[sym]['Splits'] = splits
                        splits_dict = {dt.strftime('%Y-%m-%d'): float(ratio) for dt, ratio in splits.items()}

                    # Update Cache dict
                    cache['symbols'][sym] = {
                        'Company_Name': comp_name,
                        'Market_Cap': mcap,
                        'Splits': splits_dict,
                        'Last_Updated': datetime.now().strftime('%Y-%m-%d')
                    }
                except Exception as e:
                    logger.debug(f"Error fetching details for {sym}: {e}")

            print_progress_bar(len(symbols_to_fetch_info), len(symbols_to_fetch_info), prefix='Fetching details:', suffix='Complete', length=30)
            save_market_cache(cache)

    except Exception as e:
        logger.error(f"Error fetching data from Yahoo Finance: {e}")

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
        logger.info(f"Fetching Historical Benchmark Returns since {start_dt.strftime('%d %b %Y')}...")
        
        # Start a bit earlier to ensure we catch the start date or closest trading day
        fetch_start = start_dt - pd.Timedelta(days=5)
        data = yf.download(tickers, start=fetch_start, progress=False, threads=False)
        
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
        logger.error(f"Error fetching benchmark data: {e}")
        
    return results
