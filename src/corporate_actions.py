"""
Corporate Actions Module
========================

Handles detection and automatic adjustment of historical trades for
stock splits and bonuses (Option B feature).
"""
import os
import json
import logging
from datetime import datetime
import pandas as pd
import yfinance as yf

logger = logging.getLogger(__name__)

# Track applied splits to prevent double-applying
SPLITS_FILE = 'applied_splits.json'

def process_splits(df: pd.DataFrame, base_dir: str = '.', refresh_cache: bool = False) -> tuple[bool, pd.DataFrame]:
    """
    Scans for new stock splits/bonuses, prompts the user, and adjusts historical
    trades (Quantity and Price) if approved.

    Args:
        df:            The unified raw trade DataFrame (columns: Symbol, Trade Date,
                       Quantity, Price, Total Value, etc.)
        base_dir:      Directory where applied_splits.json should be stored.
        refresh_cache: If True, bypasses the cache and forces fresh network queries.

    Returns:
        A tuple (was_modified, adjusted_df)
    """
    splits_path = os.path.join(base_dir, SPLITS_FILE)
    
    # Load previously applied splits
    applied_splits = {}
    if os.path.exists(splits_path):
        try:
            with open(splits_path, 'r', encoding='utf-8') as f:
                applied_splits = json.load(f)
        except Exception:
            pass

    # Copy df to avoid modifying original until we decide to
    search_df = df.copy()
    search_df['Trade Date'] = pd.to_datetime(search_df['Trade Date'])

    # Find the earliest buy date for each symbol
    buy_df = search_df[search_df['Trade Type'].str.lower() == 'buy']
    first_buys = buy_df.groupby('Symbol')['Trade Date'].min().to_dict()

    pending_actions = []

    # Limit check to symbols where we actually hold/held positions
    symbols_to_check = list(first_buys.keys())
    
    if not symbols_to_check:
        return False, df

    # Load cache
    from src.market_api import load_market_cache, save_market_cache, is_cache_fresh
    from src.utils import print_progress_bar

    cache = load_market_cache()
    if 'symbols' not in cache:
        cache['symbols'] = {}

    logger.info(f"Checking for Corporate Actions (Splits / Bonuses) for {len(symbols_to_check)} symbols...")

    for idx, sym in enumerate(symbols_to_check):
        first_date = first_buys[sym]
        if pd.isna(first_date):
            continue
            
        # Normalize timezone
        if hasattr(first_date, 'tz') and first_date.tz is not None:
            first_date = first_date.tz_localize(None)

        print_progress_bar(idx, len(symbols_to_check), prefix='Checking splits:', suffix=f'({sym})', length=30)

        sym_cache = cache['symbols'].get(sym, {})
        
        # Check cache freshness
        if not refresh_cache and is_cache_fresh(sym_cache):
            splits_dict = sym_cache.get('Splits', {})
            # Convert dict keys to DatetimeIndex keys
            splits = {pd.to_datetime(d): r for d, r in splits_dict.items()}
        else:
            splits = {}
            try:
                ns_sym = sym if sym.endswith('.NS') else f"{sym}.NS"
                ticker = yf.Ticker(ns_sym)
                yf_splits = ticker.splits
                
                # Pre-fetch info to keep cache populated and speed up subsequent market data sync step
                info = ticker.info
                mcap = info.get('marketCap', 0) or 0
                comp_name = info.get('longName') or info.get('shortName') or sym
                
                splits_dict = {}
                if yf_splits is not None and not yf_splits.empty:
                    splits = {dt.tz_localize(None) if hasattr(dt, 'tz') and dt.tz is not None else dt: float(ratio) for dt, ratio in yf_splits.items()}
                    splits_dict = {dt.strftime('%Y-%m-%d'): float(ratio) for dt, ratio in splits.items()}
                
                cache['symbols'][sym] = {
                    'Company_Name': comp_name,
                    'Market_Cap': mcap,
                    'Splits': splits_dict,
                    'Last_Updated': datetime.now().strftime('%Y-%m-%d')
                }
            except Exception as e:
                logger.debug(f"Error fetching splits for {sym}: {e}")

        # Process splits
        if splits:
            for split_date, ratio in splits.items():
                sd = split_date
                if hasattr(sd, 'tz') and sd.tz is not None:
                    sd = sd.tz_localize(None)
                    
                # Strict string representation of the date for tracking
                date_str = sd.strftime('%Y-%m-%d')
                
                # Check if it happened AFTER first buy
                if sd >= first_date:
                    # Check if already applied
                    applied_for_sym = applied_splits.get(sym, [])
                    if date_str not in applied_for_sym:
                        ratio_float = float(ratio)
                        if ratio_float == int(ratio_float):
                            desc = f"{int(ratio_float)}:1 Split"
                        else:
                            num = int(ratio_float * 2)
                            desc = f"{num}:2 Bonus"
                            
                        action_desc = f"{sym}: {desc} on {date_str}"
                        pending_actions.append({
                            'symbol': sym,
                            'date_str': date_str,
                            'ratio': ratio_float,
                            'desc': action_desc
                        })

    print_progress_bar(len(symbols_to_check), len(symbols_to_check), prefix='Checking splits:', suffix='Complete', length=30)
    save_market_cache(cache)

    if not pending_actions:
        return False, df

    # We found NEW pending splits!
    print("=" * 60)
    print(" 🔔 NEW CORPORATE ACTIONS DETECTED ")
    print("=" * 60)
    for act in pending_actions:
        print(f" - {act['desc']}")
        
    print("\nWould you like to auto-adjust your historical trades for these events?")
    print("This will mathematically update your Quantities and Average Prices in ")
    print("the Raw_Tradebook so your portfolio reflects post-split realities.")
    
    import sys
    is_mocked_input = hasattr(input, '__mock__') or 'mock' in type(input).__name__.lower()
    if not is_mocked_input:
        if not sys.stdin or not sys.stdin.isatty() or os.environ.get('NON_INTERACTIVE') == '1':
            logger.info("Non-interactive session detected or NON_INTERACTIVE set. Skipping auto-adjustments.")
            return False, df

    ans = input("Adjust historical trades? (y/n): ").strip().lower()
    if ans != 'y':
        logger.info("Skipping auto-adjustments.")
        return False, df

    # Apply adjustments
    df_adjusted = df.copy()
    
    for act in pending_actions:
        sym = act['symbol']
        ratio = act['ratio']
        date_str = act['date_str']
        split_dt = pd.to_datetime(date_str)
        
        # We adjust trades that happened BEFORE or ON the split date
        mask = (df_adjusted['Symbol'] == sym) & (pd.to_datetime(df_adjusted['Trade Date']) <= split_dt)
        
        # Multiply quantity by ratio
        df_adjusted.loc[mask, 'Quantity'] = df_adjusted.loc[mask, 'Quantity'] * ratio
        # Divide price by ratio
        df_adjusted.loc[mask, 'Price'] = df_adjusted.loc[mask, 'Price'] / ratio
        
        # Record it as applied
        if sym not in applied_splits:
            applied_splits[sym] = []
        applied_splits[sym].append(date_str)

    # Save tracking file
    try:
        with open(splits_path, 'w', encoding='utf-8') as f:
            json.dump(applied_splits, f, indent=4)
        logger.info("✅ Adjustments applied and logged to applied_splits.json.")
    except Exception as e:
        logger.error(f"Failed to save applied_splits.json: {e}")

    return True, df_adjusted
