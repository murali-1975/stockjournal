"""
Corporate Actions Module
========================

Handles detection and automatic adjustment of historical trades for
stock splits and bonuses (Option B feature).
"""
import os
import json
import pandas as pd
import yfinance as yf

# Track applied splits to prevent double-applying
SPLITS_FILE = 'applied_splits.json'

def process_splits(df: pd.DataFrame, base_dir: str = '.') -> tuple[bool, pd.DataFrame]:
    """
    Scans for new stock splits/bonuses, prompts the user, and adjusts historical
    trades (Quantity and Price) if approved.

    Args:
        df:       The unified raw trade DataFrame (columns: Symbol, Trade Date,
                  Quantity, Price, Total Value, etc.)
        base_dir: Directory where applied_splits.json should be stored.

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

    # Find the earliest buy date for each symbol
    buy_df = df[df['Trade Type'].str.lower() == 'buy']
    first_buys = buy_df.groupby('Symbol')['Trade Date'].min().to_dict()

    pending_actions = []

    print("\nChecking for Corporate Actions (Splits / Bonuses)...")
    
    # Limit check to symbols where we actually hold/held positions
    symbols_to_check = list(first_buys.keys())
    
    # We shouldn't hit the API too hard sequentially if possible, but yfinance
    # ticker.splits is relatively fast.
    for sym in symbols_to_check:
        first_date = first_buys[sym]
        if pd.isna(first_date):
            continue
            
        # Normalize timezone
        if hasattr(first_date, 'tz') and first_date.tz is not None:
            first_date = first_date.tz_localize(None)

        try:
            ns_sym = sym if sym.endswith('.NS') else f"{sym}.NS"
            ticker = yf.Ticker(ns_sym)
            splits = ticker.splits
            
            if splits is None or (hasattr(splits, 'empty') and splits.empty):
                continue
                
            # Process each split found
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
                        
        except Exception:
            pass

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
    
    ans = input("Adjust historical trades? (y/n): ").strip().lower()
    if ans != 'y':
        print("Skipping auto-adjustments.")
        return False, df

    # Apply adjustments
    df_adjusted = df.copy()
    
    for act in pending_actions:
        sym = act['symbol']
        ratio = act['ratio']
        date_str = act['date_str']
        split_dt = pd.to_datetime(date_str)
        
        # We adjust trades that happened BEFORE or ON the split date
        # (Technically trades ON the split date usually trade ex-split, but
        # simple assumption is that anything up to the date needs adjusting)
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
        print("✅ Adjustments applied and logged to applied_splits.json.")
    except Exception as e:
        print(f"Failed to save applied_splits.json: {e}")

    return True, df_adjusted
