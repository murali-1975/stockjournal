"""
Calculations Module
===================

Contains the core mathematical and analytical logic for trade processing:

    1. **Trade Grouping** - Aggregates raw trades by date, symbol, and type.
       Assigns Tranche/Cheat labels based on configurable thresholds.

    2. **Portfolio Calculation** - Computes Average Buy/Sell Prices,
       Current Holdings, Invested Values, and PnL (Realized + Unrealized).

    3. **Stop Loss (SL)** - Determines dynamic stop-loss levels based on
       the highest Tranche reached for each symbol:
           - Tranch 1:  -10% from average buy price
           - Tranch 2:  Average price of Tranch 1 buys
           - Tranch 3:  Average price of Tranch 1+2+3 buys
           - Tranch >3: EMA 21 value
"""

import re

import numpy as np
import pandas as pd

from .data_io import load_equity_master
from .market_api import fetch_market_data_from_yahoo


def process_grouped_trades(df: pd.DataFrame, config: dict = None) -> pd.DataFrame:
    """
    Groups raw trades by Date, Symbol, and Trade Type to aggregate daily
    quantities and prices. Applies stateful Tranche and Cheat tracking
    per symbol if a config with TRANCH size is provided.

    The Tranche/Cheat classification algorithm works as follows:
        - A buy whose total value is within ±TRANCH_TOLERANCE of a multiple
          of TRANCH size is labeled 'Tranch N'.
        - A buy whose total value is ≤ CHEAT size is labeled 'Cheat N'.
        - Accumulated Cheat values that sum to a full Tranche trigger an
          automatic Tranche increment.
        - All sell transactions are labeled 'N/A'.

    Args:
        df:     The raw trade DataFrame (must have Trade Date, Symbol,
                Trade Type, Quantity, Price columns).
        config: Configuration dict with keys TRANCH, CHEAT, TRANCH_TOLERANCE.

    Returns:
        A sorted DataFrame with columns: Trade Date, Symbol, Trade Type,
        Total_Quantity, Average_Price, Total_Value, and optionally
        Tranches/Cheat.
    """
    df_copy = df.copy()
    df_copy['Total Value'] = df_copy['Quantity'] * df_copy['Price']

    grouped_df = df_copy.groupby(['Trade Date', 'Symbol', 'Trade Type']).agg(
        Total_Quantity=('Quantity', 'sum'),
        Total_Value=('Total Value', 'sum'),
    ).reset_index()

    # Sort by date so we process state chronologically
    grouped_df = grouped_df.sort_values(by=['Symbol', 'Trade Date'])

    grouped_df['Average_Price'] = grouped_df['Total_Value'] / grouped_df['Total_Quantity']
    grouped_df['Average_Price'] = grouped_df['Average_Price'].round(2)
    grouped_df['Total_Value'] = grouped_df['Total_Value'].round(2)

    if config and config.get('TRANCH'):
        tranch_size = config.get('TRANCH')
        cheat_size = config.get('CHEAT')

        # State trackers per symbol
        symbol_state = {}

        labels = []
        for index, row in grouped_df.iterrows():
            symbol = row['Symbol']
            trade_type = str(row['Trade Type']).lower()
            total_value = row['Total_Value']
            total_quantity = row['Total_Quantity']

            if symbol not in symbol_state:
                symbol_state[symbol] = {
                    'tranch_count': 0,
                    'cheat_count': 0,
                    'accumulated_cheat_value': 0.0,
                    'current_quantity': 0.0
                }

            # Update quantity trackers chronologically to detect full exit reset triggers
            if trade_type == 'buy':
                if symbol_state[symbol]['current_quantity'] <= 1e-5:
                    symbol_state[symbol]['tranch_count'] = 0
                    symbol_state[symbol]['cheat_count'] = 0
                    symbol_state[symbol]['accumulated_cheat_value'] = 0.0
                    symbol_state[symbol]['current_quantity'] = 0.0
                symbol_state[symbol]['current_quantity'] += total_quantity
            elif trade_type == 'sell':
                symbol_state[symbol]['current_quantity'] -= total_quantity
                if symbol_state[symbol]['current_quantity'] < 1e-5:
                    symbol_state[symbol]['current_quantity'] = 0.0

            if trade_type != 'buy':
                labels.append('N/A')
                continue

            # Configure dynamic tolerance
            tranch_tolerance_ratio = config.get('TRANCH_TOLERANCE', 0.10)

            # Check for Tranch match (+/- tolerance)
            approx_tranch_multiplier = round(total_value / tranch_size) if tranch_size > 0 else 0

            matched_tranch = False

            if approx_tranch_multiplier > 0:
                expected_tranch_value = approx_tranch_multiplier * tranch_size
                tranch_tolerance = tranch_tolerance_ratio * expected_tranch_value

                if abs(total_value - expected_tranch_value) <= tranch_tolerance:
                    symbol_state[symbol]['tranch_count'] += approx_tranch_multiplier
                    labels.append(f"Tranch {symbol_state[symbol]['tranch_count']}")
                    matched_tranch = True

            if not matched_tranch:
                is_cheat = False
                if cheat_size and cheat_size > 0 and total_value <= cheat_size:
                    is_cheat = True

                if is_cheat:
                    symbol_state[symbol]['cheat_count'] += 1

                    # Add to accumulated cheat value
                    symbol_state[symbol]['accumulated_cheat_value'] += total_value

                    if tranch_size > 0:
                        approx_tranches = round(symbol_state[symbol]['accumulated_cheat_value'] / tranch_size)
                        if approx_tranches > 0:
                            expected_val = approx_tranches * tranch_size
                            if abs(symbol_state[symbol]['accumulated_cheat_value'] - expected_val) <= (tranch_tolerance_ratio * expected_val):
                                symbol_state[symbol]['tranch_count'] += approx_tranches
                                symbol_state[symbol]['accumulated_cheat_value'] -= expected_val
                                if symbol_state[symbol]['accumulated_cheat_value'] < 0:
                                    symbol_state[symbol]['accumulated_cheat_value'] = 0

                    labels.append(f"Cheat {symbol_state[symbol]['cheat_count']}")
                else:
                    # Not a cheat → must be a tranche by default
                    mult = max(1, round(total_value / tranch_size) if tranch_size > 0 else 1)
                    symbol_state[symbol]['tranch_count'] += mult
                    labels.append(f"Tranch {symbol_state[symbol]['tranch_count']}")

        grouped_df['Tranches/Cheat'] = labels
        res_df = grouped_df[['Trade Date', 'Symbol', 'Trade Type', 'Total_Quantity', 'Average_Price', 'Total_Value', 'Tranches/Cheat']]
        return res_df.sort_values(by='Trade Date')
    else:
        res_df = grouped_df[['Trade Date', 'Symbol', 'Trade Type', 'Total_Quantity', 'Average_Price', 'Total_Value']]
        return res_df.sort_values(by='Trade Date')


def _get_stop_loss(row: pd.Series, grouped_df: pd.DataFrame) -> float:
    """
    Determines the dynamic Stop Loss price for a single portfolio row.

    The SL is based on the highest Tranche level reached for that symbol:
        - Tranch 1:  -10% from average buy price
        - Tranch 2:  Average price of all Tranch 1 buys
        - Tranch 3:  Average price of all Tranch 1, 2, 3 buys
        - Tranch >3: EMA 21 value (trend-following stop)

    Args:
        row:        A single row from the Current Portfolio DataFrame.
        grouped_df: The grouped transaction DataFrame with Tranch labels.

    Returns:
        The calculated stop-loss price, rounded to 2 decimal places.
    """
    sym = row['Symbol']
    avg_buy = row['Average_Buy_Price']
    ema21 = row['EMA21']

    if 'Tranches/Cheat' not in grouped_df.columns:
        return round(avg_buy * 0.9, 2)

    # Get all buys for this symbol
    buys = grouped_df[(grouped_df['Symbol'] == sym) & (grouped_df['Trade Type'] == 'buy')]

    tranch_nums = []
    for label in buys['Tranches/Cheat']:
        if pd.isna(label):
            continue
        m = re.match(r'Tranch\s+(\d+)', str(label))
        if m:
            tranch_nums.append(int(m.group(1)))

    if not tranch_nums:
        return round(avg_buy * 0.9, 2)

    max_tranch = max(tranch_nums)
    if max_tranch == 1:
        return round(avg_buy * 0.9, 2)
    elif max_tranch == 2:
        t1_buys = buys[buys['Tranches/Cheat'] == 'Tranch 1']
        if not t1_buys.empty:
            return round((t1_buys['Total_Value'].sum() / t1_buys['Total_Quantity'].sum()), 2)
        return round(avg_buy * 0.9, 2)
    elif max_tranch == 3:
        t123_buys = buys[buys['Tranches/Cheat'].isin(['Tranch 1', 'Tranch 2', 'Tranch 3'])]
        if not t123_buys.empty:
            return round((t123_buys['Total_Value'].sum() / t123_buys['Total_Quantity'].sum()), 2)
        return round(avg_buy, 2)
    else:  # > 3
        return round(ema21, 2)


def _classify_market_cap(market_cap: float, config: dict) -> str:
    """
    Classifies a stock based on its market capitalization.

    Uses thresholds from config (SMALL_CAP, MEDIUM_CAP, LARGE_CAP).
    Falls back to SEBI-standard defaults if config is missing.

    Args:
        market_cap: The market capitalization in absolute value (e.g., 3.47e11).
        config:     Configuration dict containing cap threshold definitions.

    Returns:
        One of 'Large Cap', 'Mid Cap', 'Small Cap', or '' if market cap is 0.
    """
    if market_cap <= 0:
        return ''

    # Extract thresholds from config or use SEBI defaults
    small_cap_cfg = config.get('SMALL_CAP', {})
    large_cap_cfg = config.get('LARGE_CAP', {})

    # Default SEBI thresholds (in absolute numbers)
    small_cap_upper = 347_000_000_000    # ₹34,700 Cr
    large_cap_lower = 1_050_000_000_000  # ₹1,05,000 Cr

    if isinstance(small_cap_cfg, dict) and small_cap_cfg.get('type') == 'below':
        small_cap_upper = small_cap_cfg['value']
    if isinstance(large_cap_cfg, dict) and large_cap_cfg.get('type') == 'above':
        large_cap_lower = large_cap_cfg['value']

    if market_cap >= large_cap_lower:
        return 'Large Cap'
    elif market_cap < small_cap_upper:
        return 'Small Cap'
    else:
        return 'Mid Cap'


def _get_split_info(symbol: str, first_buy_date, market_data: dict, applied_splits: dict) -> tuple[str, str]:
    """
    Checks if a stock had any splits/bonuses after the first buy date.

    Args:
        symbol:         The stock symbol.
        first_buy_date: The earliest buy date for this stock (datetime or NaT).
        market_data:    Dict from fetch_market_data_from_yahoo with 'Splits' key.
        applied_splits: Dict mapping symbols to lists of applied split dates.

    Returns:
        A tuple: (split_description_string, adj_required_string).
        e.g. ('2:1 Split on 2024-06-22', 'Yes')
    """
    sym_data = market_data.get(symbol, {})
    splits = sym_data.get('Splits')
    if splits is None or (hasattr(splits, 'empty') and splits.empty):
        return '', 'No'

    if pd.isna(first_buy_date):
        return '', 'No'

    # Normalize first_buy_date to timezone-naive for comparison
    if hasattr(first_buy_date, 'tz') and first_buy_date.tz is not None:
        first_buy_date = first_buy_date.tz_localize(None)

    relevant = []
    adj_required = 'No'
    applied_for_sym = applied_splits.get(symbol, [])

    for split_date, ratio in splits.items():
        # Normalize split_date to timezone-naive
        sd = split_date
        if hasattr(sd, 'tz') and sd.tz is not None:
            sd = sd.tz_localize(None)

        if sd >= first_buy_date:
            date_str = sd.strftime('%Y-%m-%d')
            ratio_float = float(ratio)
            if ratio_float == int(ratio_float):
                desc = f"{int(ratio_float)}:1 Split on {date_str}"
            else:
                # Express as fraction for bonus, e.g. 1.5 → 3:2
                num = int(ratio_float * 2)
                desc = f"{num}:2 Bonus on {date_str}"
            relevant.append(desc)
            
            if date_str not in applied_for_sym:
                adj_required = 'Yes'

    return ' | '.join(relevant), adj_required


def _get_latest_tranche_cheat(symbol: str, grouped_df: pd.DataFrame) -> str:
    """
    Returns the latest (highest-numbered) Tranche or Cheat label for a symbol.

    Scans the grouped transaction DataFrame for all buy rows matching the
    given symbol and finds the label with the highest number. If a symbol
    has both Tranche and Cheat entries, the one with the higher number is
    returned. If numbers are equal, Tranche takes priority.

    Args:
        symbol:     The stock ticker symbol to look up.
        grouped_df: The grouped transaction DataFrame with 'Tranches/Cheat' column.

    Returns:
        The latest label string (e.g., 'Tranch 3', 'Cheat 2'), or '' if none found.
    """
    import re

    if 'Tranches/Cheat' not in grouped_df.columns:
        return ''

    sym_buys = grouped_df[
        (grouped_df['Symbol'] == symbol) &
        (grouped_df['Trade Type'].str.lower() == 'buy')
    ]

    if sym_buys.empty:
        return ''

    max_label = ''
    max_num = 0

    for label in sym_buys['Tranches/Cheat']:
        if label == 'N/A' or pd.isna(label):
            continue
        match = re.match(r'(Tranch|Cheat)\s+(\d+)', str(label))
        if match:
            num = int(match.group(2))
            label_type = match.group(1)
            # Higher number wins; if equal, Tranch > Cheat
            if num > max_num or (num == max_num and label_type == 'Tranch'):
                max_num = num
                max_label = label

    return max_label


def calculate_portfolios(df: pd.DataFrame, grouped_df: pd.DataFrame, config: dict = None) -> tuple:
    """
    Calculates current holdings, overall trade summary, and PnL statistics.

    This function:
        1. Aggregates all buy and sell transactions per symbol.
        2. Computes average buy/sell prices and invested values.
        3. Fetches live market data (LTP, EMAs, Market Cap) from Yahoo Finance.
        4. Classifies each stock as Large Cap / Mid Cap / Small Cap.
        5. Calculates Realized PnL (from sold positions), Unrealized PnL
           (from current holdings), and Total PnL with percentage.
        6. Builds a Current Portfolio view (active positions only) with
           dynamic Stop Loss values.
        7. Drops EMA columns from the Overall Portfolio (they only appear
           in Current Portfolio).

    Args:
        df:         The raw (or merged) trade DataFrame.
        grouped_df: The grouped transaction DataFrame with Tranche labels.
        config:     Configuration dict with cap classification thresholds.

    Returns:
        A tuple of (current_portfolio_df, overall_portfolio_df).
    """
    if config is None:
        config = {}
    df_copy = df.copy()
    df_copy['Trade Type'] = df_copy['Trade Type'].str.lower()
    df_copy['Trade Date'] = pd.to_datetime(df_copy['Trade Date'])
    df_copy['Total Value'] = df_copy['Quantity'] * df_copy['Price']
    today = pd.Timestamp.now().normalize()

    # --- First Buy Date per symbol ---
    first_buy = df_copy[df_copy['Trade Type'] == 'buy'].groupby('Symbol')['Trade Date'].min().reset_index()
    first_buy.columns = ['Symbol', 'First_Buy_Date']

    # --- Last Sell Date per symbol ---
    sells_dates = df_copy[df_copy['Trade Type'] == 'sell'].groupby('Symbol')['Trade Date'].max().reset_index()
    sells_dates.columns = ['Symbol', 'Last_Sell_Date']

    # --- Aggregating Buys ---
    buys = df_copy[df_copy['Trade Type'] == 'buy'].groupby('Symbol').agg(
        Total_Buy_Quantity=('Quantity', 'sum'),
        Total_Buy_Value=('Total Value', 'sum')
    ).reset_index()
    buys['Average_Buy_Price'] = (buys['Total_Buy_Value'] / buys['Total_Buy_Quantity']).round(2)

    # --- Aggregating Sells ---
    sells = df_copy[df_copy['Trade Type'] == 'sell'].groupby('Symbol').agg(
        Total_Sell_Quantity=('Quantity', 'sum'),
        Total_Sell_Value=('Total Value', 'sum')
    ).reset_index()
    sells['Average_Sell_Price'] = (sells['Total_Sell_Value'] / sells['Total_Sell_Quantity']).round(2)

    # --- Merge Buy and Sell stats ---
    overall_df = pd.merge(buys, sells, on='Symbol', how='outer').fillna(0)

    overall_df['Current_Quantity'] = overall_df['Total_Buy_Quantity'] - overall_df['Total_Sell_Quantity']
    overall_df['Invested_Value'] = (overall_df['Current_Quantity'] * overall_df['Average_Buy_Price']).round(2)

    # --- TF Sector & Classification from Equity Master (loaded early for EMA classification) ---
    equity_master = load_equity_master(config)
    classifications = {}
    if equity_master is not None:
        classifications = dict(zip(equity_master['Symbol'], equity_master['TF_Classification']))

    # --- Fetch Market Data (LTP & EMAs) for all symbols ---
    symbols = overall_df['Symbol'].tolist()
    market_data = fetch_market_data_from_yahoo(symbols, classifications=classifications)

    overall_df['LTP'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('LTP', 0.0))
    overall_df['Prev_Week_Close'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('Prev_Week_Close', 0.0))
    overall_df['EMA9'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA9', 0.0))
    overall_df['EMA10'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA10', 0.0))
    overall_df['EMA11'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA11', 0.0))
    overall_df['EMA21'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA21', 0.0))

    # --- Market Cap Classification ---
    overall_df['Market_Cap'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('Market_Cap', 0))
    overall_df['Cap'] = overall_df['Market_Cap'].apply(lambda mc: _classify_market_cap(mc, config))

    # --- Latest Tranche/Cheat per symbol ---
    overall_df['Latest_Tranche'] = overall_df['Symbol'].apply(
        lambda sym: _get_latest_tranche_cheat(sym, grouped_df)
    )

    # --- TF Sector & Classification Merge ---
    if equity_master is not None:
        overall_df = overall_df.merge(equity_master, on='Symbol', how='left')
        overall_df['TF_Sector'] = overall_df['TF_Sector'].fillna('')
        overall_df['TF_Classification'] = overall_df['TF_Classification'].fillna('')
    else:
        overall_df['TF_Sector'] = ''
        overall_df['TF_Classification'] = ''

    overall_df['Current_Value'] = (overall_df['Current_Quantity'] * overall_df['LTP']).round(2)

    # --- PnL Calculations ---
    overall_df['Realized_PnL'] = (overall_df['Total_Sell_Value'] - (overall_df['Total_Sell_Quantity'] * overall_df['Average_Buy_Price'])).round(2)
    overall_df['Unrealized_PnL'] = (overall_df['Current_Value'] - overall_df['Invested_Value']).round(2)
    overall_df['Total_PnL'] = (overall_df['Realized_PnL'] + overall_df['Unrealized_PnL']).round(2)

    overall_df['Total_PnL_Percentage'] = np.where(
        overall_df['Total_Buy_Value'] > 0,
        (overall_df['Total_PnL'] / overall_df['Total_Buy_Value']),
        0
    )

    # --- Holding Period (days) ---
    overall_df = overall_df.merge(first_buy, on='Symbol', how='left')
    overall_df = overall_df.merge(sells_dates, on='Symbol', how='left')
    # If fully sold (Current_Quantity == 0), use last sell date; otherwise use today
    overall_df['Holding_End'] = overall_df.apply(
        lambda row: row['Last_Sell_Date'] if row['Current_Quantity'] == 0 and pd.notna(row['Last_Sell_Date']) else today,
        axis=1
    )
    overall_df['Holding_Period'] = (overall_df['Holding_End'] - overall_df['First_Buy_Date']).dt.days
    overall_df['Holding_Period'] = overall_df['Holding_Period'].fillna(0).astype(int)

    # --- Load applied splits to correctly set Adj_Required flag ---
    import json
    import os
    applied_splits = {}
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    splits_path = os.path.join(base_dir, 'applied_splits.json')
    if os.path.exists(splits_path):
        try:
            with open(splits_path, 'r', encoding='utf-8') as f:
                applied_splits = json.load(f)
        except Exception:
            pass

    # --- Split/Bonus Detection ---
    def assign_split_info(row):
        return _get_split_info(row['Symbol'], row.get('First_Buy_Date'), market_data, applied_splits)

    split_results = overall_df.apply(assign_split_info, axis=1)
    overall_df['Split_Info'] = [res[0] for res in split_results]
    overall_df['Adj_Required'] = [res[1] for res in split_results]

    # --- Format and Order Columns for Overall Portfolio ---
    cols_order = [
        'Symbol', 'Cap', 'TF_Sector', 'TF_Classification', 'Latest_Tranche',
        'Total_Buy_Quantity', 'Total_Buy_Value', 'Average_Buy_Price',
        'Total_Sell_Quantity', 'Total_Sell_Value', 'Average_Sell_Price',
        'Current_Quantity', 'Invested_Value', 'LTP', 'Current_Value',
        'Realized_PnL', 'Unrealized_PnL', 'Total_PnL', 'Total_PnL_Percentage',
        'Holding_Period', 'Split_Info', 'Adj_Required',
        'Prev_Week_Close', 'EMA9', 'EMA10', 'EMA11', 'EMA21'
    ]
    overall_df = overall_df[cols_order].sort_values(by='Symbol')

    # --- Current Portfolio View (only active positions) ---
    portfolio_df = overall_df[overall_df['Current_Quantity'] > 0][
        ['Symbol', 'Cap', 'TF_Sector', 'TF_Classification', 'Latest_Tranche',
         'Current_Quantity', 'Average_Buy_Price', 'Invested_Value', 'LTP',
         'Prev_Week_Close', 'EMA9', 'EMA10', 'EMA11', 'EMA21', 'Current_Value', 'Unrealized_PnL',
         'Holding_Period', 'Split_Info', 'Adj_Required']
    ].copy()

    # --- Apply Stop Loss ---
    portfolio_df['SL'] = portfolio_df.apply(lambda row: _get_stop_loss(row, grouped_df), axis=1)

    # --- LTP vs SL Difference ---
    portfolio_df['LTP_SL_Diff'] = (portfolio_df['LTP'] - portfolio_df['SL']).round(2)
    portfolio_df['LTP_SL_Diff_Pct'] = np.where(
        portfolio_df['LTP'] > 0,
        ((portfolio_df['LTP'] - portfolio_df['SL']) / portfolio_df['LTP']),
        0
    )

    # Reorder Current Portfolio columns
    port_cols = ['Symbol', 'Cap', 'TF_Sector', 'TF_Classification', 'Latest_Tranche',
                 'Current_Quantity', 'Average_Buy_Price', 'SL',
                 'LTP_SL_Diff', 'LTP_SL_Diff_Pct', 'Invested_Value', 'LTP',
                 'Prev_Week_Close', 'EMA9', 'EMA10', 'EMA11', 'EMA21', 'Current_Value', 'Unrealized_PnL',
                 'Holding_Period', 'Split_Info', 'Adj_Required']
    portfolio_df = portfolio_df[port_cols]

    # Remove EMA and Previous Week Close columns from Overall Portfolio (only needed in Current Portfolio)
    overall_df = overall_df.drop(columns=['EMA9', 'EMA10', 'EMA11', 'EMA21', 'Prev_Week_Close'])

    return portfolio_df, overall_df
