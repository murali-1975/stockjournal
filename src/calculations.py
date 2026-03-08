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

            if symbol not in symbol_state:
                symbol_state[symbol] = {'tranch_count': 0, 'cheat_count': 0, 'accumulated_cheat_value': 0}

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
    df_copy['Total Value'] = df_copy['Quantity'] * df_copy['Price']

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

    # --- Fetch Market Data (LTP & EMAs) for all symbols ---
    symbols = overall_df['Symbol'].tolist()
    market_data = fetch_market_data_from_yahoo(symbols)

    overall_df['LTP'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('LTP', 0.0))
    overall_df['EMA9'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA9', 0.0))
    overall_df['EMA10'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA10', 0.0))
    overall_df['EMA11'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA11', 0.0))
    overall_df['EMA21'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA21', 0.0))

    # --- Market Cap Classification ---
    overall_df['Market_Cap'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('Market_Cap', 0))
    overall_df['Cap'] = overall_df['Market_Cap'].apply(lambda mc: _classify_market_cap(mc, config))

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

    # --- Format and Order Columns for Overall Portfolio ---
    cols_order = [
        'Symbol', 'Cap', 'Total_Buy_Quantity', 'Total_Buy_Value', 'Average_Buy_Price',
        'Total_Sell_Quantity', 'Total_Sell_Value', 'Average_Sell_Price',
        'Current_Quantity', 'Invested_Value', 'LTP', 'Current_Value',
        'Realized_PnL', 'Unrealized_PnL', 'Total_PnL', 'Total_PnL_Percentage',
        'EMA9', 'EMA10', 'EMA11', 'EMA21'
    ]
    overall_df = overall_df[cols_order].sort_values(by='Symbol')

    # --- Current Portfolio View (only active positions) ---
    portfolio_df = overall_df[overall_df['Current_Quantity'] > 0][
        ['Symbol', 'Cap', 'Current_Quantity', 'Average_Buy_Price', 'Invested_Value', 'LTP',
         'EMA9', 'EMA10', 'EMA11', 'EMA21', 'Current_Value', 'Unrealized_PnL']
    ].copy()

    # --- Apply Stop Loss ---
    portfolio_df['SL'] = portfolio_df.apply(lambda row: _get_stop_loss(row, grouped_df), axis=1)

    # Reorder Current Portfolio columns
    port_cols = ['Symbol', 'Cap', 'Current_Quantity', 'Average_Buy_Price', 'SL', 'Invested_Value', 'LTP',
                 'EMA9', 'EMA10', 'EMA11', 'EMA21', 'Current_Value', 'Unrealized_PnL']
    portfolio_df = portfolio_df[port_cols]

    # Remove EMA columns from Overall Portfolio (only needed in Current Portfolio)
    overall_df = overall_df.drop(columns=['EMA9', 'EMA10', 'EMA11', 'EMA21'])

    return portfolio_df, overall_df
