import pandas as pd
import os

def load_data(input_file):
    """
    Loads trade data from an Excel file into a pandas DataFrame.
    
    Args:
        input_file (str): Path to the input Excel file.
        
    Returns:
        pd.DataFrame or None: The loaded DataFrame, or None if an error occurred.
    """
    print(f"Reading {input_file}...")
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"Error: {input_file} not found.")
        return None
    except Exception as e:
        print(f"Error reading {input_file}: {e}")
        return None

    # Check for required columns
    required_cols = ['Trade Date', 'Symbol', 'Trade Type', 'Quantity', 'Price']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        print(f"Error: The input file is missing required columns: {missing_cols}")
        return None
        
    print("Data loaded successfully.")
    return df

def load_config(config_file):
    """
    Parses a config file to extract key-value pairs and evaluates percentages.
    
    Args:
        config_file (str): Path to the config file.
        
    Returns:
        dict: A dictionary of configuration parameters.
    """
    config = {}
    raw_lines = {}
    try:
        with open(config_file, 'r') as f:
            for line in f:
                if '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    raw_lines[key] = value
                    
                    # First pass: try direct numeric conversion
                    numeric_part = value.split()[0]
                    if '%' not in numeric_part:
                        try:
                            config[key] = float(numeric_part)
                        except ValueError:
                            pass
                            
            # Second pass: evaluate percentages of TOTAL_PORTFOLIO and specific string rules
            total_portfolio = config.get('TOTAL_PORTFOLIO', 0)
            for key, value in raw_lines.items():
                if isinstance(value, str):
                    if '%' in value and 'TOTAL_PORTFOLIO' in value:
                        try:
                            percent_str = value.split('%')[0].strip()
                            percent_val = float(percent_str) / 100
                            config[key] = percent_val * total_portfolio
                        except ValueError:
                            pass
                    elif key == 'TRANCH_TOLERANCE':
                        if '%' in value:
                            val_str = value.replace('+/-', '').replace('%', '').strip()
                            try:
                                config[key] = float(val_str) / 100
                            except ValueError:
                                pass
                    elif key == 'CHEAT':
                        if '<' in value:
                            val_str = value.replace('<', '').strip()
                            try:
                                config[key] = float(val_str)
                            except ValueError:
                                pass
                        elif '%' not in value:
                            try:
                                config[key] = float(value)
                            except ValueError:
                                pass
    except FileNotFoundError:
        print(f"Config file {config_file} not found. Skipping tranche logic.")
        
    return config

def process_grouped_trades(df, config=None):
    """
    Groups trades by Date, Symbol, and Trade Type to aggregate daily quantities and prices.
    Applies stateful Tranch and Cheat tracking per symbol if config is provided.
    
    Args:
        df (pd.DataFrame): The raw trade data.
        config (dict, optional): Configuration containing TRANCH and CHEAT sizes.
        
    Returns:
        pd.DataFrame: Summarized trade data with average execution prices and labels.
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
                    # Increment Tranch count
                    symbol_state[symbol]['tranch_count'] += approx_tranch_multiplier
                    labels.append(f"Tranch {symbol_state[symbol]['tranch_count']}")
                    matched_tranch = True
            
            if not matched_tranch:
                is_cheat = False
                # User constraint: Cheat max size is <= cheat_size
                if cheat_size and cheat_size > 0 and total_value <= cheat_size:
                    is_cheat = True
                    
                if is_cheat:
                    # Increment Cheat count organically by 1 for each trade date
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
                    # If it's not a cheat, it must be a tranche by default
                    # Determine which tranche it falls into (at least 1)
                    mult = max(1, round(total_value / tranch_size) if tranch_size > 0 else 1)
                    symbol_state[symbol]['tranch_count'] += mult
                    labels.append(f"Tranch {symbol_state[symbol]['tranch_count']}")
            
        grouped_df['Tranches/Cheat'] = labels
        res_df = grouped_df[['Trade Date', 'Symbol', 'Trade Type', 'Total_Quantity', 'Average_Price', 'Total_Value', 'Tranches/Cheat']]
        return res_df.sort_values(by='Trade Date')
    else:
        res_df = grouped_df[['Trade Date', 'Symbol', 'Trade Type', 'Total_Quantity', 'Average_Price', 'Total_Value']]
        return res_df.sort_values(by='Trade Date')

def fetch_market_data_from_yahoo(symbols):
    """
    Fetches the Last Traded Price (LTP) and EMAs (9, 10, 11, 21) from Yahoo Finance.
    Assuming Indian stocks by default, appending '.NS' to the symbols.
    
    Args:
        symbols (list): A list of stock ticker symbols.
        
    Returns:
        dict: A dictionary mapping original symbol to a dict of {'LTP': val, 'EMA9': val, 'EMA10': val, 'EMA11': val, 'EMA21': val}.
    """
    import pandas as pd
    default_data = {'LTP': 0.0, 'EMA9': 0.0, 'EMA10': 0.0, 'EMA11': 0.0, 'EMA21': 0.0}
    try:
        import yfinance as yf
    except ImportError:
        print("yfinance library not found. Please install it using 'pip install yfinance'.")
        return {sym: default_data.copy() for sym in symbols}
        
    if not symbols:
        return {}
        
    print(f"Fetching Market Data (LTP & EMAs) from Yahoo Finance for {len(symbols)} symbols...")
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
                
    except Exception as e:
        print(f"Error fetching data from Yahoo Finance: {e}")
            
    return market_data

def calculate_portfolios(df, grouped_df):
    """
    Calculates the current stock holdings, overall trades, and PnL, along with Stop Loss.
    
    Args:
        df (pd.DataFrame): The raw trade data.
        grouped_df (pd.DataFrame): The grouped trade data containing Tranch labels.
        
    Returns:
        tuple: (Current Portfolio DataFrame, Overall Portfolio DataFrame)
    """
    df_copy = df.copy()
    df_copy['Trade Type'] = df_copy['Trade Type'].str.lower()
    df_copy['Total Value'] = df_copy['Quantity'] * df_copy['Price']
    
    # Aggregating Buys
    buys = df_copy[df_copy['Trade Type'] == 'buy'].groupby('Symbol').agg(
        Total_Buy_Quantity=('Quantity', 'sum'),
        Total_Buy_Value=('Total Value', 'sum')
    ).reset_index()
    buys['Average_Buy_Price'] = (buys['Total_Buy_Value'] / buys['Total_Buy_Quantity']).round(2)
    
    # Aggregating Sells
    sells = df_copy[df_copy['Trade Type'] == 'sell'].groupby('Symbol').agg(
        Total_Sell_Quantity=('Quantity', 'sum'),
        Total_Sell_Value=('Total Value', 'sum')
    ).reset_index()
    sells['Average_Sell_Price'] = (sells['Total_Sell_Value'] / sells['Total_Sell_Quantity']).round(2)
    
    # Merge Buy and Sell stats
    overall_df = pd.merge(buys, sells, on='Symbol', how='outer').fillna(0)
    
    overall_df['Current_Quantity'] = overall_df['Total_Buy_Quantity'] - overall_df['Total_Sell_Quantity']
    overall_df['Invested_Value'] = (overall_df['Current_Quantity'] * overall_df['Average_Buy_Price']).round(2)
    
    # Fetch Market Data (LTP & EMAs) for all symbols
    symbols = overall_df['Symbol'].tolist()
    market_data = fetch_market_data_from_yahoo(symbols)
    
    # Apply results
    overall_df['LTP'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('LTP', 0.0))
    overall_df['EMA9'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA9', 0.0))
    overall_df['EMA10'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA10', 0.0))
    overall_df['EMA11'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA11', 0.0))
    overall_df['EMA21'] = overall_df['Symbol'].apply(lambda x: market_data.get(x, {}).get('EMA21', 0.0))
    
    overall_df['Current_Value'] = (overall_df['Current_Quantity'] * overall_df['LTP']).round(2)
    
    # PnL Calculations
    # Realized PnL = Sold Value - (Sold Qty * Avg Buy Price)
    overall_df['Realized_PnL'] = (overall_df['Total_Sell_Value'] - (overall_df['Total_Sell_Quantity'] * overall_df['Average_Buy_Price'])).round(2)
    overall_df['Unrealized_PnL'] = (overall_df['Current_Value'] - overall_df['Invested_Value']).round(2)
    overall_df['Total_PnL'] = (overall_df['Realized_PnL'] + overall_df['Unrealized_PnL']).round(2)
    
    # Calculate Total PnL % based on the Total Buy Value
    # Using numpy's where to avoid division by zero
    import numpy as np
    overall_df['Total_PnL_Percentage'] = np.where(overall_df['Total_Buy_Value'] > 0, 
                                                 (overall_df['Total_PnL'] / overall_df['Total_Buy_Value']), 
                                                 0)
    
    # Format and Order Columns
    cols_order = [
         'Symbol', 'Total_Buy_Quantity', 'Total_Buy_Value', 'Average_Buy_Price',
         'Total_Sell_Quantity', 'Total_Sell_Value', 'Average_Sell_Price',
         'Current_Quantity', 'Invested_Value', 'LTP', 'Current_Value',
         'Realized_PnL', 'Unrealized_PnL', 'Total_PnL', 'Total_PnL_Percentage',
         'EMA9', 'EMA10', 'EMA11', 'EMA21'
    ]
    overall_df = overall_df[cols_order].sort_values(by='Symbol')
    
    # Current Portfolio View (only Active positions)
    portfolio_df = overall_df[overall_df['Current_Quantity'] > 0][
         ['Symbol', 'Current_Quantity', 'Average_Buy_Price', 'Invested_Value', 'LTP', 
          'EMA9', 'EMA10', 'EMA11', 'EMA21', 'Current_Value', 'Unrealized_PnL']
    ].copy()
    
    import re
    def get_sl(row):
        sym = row['Symbol']
        avg_buy = row['Average_Buy_Price']
        ema21 = row['EMA21']
        
        if 'Tranches/Cheat' not in grouped_df.columns:
            return round(avg_buy * 0.9, 2)
            
        # Get all buys for this symbol
        buys = grouped_df[(grouped_df['Symbol'] == sym) & (grouped_df['Trade Type'] == 'buy')]
        
        tranch_nums = []
        for label in buys['Tranches/Cheat']:
            if pd.isna(label): continue
            m = re.match(r'Tranch\s+(\d+)', str(label))
            if m:
                tranch_nums.append(int(m.group(1)))
                
        if not tranch_nums:
            return round(avg_buy * 0.9, 2)
            
        max_tranch = max(tranch_nums)
        if max_tranch == 1:
            return round(avg_buy * 0.9, 2) # -10% from purchased price
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
        else: # > 3
            return round(ema21, 2)

    portfolio_df['SL'] = portfolio_df.apply(get_sl, axis=1)
    
    # Reorder portfolio_df
    port_cols = ['Symbol', 'Current_Quantity', 'Average_Buy_Price', 'SL', 'Invested_Value', 'LTP', 
                 'EMA9', 'EMA10', 'EMA11', 'EMA21', 'Current_Value', 'Unrealized_PnL']
    portfolio_df = portfolio_df[port_cols]
    
    return portfolio_df, overall_df

def process_tradebook(input_file, output_file, config_file=None):
    """
    Main execution function to load data, process sheets, and save to Excel.
    
    Args:
        input_file (str): Path to the input tradebook Excel file.
        output_file (str): Path to save the transformed Excel file.
        config_file (str, optional): Path to the config file for tranch logic.
    """
    df = load_data(input_file)
    if df is None:
        return
        
    config = load_config(config_file) if config_file else {}
        
    print("Processing data...")
    
    grouped_df = process_grouped_trades(df, config)
    portfolio_df, overall_df = calculate_portfolios(df, grouped_df)
    
    print(f"Saving transformed data to {output_file}...")
    try:
        # Use ExcelWriter to save multiple sheets
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            grouped_df.to_excel(writer, sheet_name='Transaction', index=False)
            portfolio_df.to_excel(writer, sheet_name='Current_Portfolio', index=False)
            overall_df.to_excel(writer, sheet_name='Overall_Portfolio', index=False)
            
            # Apply Excel native currency formatting (Number format)
            inr_format = '[$₹-en-IN] #,##0.00'
            
            worksheet_trans = writer.sheets['Transaction']
            for col_idx, col_name in enumerate(grouped_df.columns, 1):
                if col_name in ['Average_Price', 'Total_Value']:
                    for row in range(2, len(grouped_df) + 2):
                        worksheet_trans.cell(row=row, column=col_idx).number_format = inr_format
                        
            worksheet_port = writer.sheets['Current_Portfolio']
            for col_idx, col_name in enumerate(portfolio_df.columns, 1):
                if col_name in ['Average_Buy_Price', 'SL', 'Invested_Value', 'LTP', 'EMA9', 'EMA10', 'EMA11', 'EMA21', 'Current_Value', 'Unrealized_PnL']:
                    for row in range(2, len(portfolio_df) + 2):
                        worksheet_port.cell(row=row, column=col_idx).number_format = inr_format
                        
            worksheet_overall = writer.sheets['Overall_Portfolio']
            for col_idx, col_name in enumerate(overall_df.columns, 1):
                if col_name in ['Total_Buy_Value', 'Average_Buy_Price', 'Total_Sell_Value', 'Average_Sell_Price', 
                                'Invested_Value', 'LTP', 'EMA9', 'EMA10', 'EMA11', 'EMA21', 'Current_Value', 'Realized_PnL', 'Unrealized_PnL', 'Total_PnL']:
                    for row in range(2, len(overall_df) + 2):
                        worksheet_overall.cell(row=row, column=col_idx).number_format = inr_format
                elif col_name == 'Total_PnL_Percentage':
                    for row in range(2, len(overall_df) + 2):
                        # Format as percentage with 2 decimal places e.g., 15.50%
                        worksheet_overall.cell(row=row, column=col_idx).number_format = '0.00%'

        print("Done!")
    except Exception as e:
        print(f"Error saving {output_file}: {e}")

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(current_dir, 'Tradebook Template.xlsx')
    output_path = os.path.join(current_dir, 'Transformed_Tradebook.xlsx')
    config_path = os.path.join(current_dir, 'input.cfg')
    
    process_tradebook(input_path, output_path, config_path)
