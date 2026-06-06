"""
Stock Journal - Google Sheets Dashboard Builder (HARDENED MIRROR)
================================================================

The definitive solution for 1:1 parity. Fixes background issues, 
missing labels, and summation gaps in the Core vs Satellite breakdown.
"""

import pandas as pd
import gspread
import numpy as np
import logging
from gspread_formatting import *

logger = logging.getLogger(__name__)

def print(*args, **kwargs):
    msg = " ".join(str(arg) for arg in args)
    msg_upper = msg.upper()
    if "ERROR" in msg_upper or "FAILED" in msg_upper:
        logger.error(msg)
    elif "WARNING" in msg_upper or "NOTE:" in msg_upper:
        logger.warning(msg)
    else:
        logger.info(msg)

# Styling Constants
HEADER_BG = {"red": 47/255, "green": 84/255, "blue": 150/255}
HEADER_FG = {"red": 1.0, "green": 1.0, "blue": 1.0}
SECTION_FG = {"red": 47/255, "green": 84/255, "blue": 150/255}
GREEN_FG = {"red": 20/255, "green": 122/255, "blue": 30/255}
RED_FG = {"red": 192/255, "green": 0.0, "blue": 0.0}

INR_FMT = "₹#,##0.00"
PCT_FMT = "0.00%"
NUM_FMT = "#,##0"


def _grid_write(grid, start_cell, values):
    """
    Writes a 2D list of values, a 1D list, or a single scalar into our 
    in-memory grid starting at the specified A1 cell coordinate.
    """
    # 1. Parse Column and Row from cell reference (e.g., 'A1' or 'I12')
    col_str = ""
    row_str = ""
    for char in start_cell:
        if char.isalpha():
            col_str += char
        else:
            row_str += char
            
    col_idx = 0
    for char in col_str:
        col_idx = col_idx * 26 + (ord(char.upper()) - ord('A') + 1)
    col_idx -= 1
    row_idx = int(row_str) - 1

    # 2. Convert values input to standardized list-of-lists
    if not isinstance(values, list):
        values = [[values]]
    elif len(values) == 0:
        return
    elif not isinstance(values[0], list):
        values = [values]

    # 3. Write into the grid structure
    for r_offset, row_vals in enumerate(values):
        grid_row = row_idx + r_offset
        # Ensure row exists in grid
        while len(grid) <= grid_row:
            grid.append(["" for _ in range(13)])
            
        for c_offset, val in enumerate(row_vals):
            grid_col = col_idx + c_offset
            # Ensure column exists in that row
            while len(grid[grid_row]) <= grid_col:
                grid[grid_row].append("")
            
            # Format numbers to standard float/int or clean string for JSON compatibility
            if pd.isna(val):
                grid[grid_row][grid_col] = ""
            elif isinstance(val, (np.integer, int)):
                grid[grid_row][grid_col] = int(val)
            elif isinstance(val, (np.floating, float)):
                grid[grid_row][grid_col] = float(val)
            else:
                grid[grid_row][grid_col] = str(val)


def build_gsheet_dashboard(sh, portfolio_df, overall_df, benchmark_returns=None):
    print("  -> Rebuilding Hardened Mirror Dashboard...")
    
    try:
        ws = sh.worksheet("Dashboard")
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Dashboard", rows="300", cols="20")

    # GLOBAL RESET: Set background to white and text to black for the entire working area
    # This prevents any ghost formatting from previous runs from hiding data.
    formats = []
    default_format = cellFormat(
        backgroundColor={"red": 1.0, "green": 1.0, "blue": 1.0},
        textFormat=textFormat(foregroundColor={"red": 0.0, "green": 0.0, "blue": 0.0}, bold=False)
    )
    formats.append({"range": "A1:M250", "format": default_format})

    set_column_widths(ws, [
        ('A', 220), ('B', 130), ('C', 120), ('D', 120), ('E', 120), ('F', 120), ('G', 120),
        ('H', 40), # Spacer
        ('I', 220), ('J', 130), ('K', 120), ('L', 120), ('M', 120)
    ])

    # Initialize the in-memory grid (250 rows, 13 columns A-M)
    grid = [["" for _ in range(13)] for _ in range(250)]

    _grid_write(grid, 'A1', [['📊 Portfolio Dashboard']])
    formats.append({"range": "A1:G1", "format": cellFormat(textFormat=textFormat(bold=True, fontSize=18, foregroundColor=SECTION_FG), horizontalAlignment='CENTER')})
    
    if not overall_df.empty:
        last_date_str = f"🕒 Last Transacted Date: {pd.Timestamp.now().strftime('%d %b %Y')}"
        _grid_write(grid, 'J2', [[last_date_str]])
        formats.append({"range": "J2:M2", "format": cellFormat(textFormat=textFormat(italic=True, fontSize=10), horizontalAlignment='RIGHT')})

    # Left Side
    # Left Side & Middle Section
    row = 3
    _, formats = _write_performance_kpis_gsheet(grid, 3, formats)
    row, formats = _write_kpi_table(grid, portfolio_df, overall_df, row, formats)
    row += 2
    row, formats = _write_top_bottom_table(grid, portfolio_df, row, formats)
    row += 2
    row, formats = _write_nearest_sl_table(grid, portfolio_df, row, formats)
    row += 2
    row, formats = _write_corporate_actions(grid, portfolio_df, overall_df, row, formats)
    row += 2
    row, formats = _write_top_cheats_table(grid, portfolio_df, row, formats)

    # Right Side
    rrow = 3
    if benchmark_returns:
        rrow, formats = _write_benchmark_returns_table(grid, benchmark_returns, rrow, formats)
        rrow += 2
    rrow, formats = _write_cap_allocation(grid, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_core_satellite_distribution(grid, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_sector_allocation(grid, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_sector_allocation_core_satellite(grid, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_pnl_breakdown_by_cap(grid, overall_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_tranche_distribution(grid, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_holding_distribution(grid, portfolio_df, rrow, formats)

    # Write the entire grid in ONE single API call
    # Find actual populated height of the grid to avoid blank trailing rows
    max_row_populated = len(grid)
    while max_row_populated > 0 and all(x == "" for x in grid[max_row_populated - 1]):
        max_row_populated -= 1
        
    if max_row_populated > 0:
        grid_to_update = grid[:max_row_populated]
        ws.update(f'A1:M{max_row_populated}', grid_to_update)

    if formats:
        # Batch apply ALL formatting at once
        format_cell_ranges(ws, [(f['range'], f['format']) for f in formats])
    print("  -> Hardened Dashboard Complete.")


def _write_kpi_table(grid, p_df, o_df, r, formats):
    _grid_write(grid, f'A{r}', [['Portfolio Summary']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'A{r+1}', [['Metric', 'Value']])
    formats.append({"range": f'A{r+1}:B{r+1}', "format": _header_style()})

    ti = p_df['Invested_Value'].sum() if 'Invested_Value' in p_df.columns else 0
    tc = p_df['Current_Value'].sum() if 'Current_Value' in p_df.columns else 0
    tr = o_df['Realized_PnL'].sum() if 'Realized_PnL' in o_df.columns else 0
    tp = tr + (tc - ti)
    pp = tp / ti if ti > 0 else 0
    
    core_v = p_df[p_df['TF_Classification'] == 'Core']['Invested_Value'].sum() if 'TF_Classification' in p_df.columns else 0
    sat_v = ti - core_v

    data = [
        ['Total Invested Value', float(ti)], ['Total Current Value', float(tc)],
        ['Unrealized PnL', float(tc-ti)], ['Realized PnL', float(tr)],
        ['Combined PnL', float(tp)], ['Combined PnL %', float(pp)],
        ['Core Allocation (%)', float(core_v/ti if ti>0 else 0)],
        ['Satellite Allocation (%)', float(sat_v/ti if ti>0 else 0)],
        ['Active Holdings', int(len(p_df))],
        ['Total Stocks Traded', int(len(o_df))]
    ]
    _grid_write(grid, f'A{r+2}', data)
    
    formats.append({"range": f'B{r+2}:B{r+6}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'B{r+7}:B{r+9}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'B{r+10}:B{r+11}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern=NUM_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'B{r+4}:B{r+6}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if tp >= 0 else RED_FG, bold=True))})
    formats.append({"range": f'A{r+11}:B{r+11}', "format": cellFormat(backgroundColor=HEADER_BG, textFormat=textFormat(foregroundColor=HEADER_FG, bold=True), horizontalAlignment='RIGHT')})
    formats.append({"range": f'A{r+2}:B{r+11}', "format": cellFormat(borders=_thin_borders())})
    return r + 13, formats


def _write_top_bottom_table(grid, df, r, formats):
    _grid_write(grid, f'A{r}', [['Top 5 Gainers / Bottom 5 Losers']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'A{r+1}', [['Symbol', 'Classification', 'Invested (₹)', 'Current (₹)', 'PnL (₹)', 'PnL %']])
    formats.append({"range": f'A{r+1}:F{r+1}', "format": _header_style()})

    if df.empty: return r + 2, formats
    top5 = df.nlargest(5, 'Unrealized_PnL')
    bot5 = df.nsmallest(5, 'Unrealized_PnL')
    comb = pd.concat([top5, bot5]).drop_duplicates(subset='Symbol').sort_values('Unrealized_PnL', ascending=False)
    
    rows = []
    for _, s in comb.iterrows():
        inv = s.get('Invested_Value', 0)
        pnl = s.get('Unrealized_PnL', 0)
        rows.append([s.get('Symbol',''), s.get('TF_Classification','Satellite'), float(inv), float(s.get('Current_Value', 0)), float(pnl), float(pnl/inv if inv > 0 else 0)])
    
    if rows:
        end_r = r + 1 + len(rows)
        _grid_write(grid, f'A{r+2}', rows)
        formats.append({"range": f'C{r+2}:E{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'F{r+2}:F{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
        for i, row_data in enumerate(rows):
            p_val = row_data[4]
            formats.append({"range": f'A{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if p_val >= 0 else RED_FG, bold=True))})
            formats.append({"range": f'E{r+2+i}:F{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if p_val >= 0 else RED_FG))})
        formats.append({"range": f'A{r+2}:F{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_nearest_sl_table(grid, df, r, formats):
    _grid_write(grid, f'A{r}', [['⚠️ Stocks Nearest to Stop Loss']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=RED_FG))})
    _grid_write(grid, f'A{r+1}', [['Symbol', 'Classification', 'LTP', 'SL', 'Diff (₹)', 'Diff (%)', 'Tranche']])
    formats.append({"range": f'A{r+1}:G{r+1}', "format": _header_style()})
    if 'LTP_SL_Diff_Pct' not in df.columns or df.empty: return r + 2, formats
    near = df.nsmallest(10, 'LTP_SL_Diff_Pct')
    rows = [[s.get('Symbol',''), s.get('TF_Classification','Satellite'), float(s.get('LTP',0)), float(s.get('SL',0)), float(s.get('LTP_SL_Diff',0)), float(s.get('LTP_SL_Diff_Pct',0)), s.get('Latest_Tranche','')] for _,s in near.iterrows()]
    if rows:
        end_r = r + 1 + len(rows)
        _grid_write(grid, f'A{r+2}', rows)
        formats.append({"range": f'C{r+2}:E{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'F{r+2}:F{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'A{r+2}:G{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_benchmark_returns_table(grid, bench, r, formats):
    _grid_write(grid, f'I{r}', [['📈 Benchmark Returns (Since Invest Start)']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'I{r+1}', [['Index Tracker (ETF)', 'Start Price (₹)', 'LTP (₹)', 'Return %']])
    formats.append({"range": f'I{r+1}:L{r+1}', "format": _header_style()})
    rows = [[k, float(v.get('Start_Price',0)), float(v.get('LTP',0)), float(v.get('Return_Pct',0))] for k,v in bench.items()]
    if rows:
        end_r = r + 1 + len(rows)
        _grid_write(grid, f'I{r+2}', rows)
        formats.append({"range": f'J{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'L{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
        for i, rd in enumerate(rows):
            formats.append({"range": f'L{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if rd[3] >= 0 else RED_FG))})
        formats.append({"range": f'I{r+2}:L{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_cap_allocation(grid, df, r, formats):
    _grid_write(grid, f'I{r}', [['Allocation by Market Cap']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'I{r+1}', [['Market Cap', 'Invested (₹)', '% of Total', 'Returns']])
    formats.append({"range": f'I{r+1}:L{r+1}', "format": _header_style()})
    if 'Cap' not in df.columns: return r + 2, formats
    grouped = df.groupby('Cap', dropna=False).agg({'Invested_Value': 'sum', 'Unrealized_PnL': 'sum'}).sort_values('Invested_Value', ascending=False)
    total_i = grouped['Invested_Value'].sum()
    rows = [[str(k) if pd.notna(k) and k!='' else '', float(v['Invested_Value']), float(v['Invested_Value']/total_i if total_i > 0 else 0), float(v['Unrealized_PnL']/v['Invested_Value'] if v['Invested_Value'] > 0 else 0)] for k,v in grouped.iterrows()]
    rows.append(['TOTAL', float(total_i), 1.0, float(grouped['Unrealized_PnL'].sum()/total_i if total_i > 0 else 0)])
    end_r = r + 1 + len(rows)
    _grid_write(grid, f'I{r+2}', rows)
    formats.append({"range": f'J{r+2}:J{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'K{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:L{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:L{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_core_satellite_distribution(grid, df, r, formats):
    _grid_write(grid, f'I{r}', [['Core & Satellite Distribution']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'I{r+1}', [['Classification', 'Invested (₹)', '% of Total']])
    formats.append({"range": f'I{r+1}:K{r+1}', "format": _header_style()})
    grouped = df.groupby('TF_Classification').agg({'Invested_Value': 'sum'}).sort_values('Invested_Value', ascending=False)
    total = grouped['Invested_Value'].sum()
    rows = [[str(k), float(v['Invested_Value']), float(v['Invested_Value']/total if total > 0 else 0)] for k, v in grouped.iterrows()]
    rows.append(['TOTAL', float(total), 1.0])
    end_r = r + 1 + len(rows)
    _grid_write(grid, f'I{r+2}', rows)
    formats.append({"range": f'J{r+2}:J{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'K{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:K{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:K{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_sector_allocation(grid, df, r, formats):
    _grid_write(grid, f'I{r}', [['Allocation by Sector']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'I{r+1}', [['Sector', 'Invested (₹)', '% of Total']])
    formats.append({"range": f'I{r+1}:K{r+1}', "format": _header_style()})
    grouped = df.groupby('TF_Sector')['Invested_Value'].sum().sort_values(ascending=False)
    total = grouped.sum()
    rows = [[str(k), float(v), float(v/total if total > 0 else 0)] for k,v in grouped.items()]
    rows.append(['TOTAL', float(total), 1.0])
    end_r = r + 1 + len(rows)
    _grid_write(grid, f'I{r+2}', rows)
    formats.append({"range": f'J{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'K{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:K{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:K{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_sector_allocation_core_satellite(grid, df, r, formats):
    _grid_write(grid, f'I{r}', [['Sector Allocation (Core vs Satellite)']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'I{r+1}', [['Sector', 'Core (₹)', 'Satellite (₹)', 'Total (₹)']])
    formats.append({"range": f'I{r+1}:L{r+1}', "format": _header_style()})
    
    # Robust Pivot: Ensure all classifications are summed into Total even if labels mismatch
    pivot = df.pivot_table(index='TF_Sector', columns='TF_Classification', values='Invested_Value', aggfunc='sum').fillna(0)
    if 'Core' not in pivot.columns: pivot['Core'] = 0
    if 'Satellite' not in pivot.columns: pivot['Satellite'] = 0
    pivot['Total'] = pivot.sum(axis=1) # Sum all columns to avoid missing data
    pivot = pivot.sort_values('Total', ascending=False)
    
    rows = [[str(sector), float(row['Core']), float(row['Satellite']), float(row['Total'])] for sector, row in pivot.iterrows()]
    rows.append(['TOTAL', float(pivot['Core'].sum()), float(pivot['Satellite'].sum()), float(pivot['Total'].sum())])
    
    end_r = r + 1 + len(rows)
    _grid_write(grid, f'I{r+2}', rows)
    formats.append({"range": f'J{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:L{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:L{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_pnl_breakdown_by_cap(grid, o_df, r, formats):
    _grid_write(grid, f'I{r}', [['PnL Breakdown by Cap']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'I{r+1}', [['Cap', 'Realized PnL', 'Unrealized PnL', 'Total PnL']])
    formats.append({"range": f'I{r+1}:L{r+1}', "format": _header_style()})
    grouped = o_df.groupby('Cap', dropna=False).agg({'Realized_PnL': 'sum', 'Unrealized_PnL': 'sum'})
    grouped['Total'] = grouped['Realized_PnL'] + grouped['Unrealized_PnL']
    rows = [[str(cap) if pd.notna(cap) and cap!='' else '', float(v['Realized_PnL']), float(v['Unrealized_PnL']), float(v['Total'])] for cap, v in grouped.iterrows()]
    rows.append(['TOTAL', float(grouped['Realized_PnL'].sum()), float(grouped['Unrealized_PnL'].sum()), float(grouped['Total'].sum())])
    end_r = r + 1 + len(rows)
    _grid_write(grid, f'I{r+2}', rows)
    formats.append({"range": f'J{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    for i, rd in enumerate(rows):
        for ci in range(1, 4):
            val = rd[ci]
            formats.append({"range": f'{chr(74+ci-1)}{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if val >= 0 else RED_FG))})
    formats.append({"range": f'I{end_r}:L{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:L{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_tranche_distribution(grid, df, r, formats):
    _grid_write(grid, f'I{r}', [['Tranche Distribution']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'I{r+1}', [['Tranche', 'Core Count', 'Satellite Count', 'Total Count', '% of Holdings']])
    formats.append({"range": f'I{r+1}:M{r+1}', "format": _header_style()})

    if 'Latest_Tranche' not in df.columns or df.empty:
        return r + 2, formats

    # Filter out empty or null tranche values
    valid_df = df[df['Latest_Tranche'].notna() & df['Latest_Tranche'].str.strip().ne('')].copy()
    if valid_df.empty:
        return r + 2, formats

    # Group by Tranche and Classification
    tranche_groups = valid_df.groupby(['Latest_Tranche', 'TF_Classification']).size().unstack(fill_value=0)
    
    # Ensure both Core and Satellite columns exist
    if 'Core' not in tranche_groups.columns:
        tranche_groups['Core'] = 0
    if 'Satellite' not in tranche_groups.columns:
        tranche_groups['Satellite'] = 0
        
    tranche_groups['Total'] = tranche_groups['Core'] + tranche_groups['Satellite']
    tranche_groups = tranche_groups.sort_index()
    
    total_holdings = tranche_groups['Total'].sum()

    rows = []
    for label, r_data in tranche_groups.iterrows():
        rows.append([
            str(label),
            int(r_data['Core']),
            int(r_data['Satellite']),
            int(r_data['Total']),
            float(r_data['Total'] / total_holdings if total_holdings > 0 else 0)
        ])

    # Add TOTAL row
    rows.append([
        'TOTAL',
        int(tranche_groups['Core'].sum()),
        int(tranche_groups['Satellite'].sum()),
        int(total_holdings),
        1.0
    ])

    end_r = r + 1 + len(rows)
    _grid_write(grid, f'I{r+2}', rows)
    formats.append({"range": f'J{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern=NUM_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'M{r+2}:M{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:M{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:M{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_holding_distribution(grid, df, r, formats):
    _grid_write(grid, f'I{r}', [['Holding Period Distribution']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'I{r+1}', [['Period', 'Count', '% of Holdings']])
    formats.append({"range": f'I{r+1}:K{r+1}', "format": _header_style()})
    hp = df['Holding_Period']
    bk = [('0-30 days', int(((hp>=0)&(hp<=30)).sum())), ('31-90 days', int(((hp>30)&(hp<=90)).sum())), ('91-180 days', int(((hp>90)&(hp<=180)).sum())), ('180+ days', int((hp>180).sum()))]
    total = len(df)
    rows = [[label, cnt, float(cnt/total if total > 0 else 0)] for label, cnt in bk]
    _grid_write(grid, f'I{r+2}', rows)
    formats.append({"range": f'J{r+2}:J{r+5}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern=NUM_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'K{r+2}:K{r+5}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{r+2}:K{r+5}', "format": cellFormat(borders=_thin_borders())})
    return r + 7, formats


def _write_corporate_actions(grid, p_df, o_df, r, formats):
    _grid_write(grid, f'A{r}', [['🔔 Corporate Actions (Splits / Bonus)']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'A{r+1}', [['Symbol', 'Split / Bonus Details', 'Adj Required']])
    formats.append({"range": f'A{r+1}:C{r+1}', "format": _header_style()})
    comb = pd.concat([p_df, o_df]).drop_duplicates(subset='Symbol')
    acts = comb[comb['Split_Info'] != ''].sort_values('Symbol')
    rows = [[str(s.get('Symbol','')), s.get('Split_Info',''), s.get('Adj_Required','No')] for _,s in acts.iterrows()]
    if rows:
        end_r = r + 1 + len(rows)
        _grid_write(grid, f'A{r+2}', rows)
        formats.append({"range": f'A{r+2}:C{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats


def _write_top_cheats_table(grid, df, r, formats):
    _grid_write(grid, f'A{r}', [['Top 5 Cheat Stocks by Holding Period']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    _grid_write(grid, f'A{r+1}', [['Stock Name', 'Cheat', 'Holding Period (Days)', 'Invested Value', 'Current Value', 'Return %', 'XIRR']])
    formats.append({"range": f'A{r+1}:G{r+1}', "format": _header_style()})

    if 'Latest_Tranche' not in df.columns or df.empty:
        return r + 2, formats

    cheat_df = df[df['Latest_Tranche'].str.startswith('Cheat', na=False)].copy()
    if cheat_df.empty:
        _grid_write(grid, f'A{r+2}', [['No active cheat stocks in portfolio']])
        return r + 3, formats

    top_cheats = cheat_df.sort_values(by='Holding_Period', ascending=False).head(5)

    rows = []
    for _, s in top_cheats.iterrows():
        rows.append([
            s.get('Symbol', ''),
            s.get('Latest_Tranche', ''),
            int(s.get('Holding_Period', 0)),
            float(s.get('Invested_Value', 0)),
            float(s.get('Current_Value', 0)),
            float(s.get('Return_Pct', 0.0)),
            float(s.get('XIRR', 0.0))
        ])

    if rows:
        end_r = r + 1 + len(rows)
        _grid_write(grid, f'A{r+2}', rows)
        formats.append({"range": f'C{r+2}:C{end_r}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern='0'), horizontalAlignment='RIGHT')})
        formats.append({"range": f'D{r+2}:E{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'F{r+2}:G{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'A{r+2}:G{end_r}', "format": cellFormat(borders=_thin_borders())})
        for i in range(len(rows)):
            formats.append({"range": f'A{r+2+i}', "format": cellFormat(textFormat=textFormat(bold=True))})
            
            # Apply dynamic color highlight (Green/Red) on Return_Pct and XIRR
            ret_val = rows[i][5]
            xirr_val = rows[i][6]
            
            formats.append({"range": f'F{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if ret_val >= 0 else RED_FG))})
            formats.append({"range": f'G{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if xirr_val >= 0 else RED_FG))})

    return r + 2 + len(rows), formats


def _write_performance_kpis_gsheet(grid, start_row, formats):
    """Writes win/loss, risk/reward, and advancing/declining metrics side-by-side with KPIs."""
    # 1. Section Title
    _grid_write(grid, f'D{start_row}', [['Trading Performance']])
    formats.append({"range": f'D{start_row}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    
    # 2. Styled Header
    _grid_write(grid, f'D{start_row+1}', [['Metric', 'Realized (Closed)', 'Unrealized (Open)']])
    formats.append({"range": f'D{start_row+1}:F{start_row+1}', "format": _header_style()})
    
    r = start_row + 2
    # 3. Main KPI formulas
    # Note: Row references in formulas correspond to their row indices.
    data = [
        ['Winning Trades', '=COUNTIF(Overall_Portfolio!P$2:P$1000, ">0")', '=COUNTIF(Current_Portfolio!U$2:U$1000, ">0")'],
        ['Losing Trades', '=COUNTIF(Overall_Portfolio!P$2:P$1000, "<0")', '=COUNTIF(Current_Portfolio!U$2:U$1000, "<0")'],
        ['Win / Loss Ratio', f'=IF(E{r+1}>0, E{r}/E{r+1}, IF(E{r}>0, "No Loss", 0))', f'=IF(F{r+1}>0, F{r}/F{r+1}, IF(F{r}>0, "No Loss", 0))'],
        ['Avg Win (₹)', f'=IF(E{r}>0, SUMIF(Overall_Portfolio!P$2:P$1000, ">0")/E{r}, 0)', f'=IF(F{r}>0, SUMIF(Current_Portfolio!U$2:U$1000, ">0")/F{r}, 0)'],
        ['Avg Loss (₹)', f'=IF(E{r+1}>0, SUMIF(Overall_Portfolio!P$2:P$1000, "<0")/E{r+1}, 0)', f'=IF(F{r+1}>0, SUMIF(Current_Portfolio!U$2:U$1000, "<0")/F{r+1}, 0)'],
        ['Risk / Reward Ratio', f'=IF(E{r+4}<0, ABS(E{r+3}/E{r+4}), 0)', f'=IF(F{r+4}<0, ABS(F{r+3}/F{r+4}), 0)']
    ]
    
    _grid_write(grid, f'D{r}', data)
    
    # Formats for Main KPIs
    formats.append({"range": f'E{r}:F{r+1}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern=NUM_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'E{r+2}:F{r+2}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern='0.00'), horizontalAlignment='RIGHT')})
    formats.append({"range": f'E{r+3}:F{r+4}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'E{r+5}:F{r+5}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern='0.00'), horizontalAlignment='RIGHT')})
    formats.append({"range": f'D{r}:F{r+5}', "format": cellFormat(borders=_thin_borders())})
    
    r = r + 6
    r += 1 # blank spacer row
    
    # 4. Advancing / Declining Sub-headers
    _grid_write(grid, f'D{r}', [[' ', 'Advancing', 'Declining']])
    formats.append({"range": f'D{r}:F{r}', "format": _header_style()})
    
    # 5. Advancing / Declining counts using dynamic SUMPRODUCT formulas
    sub_metrics = [
        ['Core (Previous month close)',
         '=SUMPRODUCT((Current_Portfolio!D$2:D$1000="Core")*(Current_Portfolio!M$2:M$1000>Current_Portfolio!AA$2:AA$1000)*(Current_Portfolio!M$2:M$1000>0))',
         '=SUMPRODUCT((Current_Portfolio!D$2:D$1000="Core")*(Current_Portfolio!M$2:M$1000<=Current_Portfolio!AA$2:AA$1000)*(Current_Portfolio!M$2:M$1000>0))'],
        ['Satellite (Previous week close)',
         '=SUMPRODUCT((Current_Portfolio!D$2:D$1000="Satellite")*(Current_Portfolio!M$2:M$1000>Current_Portfolio!O$2:O$1000)*(Current_Portfolio!M$2:M$1000>0))',
         '=SUMPRODUCT((Current_Portfolio!D$2:D$1000="Satellite")*(Current_Portfolio!M$2:M$1000<=Current_Portfolio!O$2:O$1000)*(Current_Portfolio!M$2:M$1000>0))'],
        ['Core (Previous Close)',
         '=SUMPRODUCT((Current_Portfolio!D$2:D$1000="Core")*(Current_Portfolio!M$2:M$1000>Current_Portfolio!N$2:N$1000)*(Current_Portfolio!M$2:M$1000>0))',
         '=SUMPRODUCT((Current_Portfolio!D$2:D$1000="Core")*(Current_Portfolio!M$2:M$1000<=Current_Portfolio!N$2:N$1000)*(Current_Portfolio!M$2:M$1000>0))'],
        ['Satellite (Previous Close)',
         '=SUMPRODUCT((Current_Portfolio!D$2:D$1000="Satellite")*(Current_Portfolio!M$2:M$1000>Current_Portfolio!N$2:N$1000)*(Current_Portfolio!M$2:M$1000>0))',
         '=SUMPRODUCT((Current_Portfolio!D$2:D$1000="Satellite")*(Current_Portfolio!M$2:M$1000<=Current_Portfolio!N$2:N$1000)*(Current_Portfolio!M$2:M$1000>0))']
    ]
    
    _grid_write(grid, f'D{r+1}', sub_metrics)
    formats.append({"range": f'E{r+1}:F{r+4}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern=NUM_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'D{r+1}:F{r+4}', "format": cellFormat(borders=_thin_borders())})
    
    return r + 5, formats


def _header_style():
    return cellFormat(backgroundColor=HEADER_BG, textFormat=textFormat(foregroundColor=HEADER_FG, bold=True), horizontalAlignment='CENTER', borders=borders(top=border('SOLID'), bottom=border('SOLID'), left=border('SOLID'), right=border('SOLID')))


def _thin_borders():
    return borders(top=border('SOLID'), bottom=border('SOLID'), left=border('SOLID'), right=border('SOLID'))
