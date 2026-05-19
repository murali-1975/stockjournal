"""
Stock Journal - Google Sheets Dashboard Builder (HARDENED MIRROR)
================================================================

The definitive solution for 1:1 parity. Fixes background issues, 
missing labels, and summation gaps in the Core vs Satellite breakdown.
"""

import pandas as pd
import gspread
import numpy as np
from gspread_formatting import *

# Styling Constants
HEADER_BG = {"red": 47/255, "green": 84/255, "blue": 150/255}
HEADER_FG = {"red": 1.0, "green": 1.0, "blue": 1.0}
SECTION_FG = {"red": 47/255, "green": 84/255, "blue": 150/255}
GREEN_FG = {"red": 20/255, "green": 122/255, "blue": 30/255}
RED_FG = {"red": 192/255, "green": 0.0, "blue": 0.0}

INR_FMT = "₹#,##0.00"
PCT_FMT = "0.00%"
NUM_FMT = "#,##0"

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

    ws.update('A1', [['📊 Portfolio Dashboard']])
    formats.append({"range": "A1:G1", "format": cellFormat(textFormat=textFormat(bold=True, fontSize=18, foregroundColor=SECTION_FG), horizontalAlignment='CENTER')})
    
    if not overall_df.empty:
        ws.update('J2', [[f"🕒 Last Transacted Date: {pd.Timestamp.now().strftime('%d %b %Y')}"]])
        formats.append({"range": "J2:M2", "format": cellFormat(textFormat=textFormat(italic=True, fontSize=10), horizontalAlignment='RIGHT')})

    # Left Side
    row = 3
    row, formats = _write_kpi_table(ws, portfolio_df, overall_df, row, formats)
    row += 2
    row, formats = _write_top_bottom_table(ws, portfolio_df, row, formats)
    row += 2
    row, formats = _write_nearest_sl_table(ws, portfolio_df, row, formats)
    row += 2
    row, formats = _write_corporate_actions(ws, portfolio_df, overall_df, row, formats)
    row += 2
    row, formats = _write_top_cheats_table(ws, portfolio_df, row, formats)

    # Right Side
    rrow = 3
    if benchmark_returns:
        rrow, formats = _write_benchmark_returns_table(ws, benchmark_returns, rrow, formats)
        rrow += 2
    rrow, formats = _write_cap_allocation(ws, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_core_satellite_distribution(ws, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_sector_allocation(ws, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_sector_allocation_core_satellite(ws, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_pnl_breakdown_by_cap(ws, overall_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_tranche_distribution(ws, portfolio_df, rrow, formats)
    rrow += 2
    rrow, formats = _write_holding_distribution(ws, portfolio_df, rrow, formats)

    if formats:
        # Batch apply ALL formatting at once
        format_cell_ranges(ws, [(f['range'], f['format']) for f in formats])
    print("  -> Hardened Dashboard Complete.")

def _write_kpi_table(ws, p_df, o_df, r, formats):
    ws.update(f'A{r}', [['Portfolio Summary']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'A{r+1}:B{r+1}', [['Metric', 'Value']])
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
    ws.update(f'A{r+2}:B{r+11}', data)
    
    formats.append({"range": f'B{r+2}:B{r+6}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'B{r+7}:B{r+9}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'B{r+10}:B{r+11}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern=NUM_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'B{r+4}:B{r+6}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if tp >= 0 else RED_FG, bold=True))})
    formats.append({"range": f'A{r+11}:B{r+11}', "format": cellFormat(backgroundColor=HEADER_BG, textFormat=textFormat(foregroundColor=HEADER_FG, bold=True), horizontalAlignment='RIGHT')})
    formats.append({"range": f'A{r+2}:B{r+11}', "format": cellFormat(borders=_thin_borders())})
    return r + 13, formats

def _write_top_bottom_table(ws, df, r, formats):
    ws.update(f'A{r}', [['Top 5 Gainers / Bottom 5 Losers']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'A{r+1}:F{r+1}', [['Symbol', 'Classification', 'Invested (₹)', 'Current (₹)', 'PnL (₹)', 'PnL %']])
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
        ws.update(f'A{r+2}:F{end_r}', rows)
        formats.append({"range": f'C{r+2}:E{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'F{r+2}:F{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
        for i, row_data in enumerate(rows):
            p_val = row_data[4]
            formats.append({"range": f'A{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if p_val >= 0 else RED_FG, bold=True))})
            formats.append({"range": f'E{r+2+i}:F{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if p_val >= 0 else RED_FG))})
        formats.append({"range": f'A{r+2}:F{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_nearest_sl_table(ws, df, r, formats):
    ws.update(f'A{r}', [['⚠️ Stocks Nearest to Stop Loss']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=RED_FG))})
    ws.update(f'A{r+1}:G{r+1}', [['Symbol', 'Classification', 'LTP', 'SL', 'Diff (₹)', 'Diff (%)', 'Tranche']])
    formats.append({"range": f'A{r+1}:G{r+1}', "format": _header_style()})
    if 'LTP_SL_Diff_Pct' not in df.columns or df.empty: return r + 2, formats
    near = df.nsmallest(10, 'LTP_SL_Diff_Pct')
    rows = [[s.get('Symbol',''), s.get('TF_Classification','Satellite'), float(s.get('LTP',0)), float(s.get('SL',0)), float(s.get('LTP_SL_Diff',0)), float(s.get('LTP_SL_Diff_Pct',0)), s.get('Latest_Tranche','')] for _,s in near.iterrows()]
    if rows:
        end_r = r + 1 + len(rows)
        ws.update(f'A{r+2}:G{end_r}', rows)
        formats.append({"range": f'C{r+2}:E{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'F{r+2}:F{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'A{r+2}:G{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_benchmark_returns_table(ws, bench, r, formats):
    ws.update(f'I{r}', [['📈 Benchmark Returns (Since Invest Start)']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'I{r+1}:L{r+1}', [['Index Tracker (ETF)', 'Start Price (₹)', 'LTP (₹)', 'Return %']])
    formats.append({"range": f'I{r+1}:L{r+1}', "format": _header_style()})
    rows = [[k, float(v.get('Start_Price',0)), float(v.get('LTP',0)), float(v.get('Return_Pct',0))] for k,v in bench.items()]
    if rows:
        end_r = r + 1 + len(rows)
        ws.update(f'I{r+2}:L{end_r}', rows)
        formats.append({"range": f'J{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'L{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
        for i, rd in enumerate(rows):
            formats.append({"range": f'L{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if rd[3] >= 0 else RED_FG))})
        formats.append({"range": f'I{r+2}:L{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_cap_allocation(ws, df, r, formats):
    ws.update(f'I{r}', [['Allocation by Market Cap']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'I{r+1}:L{r+1}', [['Market Cap', 'Invested (₹)', '% of Total', 'Returns']])
    formats.append({"range": f'I{r+1}:L{r+1}', "format": _header_style()})
    if 'Cap' not in df.columns: return r + 2, formats
    grouped = df.groupby('Cap', dropna=False).agg({'Invested_Value': 'sum', 'Unrealized_PnL': 'sum'}).sort_values('Invested_Value', ascending=False)
    total_i = grouped['Invested_Value'].sum()
    rows = [[str(k) if pd.notna(k) and k!='' else '', float(v['Invested_Value']), float(v['Invested_Value']/total_i if total_i > 0 else 0), float(v['Unrealized_PnL']/v['Invested_Value'] if v['Invested_Value'] > 0 else 0)] for k,v in grouped.iterrows()]
    rows.append(['TOTAL', float(total_i), 1.0, float(grouped['Unrealized_PnL'].sum()/total_i if total_i > 0 else 0)])
    end_r = r + 1 + len(rows)
    ws.update(f'I{r+2}:L{end_r}', rows)
    formats.append({"range": f'J{r+2}:J{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'K{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:L{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:L{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_core_satellite_distribution(ws, df, r, formats):
    ws.update(f'I{r}', [['Core & Satellite Distribution']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'I{r+1}:K{r+1}', [['Classification', 'Invested (₹)', '% of Total']])
    formats.append({"range": f'I{r+1}:K{r+1}', "format": _header_style()})
    grouped = df.groupby('TF_Classification').agg({'Invested_Value': 'sum'}).sort_values('Invested_Value', ascending=False)
    total = grouped['Invested_Value'].sum()
    rows = [[str(k), float(v['Invested_Value']), float(v['Invested_Value']/total if total > 0 else 0)] for k, v in grouped.iterrows()]
    rows.append(['TOTAL', float(total), 1.0])
    end_r = r + 1 + len(rows)
    ws.update(f'I{r+2}:K{end_r}', rows)
    formats.append({"range": f'J{r+2}:J{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'K{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:K{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:K{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_sector_allocation(ws, df, r, formats):
    ws.update(f'I{r}', [['Allocation by Sector']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'I{r+1}:K{r+1}', [['Sector', 'Invested (₹)', '% of Total']])
    formats.append({"range": f'I{r+1}:K{r+1}', "format": _header_style()})
    grouped = df.groupby('TF_Sector')['Invested_Value'].sum().sort_values(ascending=False)
    total = grouped.sum()
    rows = [[str(k), float(v), float(v/total if total > 0 else 0)] for k,v in grouped.items()]
    rows.append(['TOTAL', float(total), 1.0])
    end_r = r + 1 + len(rows)
    ws.update(f'I{r+2}:K{end_r}', rows)
    formats.append({"range": f'J{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'K{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:K{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:K{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_sector_allocation_core_satellite(ws, df, r, formats):
    ws.update(f'I{r}', [['Sector Allocation (Core vs Satellite)']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'I{r+1}:L{r+1}', [['Sector', 'Core (₹)', 'Satellite (₹)', 'Total (₹)']])
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
    ws.update(f'I{r+2}:L{end_r}', rows)
    formats.append({"range": f'J{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{end_r}:L{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:L{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_pnl_breakdown_by_cap(ws, o_df, r, formats):
    ws.update(f'I{r}', [['PnL Breakdown by Cap']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'I{r+1}:L{r+1}', [['Cap', 'Realized PnL', 'Unrealized PnL', 'Total PnL']])
    formats.append({"range": f'I{r+1}:L{r+1}', "format": _header_style()})
    grouped = o_df.groupby('Cap', dropna=False).agg({'Realized_PnL': 'sum', 'Unrealized_PnL': 'sum'})
    grouped['Total'] = grouped['Realized_PnL'] + grouped['Unrealized_PnL']
    rows = [[str(cap) if pd.notna(cap) and cap!='' else '', float(v['Realized_PnL']), float(v['Unrealized_PnL']), float(v['Total'])] for cap, v in grouped.iterrows()]
    rows.append(['TOTAL', float(grouped['Realized_PnL'].sum()), float(grouped['Unrealized_PnL'].sum()), float(grouped['Total'].sum())])
    end_r = r + 1 + len(rows)
    ws.update(f'I{r+2}:L{end_r}', rows)
    formats.append({"range": f'J{r+2}:L{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
    for i, rd in enumerate(rows):
        for ci in range(1, 4):
            val = rd[ci]
            formats.append({"range": f'{chr(74+ci-1)}{r+2+i}', "format": cellFormat(textFormat=textFormat(foregroundColor=GREEN_FG if val >= 0 else RED_FG))})
    formats.append({"range": f'I{end_r}:L{end_r}', "format": cellFormat(textFormat=textFormat(bold=True))})
    formats.append({"range": f'I{r+2}:L{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_tranche_distribution(ws, df, r, formats):
    ws.update(f'I{r}', [['Tranche Distribution']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'I{r+1}:K{r+1}', [['Tranche', 'Count', '% of Holdings']])
    formats.append({"range": f'I{r+1}:K{r+1}', "format": _header_style()})
    cnts = df['Latest_Tranche'].value_counts().sort_index()
    total = cnts.sum()
    rows = [[str(k), int(v), float(v/total if total > 0 else 0)] for k,v in cnts.items()]
    if rows:
        end_r = r + 1 + len(rows)
        ws.update(f'I{r+2}:K{end_r}', rows)
        formats.append({"range": f'J{r+2}:J{end_r}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern=NUM_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'K{r+2}:K{end_r}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'I{r+2}:K{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_holding_distribution(ws, df, r, formats):
    ws.update(f'I{r}', [['Holding Period Distribution']])
    formats.append({"range": f'I{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'I{r+1}:K{r+1}', [['Period', 'Count', '% of Holdings']])
    formats.append({"range": f'I{r+1}:K{r+1}', "format": _header_style()})
    hp = df['Holding_Period']
    bk = [('0-30 days', int(((hp>=0)&(hp<=30)).sum())), ('31-90 days', int(((hp>30)&(hp<=90)).sum())), ('91-180 days', int(((hp>90)&(hp<=180)).sum())), ('180+ days', int((hp>180).sum()))]
    total = len(df)
    rows = [[label, cnt, float(cnt/total if total > 0 else 0)] for label, cnt in bk]
    ws.update(f'I{r+2}:K{r+5}', rows)
    formats.append({"range": f'J{r+2}:J{r+5}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern=NUM_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'K{r+2}:K{r+5}', "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern=PCT_FMT), horizontalAlignment='RIGHT')})
    formats.append({"range": f'I{r+2}:K{r+5}', "format": cellFormat(borders=_thin_borders())})
    return r + 7, formats

def _write_corporate_actions(ws, p_df, o_df, r, formats):
    ws.update(f'A{r}', [['🔔 Corporate Actions (Splits / Bonus)']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'A{r+1}:C{r+1}', [['Symbol', 'Split / Bonus Details', 'Adj Required']])
    formats.append({"range": f'A{r+1}:C{r+1}', "format": _header_style()})
    comb = pd.concat([p_df, o_df]).drop_duplicates(subset='Symbol')
    acts = comb[comb['Split_Info'] != ''].sort_values('Symbol')
    rows = [[str(s.get('Symbol','')), s.get('Split_Info',''), s.get('Adj_Required','No')] for _,s in acts.iterrows()]
    if rows:
        end_r = r + 1 + len(rows)
        ws.update(f'A{r+2}:C{end_r}', rows)
        formats.append({"range": f'A{r+2}:C{end_r}', "format": cellFormat(borders=_thin_borders())})
    return r + 2 + len(rows), formats

def _write_top_cheats_table(ws, df, r, formats):
    ws.update(f'A{r}', [['Top 5 Cheat Stocks by Holding Period']])
    formats.append({"range": f'A{r}', "format": cellFormat(textFormat=textFormat(bold=True, fontSize=12, foregroundColor=SECTION_FG))})
    ws.update(f'A{r+1}:E{r+1}', [['Stock Name', 'Cheat', 'Holding Period (Days)', 'Invested Value', 'Current Value']])
    formats.append({"range": f'A{r+1}:E{r+1}', "format": _header_style()})

    if 'Latest_Tranche' not in df.columns or df.empty:
        return r + 2, formats

    cheat_df = df[df['Latest_Tranche'].str.startswith('Cheat', na=False)].copy()
    if cheat_df.empty:
        ws.update(f'A{r+2}', [['No active cheat stocks in portfolio']])
        return r + 3, formats

    top_cheats = cheat_df.sort_values(by='Holding_Period', ascending=False).head(5)

    rows = []
    for _, s in top_cheats.iterrows():
        rows.append([
            s.get('Symbol', ''),
            s.get('Latest_Tranche', ''),
            int(s.get('Holding_Period', 0)),
            float(s.get('Invested_Value', 0)),
            float(s.get('Current_Value', 0))
        ])

    if rows:
        end_r = r + 1 + len(rows)
        ws.update(f'A{r+2}:E{end_r}', rows)
        formats.append({"range": f'C{r+2}:C{end_r}', "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern='0'), horizontalAlignment='RIGHT')})
        formats.append({"range": f'D{r+2}:E{end_r}', "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern=INR_FMT), horizontalAlignment='RIGHT')})
        formats.append({"range": f'A{r+2}:E{end_r}', "format": cellFormat(borders=_thin_borders())})
        for i in range(len(rows)):
            formats.append({"range": f'A{r+2+i}', "format": cellFormat(textFormat=textFormat(bold=True))})

    return r + 2 + len(rows), formats

def _header_style():
    return cellFormat(backgroundColor=HEADER_BG, textFormat=textFormat(foregroundColor=HEADER_FG, bold=True), horizontalAlignment='CENTER', borders=borders(top=border('SOLID'), bottom=border('SOLID'), left=border('SOLID'), right=border('SOLID')))

def _thin_borders():
    return borders(top=border('SOLID'), bottom=border('SOLID'), left=border('SOLID'), right=border('SOLID'))
