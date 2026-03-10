"""
Dashboard Module
================

Creates a 'Dashboard' sheet with pre-computed summary tables that are
ready for the user to chart manually in Excel.

Tables included:
    1. Portfolio KPIs
    2. Cap-wise Allocation
    3. Core & Satellite Distribution
    4. Sector-wise Allocation
    5. Top 5 Gainers / Bottom 5 Losers
    6. Realized vs Unrealized PnL by Cap
    7. Stocks Nearest to Stop Loss
    8. Tranche Distribution
    9. Holding Period Distribution
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd


# ─── Styling Constants ──────────────────────────────────────────────
TITLE_FONT = Font(name='Calibri', bold=True, size=14, color='2F5496')
SECTION_FONT = Font(name='Calibri', bold=True, size=12, color='2F5496')
HEADER_FONT = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
HEADER_FILL = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
LABEL_FONT = Font(name='Calibri', bold=True, size=11)
VALUE_FONT = Font(name='Calibri', size=11)
GREEN_FONT = Font(name='Calibri', bold=True, color='147A1E')
RED_FONT = Font(name='Calibri', bold=True, color='C00000')
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)
INR_FMT = '[$₹-en-IN] #,##0.00'
PCT_FMT = '0.00%'


def create_dashboard(wb, portfolio_df: pd.DataFrame, overall_df: pd.DataFrame) -> None:
    """
    Creates (or replaces) a 'Dashboard' sheet with pre-computed summary tables.

    Args:
        wb:           An open openpyxl Workbook instance.
        portfolio_df: The Current Portfolio DataFrame.
        overall_df:   The Overall Portfolio DataFrame.
    """
    if 'Dashboard' in wb.sheetnames:
        del wb['Dashboard']

    ws = wb.create_sheet('Dashboard', 0)

    # Column widths
    widths = {'A': 24, 'B': 18, 'C': 18, 'D': 18, 'E': 18, 'F': 18,
              'G': 4,  # spacer
              'H': 24, 'I': 18, 'J': 18, 'K': 18}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ── Title ────────────────────────────────────────────────────────
    ws.merge_cells('A1:F1')
    ws['A1'].value = '📊 Portfolio Dashboard'
    ws['A1'].font = Font(name='Calibri', bold=True, size=18, color='2F5496')
    ws['A1'].alignment = Alignment(horizontal='center')

    # ══════════════════════════════════════════════════════════════════
    #  LEFT SIDE (Columns A-F) — KPIs, Gainers/Losers, Risk Table
    # ══════════════════════════════════════════════════════════════════

    row = 3
    row = _write_kpi_table(ws, portfolio_df, overall_df, row)
    row += 2
    row = _write_top_bottom_table(ws, portfolio_df, row)
    row += 2
    row = _write_nearest_sl_table(ws, portfolio_df, row)
    row += 2
    row = _write_corporate_actions(ws, portfolio_df, overall_df, row)

    # ══════════════════════════════════════════════════════════════════
    #  RIGHT SIDE (Columns H-K) — Allocation & Distribution tables
    # ══════════════════════════════════════════════════════════════════

    rrow = 3
    rrow = _write_cap_allocation(ws, portfolio_df, rrow)
    rrow += 2
    rrow = _write_classification_allocation(ws, portfolio_df, rrow)
    rrow += 2
    rrow = _write_sector_allocation(ws, portfolio_df, rrow)
    rrow += 2
    rrow = _write_core_satellite_sector_allocation(ws, portfolio_df, rrow)
    rrow += 2
    rrow = _write_pnl_by_cap(ws, overall_df, rrow)
    rrow += 2
    rrow = _write_tranche_distribution(ws, portfolio_df, rrow)
    rrow += 2
    rrow = _write_holding_distribution(ws, portfolio_df, rrow)

    print("Dashboard sheet created.")


# ═══════════════════════════════════════════════════════════════════
#  Helper: Write a styled table header row
# ═══════════════════════════════════════════════════════════════════

def _styled_header(ws, row, col_start, headers):
    """Writes a styled header row and returns the next row."""
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col_start + i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center')
        c.border = THIN_BORDER
    return row + 1


def _section_title(ws, row, col, title):
    """Writes a section title and returns the next row."""
    c = ws.cell(row=row, column=col, value=title)
    c.font = SECTION_FONT
    return row + 1


def _data_cell(ws, row, col, value, fmt=None, font=None):
    """Writes a formatted data cell."""
    c = ws.cell(row=row, column=col, value=value)
    c.border = THIN_BORDER
    if fmt:
        c.number_format = fmt
    if font:
        c.font = font
    return c


# ═══════════════════════════════════════════════════════════════════
#  Table 1: Portfolio KPIs
# ═══════════════════════════════════════════════════════════════════

def _write_kpi_table(ws, portfolio_df, overall_df, row):
    """Writes KPIs. Returns next free row."""
    row = _section_title(ws, row, 1, 'Portfolio Summary')
    row = _styled_header(ws, row, 1, ['Metric', 'Value'])

    ti = portfolio_df['Invested_Value'].sum() if 'Invested_Value' in portfolio_df.columns else 0
    tc = portfolio_df['Current_Value'].sum() if 'Current_Value' in portfolio_df.columns else 0
    tu = portfolio_df['Unrealized_PnL'].sum() if 'Unrealized_PnL' in portfolio_df.columns else 0
    tr = overall_df['Realized_PnL'].sum() if 'Realized_PnL' in overall_df.columns else 0
    tp = tr + tu
    pp = tp / ti if ti > 0 else 0

    core_ti = portfolio_df[portfolio_df['TF_Classification'].str.contains('Core', case=False, na=False)]['Invested_Value'].sum() if 'TF_Classification' in portfolio_df.columns else 0
    sat_ti = portfolio_df[portfolio_df['TF_Classification'].str.contains('Satellite', case=False, na=False)]['Invested_Value'].sum() if 'TF_Classification' in portfolio_df.columns else 0
    core_pct = core_ti / ti if ti > 0 else 0
    sat_pct = sat_ti / ti if ti > 0 else 0

    kpis = [
        ('Total Invested Value', ti, INR_FMT),
        ('Total Current Value', tc, INR_FMT),
        ('Unrealized PnL', tu, INR_FMT),
        ('Realized PnL', tr, INR_FMT),
        ('Combined PnL', tp, INR_FMT),
        ('Combined PnL %', pp, PCT_FMT),
        ('Core Allocation (%)', core_pct, PCT_FMT),
        ('Satellite Allocation (%)', sat_pct, PCT_FMT),
        ('Active Holdings', len(portfolio_df), '0'),
        ('Total Stocks Traded', len(overall_df), '0'),
    ]

    for label, value, fmt in kpis:
        _data_cell(ws, row, 1, label, font=LABEL_FONT)
        pnl_font = None
        if isinstance(value, (int, float)) and 'PnL' in label and value != 0:
            pnl_font = GREEN_FONT if value > 0 else RED_FONT
        _data_cell(ws, row, 2, value, fmt=fmt, font=pnl_font)
        row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 2: Top 5 Gainers / Bottom 5 Losers
# ═══════════════════════════════════════════════════════════════════

def _write_top_bottom_table(ws, df, row):
    """Writes top gainers and bottom losers. Returns next free row."""
    row = _section_title(ws, row, 1, 'Top 5 Gainers / Bottom 5 Losers')
    row = _styled_header(ws, row, 1, ['Symbol', 'Invested (₹)', 'Current (₹)', 'PnL (₹)', 'PnL %'])

    if 'Unrealized_PnL' not in df.columns or len(df) == 0:
        _data_cell(ws, row, 1, 'No data')
        return row + 1

    top5 = df.nlargest(5, 'Unrealized_PnL')
    bot5 = df.nsmallest(5, 'Unrealized_PnL')
    combined = pd.concat([top5, bot5]).drop_duplicates(subset='Symbol')
    combined = combined.sort_values('Unrealized_PnL', ascending=False)

    for _, s in combined.iterrows():
        _data_cell(ws, row, 1, s.get('Symbol', ''), font=LABEL_FONT)
        _data_cell(ws, row, 2, s.get('Invested_Value', 0), fmt=INR_FMT)
        _data_cell(ws, row, 3, s.get('Current_Value', 0), fmt=INR_FMT)

        pnl = s.get('Unrealized_PnL', 0)
        pnl_font = GREEN_FONT if pnl > 0 else RED_FONT if pnl < 0 else None
        _data_cell(ws, row, 4, pnl, fmt=INR_FMT, font=pnl_font)

        inv = s.get('Invested_Value', 0)
        pct = pnl / inv if inv > 0 else 0
        _data_cell(ws, row, 5, pct, fmt=PCT_FMT, font=pnl_font)
        row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 3: Stocks Nearest to Stop Loss
# ═══════════════════════════════════════════════════════════════════

def _write_nearest_sl_table(ws, df, row):
    """Writes stocks closest to SL. Returns next free row."""
    row = _section_title(ws, row, 1, '⚠️ Stocks Nearest to Stop Loss')
    row = _styled_header(ws, row, 1, ['Symbol', 'LTP', 'SL', 'Diff (₹)', 'Diff (%)', 'Tranche'])

    if 'LTP_SL_Diff_Pct' not in df.columns or len(df) == 0:
        _data_cell(ws, row, 1, 'No data')
        return row + 1

    nearest = df.nsmallest(10, 'LTP_SL_Diff_Pct')
    col_keys = ['Symbol', 'LTP', 'SL', 'LTP_SL_Diff', 'LTP_SL_Diff_Pct', 'Latest_Tranche']
    formats = [None, INR_FMT, INR_FMT, INR_FMT, PCT_FMT, None]

    for _, s in nearest.iterrows():
        for ci, (key, fmt) in enumerate(zip(col_keys, formats), 1):
            val = s.get(key, '')
            font = None
            if key == 'LTP_SL_Diff_Pct' and isinstance(val, (int, float)):
                font = RED_FONT if val < 0.05 else (GREEN_FONT if val > 0.15 else None)
            _data_cell(ws, row, ci, val, fmt=fmt, font=font)
        row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 4: Cap-wise Allocation (RIGHT SIDE)
# ═══════════════════════════════════════════════════════════════════

def _write_cap_allocation(ws, df, row):
    """Cap-wise allocation. Returns next free row."""
    row = _section_title(ws, row, 8, 'Allocation by Market Cap')
    row = _styled_header(ws, row, 8, ['Market Cap', 'Invested (₹)', '% of Total'])

    if 'Cap' not in df.columns or 'Invested_Value' not in df.columns:
        _data_cell(ws, row, 8, 'No data')
        return row + 1

    grouped = df.groupby('Cap')['Invested_Value'].sum().sort_values(ascending=False)
    grouped = grouped[grouped.index != '']
    total = grouped.sum()

    for label, val in grouped.items():
        _data_cell(ws, row, 8, str(label), font=LABEL_FONT)
        _data_cell(ws, row, 9, round(val, 2), fmt=INR_FMT)
        _data_cell(ws, row, 10, val / total if total > 0 else 0, fmt=PCT_FMT)
        row += 1

    # Total row
    _data_cell(ws, row, 8, 'TOTAL', font=LABEL_FONT)
    _data_cell(ws, row, 9, round(total, 2), fmt=INR_FMT, font=LABEL_FONT)
    _data_cell(ws, row, 10, 1.0, fmt=PCT_FMT, font=LABEL_FONT)
    row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 5: Core & Satellite Distribution (RIGHT SIDE)
# ═══════════════════════════════════════════════════════════════════

def _write_classification_allocation(ws, df, row):
    """Core & Satellite distribution. Returns next free row."""
    row = _section_title(ws, row, 8, 'Core & Satellite Distribution')
    row = _styled_header(ws, row, 8, ['Classification', 'Invested (₹)', '% of Total'])

    if 'TF_Classification' not in df.columns or 'Invested_Value' not in df.columns:
        _data_cell(ws, row, 8, 'No data')
        return row + 1

    grouped = df.groupby('TF_Classification')['Invested_Value'].sum().sort_values(ascending=False)
    grouped = grouped[grouped.index != '']
    total = grouped.sum()

    for label, val in grouped.items():
        _data_cell(ws, row, 8, str(label), font=LABEL_FONT)
        _data_cell(ws, row, 9, round(val, 2), fmt=INR_FMT)
        _data_cell(ws, row, 10, val / total if total > 0 else 0, fmt=PCT_FMT)
        row += 1

    _data_cell(ws, row, 8, 'TOTAL', font=LABEL_FONT)
    _data_cell(ws, row, 9, round(total, 2), fmt=INR_FMT, font=LABEL_FONT)
    _data_cell(ws, row, 10, 1.0, fmt=PCT_FMT, font=LABEL_FONT)
    row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 5: Sector-wise Allocation (RIGHT SIDE)
# ═══════════════════════════════════════════════════════════════════

def _write_sector_allocation(ws, df, row):
    """Sector-wise allocation. Returns next free row."""
    row = _section_title(ws, row, 8, 'Allocation by Sector')
    row = _styled_header(ws, row, 8, ['Sector', 'Invested (₹)', '% of Total'])

    if 'TF_Sector' not in df.columns or 'Invested_Value' not in df.columns:
        _data_cell(ws, row, 8, 'No data')
        return row + 1

    grouped = df.groupby('TF_Sector')['Invested_Value'].sum().sort_values(ascending=False)
    grouped = grouped[grouped.index != '']
    total = grouped.sum()

    for label, val in grouped.items():
        _data_cell(ws, row, 8, str(label), font=LABEL_FONT)
        _data_cell(ws, row, 9, round(val, 2), fmt=INR_FMT)
        _data_cell(ws, row, 10, val / total if total > 0 else 0, fmt=PCT_FMT)
        row += 1

    _data_cell(ws, row, 8, 'TOTAL', font=LABEL_FONT)
    _data_cell(ws, row, 9, round(total, 2), fmt=INR_FMT, font=LABEL_FONT)
    _data_cell(ws, row, 10, 1.0, fmt=PCT_FMT, font=LABEL_FONT)
    row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 5b: Core vs Satellite Sector Allocation (RIGHT SIDE)
# ═══════════════════════════════════════════════════════════════════

def _write_core_satellite_sector_allocation(ws, df, row):
    """Sector allocation split by Core and Satellite. Returns next free row."""
    row = _section_title(ws, row, 8, 'Sector Allocation (Core vs Satellite)')
    row = _styled_header(ws, row, 8, ['Sector', 'Core (₹)', 'Satellite (₹)', 'Total (₹)'])

    if 'TF_Sector' not in df.columns or 'TF_Classification' not in df.columns or 'Invested_Value' not in df.columns:
        _data_cell(ws, row, 8, 'No data')
        return row + 1

    valid_df = df[df['TF_Sector'] != ''].copy()
    if valid_df.empty:
        _data_cell(ws, row, 8, 'No sector data available')
        return row + 1

    core_mask = valid_df['TF_Classification'].str.contains('Core', case=False, na=False)
    sat_mask = valid_df['TF_Classification'].str.contains('Satellite', case=False, na=False)

    core_df = valid_df[core_mask]
    sat_df = valid_df[sat_mask]

    core_grouped = core_df.groupby('TF_Sector')['Invested_Value'].sum()
    sat_grouped = sat_df.groupby('TF_Sector')['Invested_Value'].sum()

    combined = pd.DataFrame({'Core': core_grouped, 'Satellite': sat_grouped}).fillna(0)
    combined['Total'] = combined['Core'] + combined['Satellite']
    combined = combined[combined['Total'] > 0]
    combined = combined.sort_values(by='Total', ascending=False)

    total_core = combined['Core'].sum()
    total_sat = combined['Satellite'].sum()
    total_all = combined['Total'].sum()

    for sector, row_data in combined.iterrows():
        _data_cell(ws, row, 8, str(sector), font=LABEL_FONT)
        _data_cell(ws, row, 9, round(row_data['Core'], 2), fmt=INR_FMT)
        _data_cell(ws, row, 10, round(row_data['Satellite'], 2), fmt=INR_FMT)
        _data_cell(ws, row, 11, round(row_data['Total'], 2), fmt=INR_FMT)
        row += 1

    _data_cell(ws, row, 8, 'TOTAL', font=LABEL_FONT)
    _data_cell(ws, row, 9, round(total_core, 2), fmt=INR_FMT, font=LABEL_FONT)
    _data_cell(ws, row, 10, round(total_sat, 2), fmt=INR_FMT, font=LABEL_FONT)
    _data_cell(ws, row, 11, round(total_all, 2), fmt=INR_FMT, font=LABEL_FONT)
    row += 1

    return row




# ═══════════════════════════════════════════════════════════════════
#  Table 6: Realized vs Unrealized PnL by Cap (RIGHT SIDE)
# ═══════════════════════════════════════════════════════════════════

def _write_pnl_by_cap(ws, overall_df, row):
    """PnL comparison by cap. Returns next free row."""
    row = _section_title(ws, row, 8, 'PnL Breakdown by Cap')
    row = _styled_header(ws, row, 8, ['Cap', 'Realized PnL', 'Unrealized PnL', 'Total PnL'])

    if 'Cap' not in overall_df.columns:
        _data_cell(ws, row, 8, 'No data')
        return row + 1

    grouped = overall_df.groupby('Cap').agg(
        Realized=('Realized_PnL', 'sum'),
        Unrealized=('Unrealized_PnL', 'sum')
    ).sort_index()

    grand_r = grand_u = 0
    for cat, vals in grouped.iterrows():
        r_val = round(vals['Realized'], 2)
        u_val = round(vals['Unrealized'], 2)
        t_val = r_val + u_val
        grand_r += r_val
        grand_u += u_val

        _data_cell(ws, row, 8, str(cat), font=LABEL_FONT)
        _data_cell(ws, row, 9, r_val, fmt=INR_FMT,
                   font=GREEN_FONT if r_val > 0 else (RED_FONT if r_val < 0 else None))
        _data_cell(ws, row, 10, u_val, fmt=INR_FMT,
                   font=GREEN_FONT if u_val > 0 else (RED_FONT if u_val < 0 else None))
        _data_cell(ws, row, 11, t_val, fmt=INR_FMT,
                   font=GREEN_FONT if t_val > 0 else (RED_FONT if t_val < 0 else None))
        row += 1

    # Total row
    grand_t = grand_r + grand_u
    _data_cell(ws, row, 8, 'TOTAL', font=LABEL_FONT)
    _data_cell(ws, row, 9, grand_r, fmt=INR_FMT, font=LABEL_FONT)
    _data_cell(ws, row, 10, grand_u, fmt=INR_FMT, font=LABEL_FONT)
    _data_cell(ws, row, 11, grand_t, fmt=INR_FMT, font=LABEL_FONT)
    row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 7: Tranche Distribution (RIGHT SIDE)
# ═══════════════════════════════════════════════════════════════════

def _write_tranche_distribution(ws, df, row):
    """Tranche distribution counts. Returns next free row."""
    row = _section_title(ws, row, 8, 'Tranche Distribution')
    row = _styled_header(ws, row, 8, ['Tranche', 'Count', '% of Holdings'])

    if 'Latest_Tranche' not in df.columns or len(df) == 0:
        _data_cell(ws, row, 8, 'No data')
        return row + 1

    counts = df['Latest_Tranche'].value_counts().sort_index()
    counts = counts[counts.index != '']
    total = counts.sum()

    for label, cnt in counts.items():
        _data_cell(ws, row, 8, str(label), font=LABEL_FONT)
        _data_cell(ws, row, 9, int(cnt))
        _data_cell(ws, row, 10, cnt / total if total > 0 else 0, fmt=PCT_FMT)
        row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 8: Holding Period Distribution (RIGHT SIDE)
# ═══════════════════════════════════════════════════════════════════

def _write_holding_distribution(ws, df, row):
    """Holding period bucket counts. Returns next free row."""
    row = _section_title(ws, row, 8, 'Holding Period Distribution')
    row = _styled_header(ws, row, 8, ['Period', 'Count', '% of Holdings'])

    total = len(df)
    if 'Holding_Period' not in df.columns or total == 0:
        _data_cell(ws, row, 8, 'No data')
        return row + 1

    hp = df['Holding_Period']
    buckets = [
        ('0 - 30 days', int(((hp >= 0) & (hp <= 30)).sum())),
        ('31 - 90 days', int(((hp > 30) & (hp <= 90)).sum())),
        ('91 - 180 days', int(((hp > 90) & (hp <= 180)).sum())),
        ('180+ days', int((hp > 180).sum())),
    ]

    for label, cnt in buckets:
        _data_cell(ws, row, 8, label, font=LABEL_FONT)
        _data_cell(ws, row, 9, cnt)
        _data_cell(ws, row, 10, cnt / total if total > 0 else 0, fmt=PCT_FMT)
        row += 1

    return row


# ═══════════════════════════════════════════════════════════════════
#  Table 10: Corporate Actions (LEFT SIDE)
# ═══════════════════════════════════════════════════════════════════

def _write_corporate_actions(ws, portfolio_df, overall_df, row):
    """Writes stocks with splits/bonuses since purchase. Returns next free row."""
    row = _section_title(ws, row, 1, '🔔 Corporate Actions (Splits / Bonus)')
    row = _styled_header(ws, row, 1, ['Symbol', 'Split / Bonus Details', 'Adj Required'])

    # Combine both portfolios and filter for rows with split info
    all_stocks = pd.concat([portfolio_df, overall_df]).drop_duplicates(subset='Symbol')

    if 'Split_Info' not in all_stocks.columns or len(all_stocks) == 0:
        _data_cell(ws, row, 1, 'No corporate actions detected')
        return row + 1

    actions = all_stocks[all_stocks['Split_Info'] != ''].sort_values('Symbol')

    if len(actions) == 0:
        _data_cell(ws, row, 1, 'No splits/bonuses detected since purchase')
        return row + 1

    for _, s in actions.iterrows():
        _data_cell(ws, row, 1, s.get('Symbol', ''), font=LABEL_FONT)
        _data_cell(ws, row, 2, s.get('Split_Info', ''))
        adj = s.get('Adj_Required', 'No')
        _data_cell(ws, row, 3, adj,
                   font=RED_FONT if adj == 'Yes' else None)
        row += 1

    return row
