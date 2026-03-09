"""
Dashboard Module
================

Creates a Dashboard worksheet with summary KPIs and native Excel charts
using openpyxl. Chart source data is placed in small tables directly
above or beside each chart for reliable rendering.

Charts included:
    - Cap-wise Allocation (Pie)
    - Sector-wise Allocation (Pie)
    - Top 5 Gainers / Bottom 5 Losers (Bar)
    - Realized vs Unrealized PnL (Bar)
    - Tranche Distribution (Bar)
    - Holding Period Distribution (Bar)

Also produces:
    - Summary KPIs table (rows 2-12)
    - Stocks Nearest to Stop Loss table
"""

from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd


# ─── Styling Constants ──────────────────────────────────────────────
HEADER_FONT = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
HEADER_FILL = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
SECTION_FONT = Font(name='Calibri', bold=True, size=12, color='2F5496')
KPI_LABEL_FONT = Font(name='Calibri', bold=True, size=11)
KPI_VALUE_FONT = Font(name='Calibri', bold=True, size=12, color='2F5496')
RED_FONT = Font(name='Calibri', bold=True, color='C00000')
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)
INR_FMT = '[$₹-en-IN] #,##0.00'
PCT_FMT = '0.00%'


def create_dashboard(wb, portfolio_df: pd.DataFrame, overall_df: pd.DataFrame) -> None:
    """
    Creates (or replaces) a 'Dashboard' sheet in the workbook with KPIs and charts.

    Args:
        wb:           An open openpyxl Workbook instance.
        portfolio_df: The Current Portfolio DataFrame.
        overall_df:   The Overall Portfolio DataFrame.
    """
    if 'Dashboard' in wb.sheetnames:
        del wb['Dashboard']

    ws = wb.create_sheet('Dashboard', 0)

    # ── Set column widths ────────────────────────────────────────────
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 18
    for c in range(3, 20):
        ws.column_dimensions[get_column_letter(c)].width = 14

    # ── Row tracker — we build the dashboard top-to-bottom ───────────
    row = 1

    # ══════════ Section 1: KPIs (rows 1-12) ══════════════════════════
    row = _write_kpis(ws, portfolio_df, overall_df, row)
    row += 2  # gap

    # ══════════ Section 2: Cap Allocation Pie (left) + Sector Pie (right)
    # Write data tables, then charts
    cap_data_start = row
    row = _write_section_header(ws, 'Portfolio Allocation', 1, row)
    cap_data_end = _write_grouped_table(ws, portfolio_df, 'Cap', 'Invested_Value',
                                        1, row, 'Market Cap', 'Invested (₹)')
    sector_data_end = _write_grouped_table(ws, portfolio_df, 'TF_Sector', 'Invested_Value',
                                           5, cap_data_start + 1, 'Sector', 'Invested (₹)')

    # Move row past both tables
    data_end = max(cap_data_end, sector_data_end)
    chart_row = data_end + 1

    _add_pie_chart(ws, 'By Market Cap', 1, cap_data_start + 1, cap_data_end,
                   anchor=f'A{chart_row}')
    _add_pie_chart(ws, 'By Sector', 5, cap_data_start + 1, sector_data_end,
                   anchor=f'H{chart_row}')
    row = chart_row + 16  # charts take ~15 rows

    # ══════════ Section 3: Top Gainers/Losers ════════════════════════
    row = _write_section_header(ws, 'Top Gainers / Bottom Losers', 1, row)
    gl_data_start = row
    gl_data_end = _write_top_bottom_table(ws, portfolio_df, 'Unrealized_PnL', 1, row)
    chart_row = gl_data_end + 1
    _add_bar_chart(ws, 'Unrealized PnL (₹)', 1, gl_data_start, gl_data_end,
                   anchor=f'A{chart_row}', width=24, height=13)
    row = chart_row + 16

    # ══════════ Section 4: Realized vs Unrealized PnL ════════════════
    row = _write_section_header(ws, 'Realized vs Unrealized PnL by Cap', 1, row)
    pnl_data_start = row
    pnl_data_end = _write_pnl_table(ws, overall_df, 1, row)
    chart_row = pnl_data_end + 1
    _add_double_bar_chart(ws, 'PnL Comparison (₹)', 1, pnl_data_start, pnl_data_end,
                          anchor=f'A{chart_row}')
    row = chart_row + 16

    # ══════════ Section 5: Stocks Nearest SL ═════════════════════════
    row = _write_nearest_sl_table(ws, portfolio_df, row)
    row += 2

    # ══════════ Section 6: Tranche Distribution ══════════════════════
    row = _write_section_header(ws, 'Tranche Distribution', 1, row)
    tr_data_start = row
    tr_data_end = _write_count_table(ws, portfolio_df, 'Latest_Tranche', 1, row,
                                     'Tranche', 'Count')
    chart_row = tr_data_end + 1
    _add_bar_chart(ws, 'Holdings per Tranche', 1, tr_data_start, tr_data_end,
                   anchor=f'A{chart_row}')

    # Holding Period Distribution (beside Tranche)
    hp_data_start = tr_data_start
    hp_header_row = hp_data_start - 1  # reuse the same section header row area
    hp_data_end = _write_holding_buckets(ws, portfolio_df, 5, tr_data_start)
    _add_bar_chart(ws, 'Holding Period (Days)', 5, hp_data_start, hp_data_end,
                   anchor=f'H{chart_row}')

    print("Dashboard sheet created.")


# ═══════════════════════════════════════════════════════════════════
#  KPI Table
# ═══════════════════════════════════════════════════════════════════

def _write_kpis(ws, portfolio_df, overall_df, start_row):
    """Writes summary KPI metrics. Returns the next free row."""
    # Title
    ws.merge_cells(f'A{start_row}:B{start_row}')
    title = ws.cell(row=start_row, column=1, value='📊 Portfolio Dashboard')
    title.font = Font(name='Calibri', bold=True, size=16, color='2F5496')
    title.alignment = Alignment(horizontal='center')

    hr = start_row + 1
    # blank row
    hr = start_row + 2
    for ci, h in enumerate(['Metric', 'Value'], 1):
        c = ws.cell(row=hr, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center')
        c.border = THIN_BORDER

    # KPI data
    ti = portfolio_df['Invested_Value'].sum() if 'Invested_Value' in portfolio_df.columns else 0
    tc = portfolio_df['Current_Value'].sum() if 'Current_Value' in portfolio_df.columns else 0
    tu = portfolio_df['Unrealized_PnL'].sum() if 'Unrealized_PnL' in portfolio_df.columns else 0
    tr = overall_df['Realized_PnL'].sum() if 'Realized_PnL' in overall_df.columns else 0
    tp = tr + tu
    pp = tp / ti if ti > 0 else 0

    kpis = [
        ('Total Invested Value', ti, INR_FMT),
        ('Total Current Value', tc, INR_FMT),
        ('Unrealized PnL', tu, INR_FMT),
        ('Realized PnL', tr, INR_FMT),
        ('Combined PnL', tp, INR_FMT),
        ('Combined PnL %', pp, PCT_FMT),
        ('Active Holdings', len(portfolio_df), '0'),
        ('Total Stocks Traded', len(overall_df), '0'),
    ]

    r = hr + 1
    for label, value, fmt in kpis:
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = KPI_LABEL_FONT
        lc.border = THIN_BORDER

        vc = ws.cell(row=r, column=2, value=value)
        vc.font = KPI_VALUE_FONT
        vc.number_format = fmt
        vc.alignment = Alignment(horizontal='right')
        vc.border = THIN_BORDER
        r += 1

    return r


# ═══════════════════════════════════════════════════════════════════
#  Data Table Writers
# ═══════════════════════════════════════════════════════════════════

def _write_section_header(ws, title, col, row):
    """Writes a bold section header and returns the next row."""
    c = ws.cell(row=row, column=col, value=title)
    c.font = SECTION_FONT
    return row + 1


def _write_grouped_table(ws, df, group_col, value_col, col_start, row_start,
                         label_header, value_header):
    """Groups df, sums values, writes as a mini table. Returns last data row."""
    # Headers
    hc1 = ws.cell(row=row_start, column=col_start, value=label_header)
    hc2 = ws.cell(row=row_start, column=col_start + 1, value=value_header)
    for c in (hc1, hc2):
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = THIN_BORDER

    if group_col not in df.columns or value_col not in df.columns:
        ws.cell(row=row_start + 1, column=col_start, value='N/A')
        ws.cell(row=row_start + 1, column=col_start + 1, value=0)
        return row_start + 1

    grouped = df.groupby(group_col)[value_col].sum().sort_values(ascending=False)
    grouped = grouped[grouped.index != '']

    r = row_start + 1
    for label, val in grouped.items():
        ws.cell(row=r, column=col_start, value=str(label)).border = THIN_BORDER
        vc = ws.cell(row=r, column=col_start + 1, value=round(val, 2))
        vc.number_format = INR_FMT
        vc.border = THIN_BORDER
        r += 1

    return r - 1


def _write_count_table(ws, df, group_col, col_start, row_start,
                       label_header, value_header):
    """Groups df by column, counts rows. Returns last data row."""
    hc1 = ws.cell(row=row_start, column=col_start, value=label_header)
    hc2 = ws.cell(row=row_start, column=col_start + 1, value=value_header)
    for c in (hc1, hc2):
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = THIN_BORDER

    if group_col not in df.columns:
        ws.cell(row=row_start + 1, column=col_start, value='N/A')
        ws.cell(row=row_start + 1, column=col_start + 1, value=0)
        return row_start + 1

    counts = df[group_col].value_counts().sort_index()
    counts = counts[counts.index != '']

    r = row_start + 1
    for label, cnt in counts.items():
        ws.cell(row=r, column=col_start, value=str(label)).border = THIN_BORDER
        ws.cell(row=r, column=col_start + 1, value=int(cnt)).border = THIN_BORDER
        r += 1

    return r - 1


def _write_top_bottom_table(ws, df, value_col, col_start, row_start):
    """Top 5 gainers + bottom 5 losers. Returns last data row."""
    hc1 = ws.cell(row=row_start, column=col_start, value='Symbol')
    hc2 = ws.cell(row=row_start, column=col_start + 1, value='PnL (₹)')
    for c in (hc1, hc2):
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = THIN_BORDER

    if value_col not in df.columns or 'Symbol' not in df.columns or len(df) == 0:
        ws.cell(row=row_start + 1, column=col_start, value='N/A')
        ws.cell(row=row_start + 1, column=col_start + 1, value=0)
        return row_start + 1

    top5 = df.nlargest(5, value_col)
    bot5 = df.nsmallest(5, value_col)
    combined = pd.concat([top5, bot5]).drop_duplicates(subset='Symbol')
    combined = combined.sort_values(value_col, ascending=False)

    r = row_start + 1
    for _, stock in combined.iterrows():
        ws.cell(row=r, column=col_start, value=stock['Symbol']).border = THIN_BORDER
        vc = ws.cell(row=r, column=col_start + 1, value=round(stock[value_col], 2))
        vc.number_format = INR_FMT
        vc.border = THIN_BORDER
        # Color code: green for gain, red for loss
        if stock[value_col] < 0:
            vc.font = RED_FONT
        r += 1

    return r - 1


def _write_pnl_table(ws, overall_df, col_start, row_start):
    """Realized vs Unrealized PnL grouped by Cap. Returns last data row."""
    headers = ['Category', 'Realized PnL', 'Unrealized PnL']
    for i, h in enumerate(headers):
        c = ws.cell(row=row_start, column=col_start + i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = THIN_BORDER

    if 'Cap' not in overall_df.columns:
        ws.cell(row=row_start + 1, column=col_start, value='Total')
        ws.cell(row=row_start + 1, column=col_start + 1, value=0)
        ws.cell(row=row_start + 1, column=col_start + 2, value=0)
        return row_start + 1

    grouped = overall_df.groupby('Cap').agg(
        Realized=('Realized_PnL', 'sum'),
        Unrealized=('Unrealized_PnL', 'sum')
    ).sort_index()

    r = row_start + 1
    for cat, vals in grouped.iterrows():
        ws.cell(row=r, column=col_start, value=str(cat)).border = THIN_BORDER
        rc = ws.cell(row=r, column=col_start + 1, value=round(vals['Realized'], 2))
        rc.number_format = INR_FMT
        rc.border = THIN_BORDER
        uc = ws.cell(row=r, column=col_start + 2, value=round(vals['Unrealized'], 2))
        uc.number_format = INR_FMT
        uc.border = THIN_BORDER
        r += 1

    return r - 1


def _write_holding_buckets(ws, df, col_start, row_start):
    """Writes holding period bucket counts. Returns last data row."""
    hc1 = ws.cell(row=row_start, column=col_start, value='Period')
    hc2 = ws.cell(row=row_start, column=col_start + 1, value='Count')
    for c in (hc1, hc2):
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = THIN_BORDER

    if 'Holding_Period' not in df.columns:
        buckets = [('0-30 days', 0), ('31-90 days', 0), ('91-180 days', 0), ('180+ days', 0)]
    else:
        hp = df['Holding_Period']
        buckets = [
            ('0-30 days', int(((hp >= 0) & (hp <= 30)).sum())),
            ('31-90 days', int(((hp > 30) & (hp <= 90)).sum())),
            ('91-180 days', int(((hp > 90) & (hp <= 180)).sum())),
            ('180+ days', int((hp > 180).sum())),
        ]

    r = row_start + 1
    for label, cnt in buckets:
        ws.cell(row=r, column=col_start, value=label).border = THIN_BORDER
        ws.cell(row=r, column=col_start + 1, value=cnt).border = THIN_BORDER
        r += 1

    return r - 1


# ═══════════════════════════════════════════════════════════════════
#  Risk Table (visible — Stocks Nearest to SL)
# ═══════════════════════════════════════════════════════════════════

def _write_nearest_sl_table(ws, portfolio_df, start_row):
    """Writes stocks closest to stop loss. Returns next free row."""
    title = ws.cell(row=start_row, column=1, value='⚠️ Stocks Nearest to Stop Loss')
    title.font = Font(name='Calibri', bold=True, size=13, color='C00000')

    headers = ['Symbol', 'LTP', 'SL', 'Diff (₹)', 'Diff (%)', 'Tranche']
    col_keys = ['Symbol', 'LTP', 'SL', 'LTP_SL_Diff', 'LTP_SL_Diff_Pct', 'Latest_Tranche']
    formats = [None, INR_FMT, INR_FMT, INR_FMT, PCT_FMT, None]

    hr = start_row + 1
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=hr, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center')
        c.border = THIN_BORDER

    if 'LTP_SL_Diff_Pct' not in portfolio_df.columns or len(portfolio_df) == 0:
        ws.cell(row=hr + 1, column=1, value='No data available')
        return hr + 2

    nearest = portfolio_df.nsmallest(8, 'LTP_SL_Diff_Pct')

    r = hr + 1
    for _, stock in nearest.iterrows():
        for ci, (key, fmt) in enumerate(zip(col_keys, formats), 1):
            val = stock.get(key, '')
            cell = ws.cell(row=r, column=ci, value=val)
            cell.border = THIN_BORDER
            if fmt:
                cell.number_format = fmt
            if key == 'LTP_SL_Diff_Pct' and isinstance(val, (int, float)) and val < 0.05:
                cell.font = RED_FONT
        r += 1

    return r + 1


# ═══════════════════════════════════════════════════════════════════
#  Chart Builders
# ═══════════════════════════════════════════════════════════════════

def _add_pie_chart(ws, title, col_start, row_start, last_data_row, anchor):
    """Adds a pie chart. Data: col_start=labels, col_start+1=values."""
    if last_data_row <= row_start:
        return

    chart = PieChart()
    chart.title = title
    chart.style = 10
    chart.width = 14
    chart.height = 12

    labels = Reference(ws, min_col=col_start, min_row=row_start + 1, max_row=last_data_row)
    data = Reference(ws, min_col=col_start + 1, min_row=row_start, max_row=last_data_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCatName = True
    chart.dataLabels.showVal = False

    ws.add_chart(chart, anchor)


def _add_bar_chart(ws, title, col_start, row_start, last_data_row,
                   anchor, width=18, height=12):
    """Adds a vertical bar chart."""
    if last_data_row <= row_start:
        return

    chart = BarChart()
    chart.type = 'col'
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height
    chart.y_axis.numFmt = '#,##0'

    labels = Reference(ws, min_col=col_start, min_row=row_start + 1, max_row=last_data_row)
    data = Reference(ws, min_col=col_start + 1, min_row=row_start, max_row=last_data_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.shape = 4

    ws.add_chart(chart, anchor)


def _add_double_bar_chart(ws, title, col_start, row_start, last_data_row, anchor):
    """Adds a grouped bar chart with two data series."""
    if last_data_row <= row_start:
        return

    chart = BarChart()
    chart.type = 'col'
    chart.grouping = 'clustered'
    chart.title = title
    chart.style = 10
    chart.width = 18
    chart.height = 12
    chart.y_axis.numFmt = '#,##0'

    labels = Reference(ws, min_col=col_start, min_row=row_start + 1, max_row=last_data_row)
    data = Reference(ws, min_col=col_start + 1, max_col=col_start + 2,
                     min_row=row_start, max_row=last_data_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    ws.add_chart(chart, anchor)
