"""
Dashboard Module
================

Creates a Dashboard worksheet with summary KPIs and native Excel charts
using openpyxl. All chart data is written to columns S+ (hidden) so the
charts can reference real cell ranges.

Charts included:
    - Cap-wise Allocation (Pie)
    - Sector-wise Allocation (Pie)
    - Top 5 Gainers / Bottom 5 Losers (Bar)
    - Realized vs Unrealized PnL (Bar)
    - Tranche Distribution (Bar)
    - Holding Period Distribution (Bar)

Also produces:
    - Summary KPIs table (rows 1-10)
    - Stocks Nearest to Stop Loss table (rows 44-55)
"""

from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd


# ─── Styling Constants ──────────────────────────────────────────────
HEADER_FONT = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
HEADER_FILL = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
KPI_LABEL_FONT = Font(name='Calibri', bold=True, size=11)
KPI_VALUE_FONT = Font(name='Calibri', bold=True, size=13, color='2F5496')
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin'),
)
INR_FMT = '[$₹-en-IN] #,##0.00'
PCT_FMT = '0.00%'
DATA_COL_START = 19  # Column S — hidden data area for chart sources


def create_dashboard(wb, portfolio_df: pd.DataFrame, overall_df: pd.DataFrame) -> None:
    """
    Creates (or replaces) a 'Dashboard' sheet in the workbook with KPIs and charts.

    Args:
        wb:           An open openpyxl Workbook instance.
        portfolio_df: The Current Portfolio DataFrame.
        overall_df:   The Overall Portfolio DataFrame.
    """
    # Remove existing Dashboard sheet if present
    if 'Dashboard' in wb.sheetnames:
        del wb['Dashboard']

    ws = wb.create_sheet('Dashboard', 0)  # Insert as first sheet

    # ── 1. Summary KPIs ─────────────────────────────────────────────
    _write_kpis(ws, portfolio_df, overall_df)

    # ── 2. Cap-wise Allocation Pie ───────────────────────────────────
    data_row = _write_grouped_data(ws, portfolio_df, 'Cap', 'Invested_Value',
                                   DATA_COL_START, 1, 'Cap', 'Invested Value (₹)')
    _add_pie_chart(ws, 'Portfolio by Market Cap',
                   DATA_COL_START, 1, data_row, anchor='A12')

    # ── 3. Sector-wise Allocation Pie ────────────────────────────────
    data_row = _write_grouped_data(ws, portfolio_df, 'TF_Sector', 'Invested_Value',
                                   DATA_COL_START + 3, 1, 'Sector', 'Invested Value (₹)')
    _add_pie_chart(ws, 'Portfolio by Sector',
                   DATA_COL_START + 3, 1, data_row, anchor='J12')

    # ── 4. Top 5 Gainers / Bottom 5 Losers Bar ──────────────────────
    data_row = _write_top_bottom(ws, portfolio_df, 'Unrealized_PnL',
                                 DATA_COL_START + 6, 1)
    _add_bar_chart(ws, 'Top Gainers / Bottom Losers (Unrealized PnL)',
                   DATA_COL_START + 6, 1, data_row, anchor='A28')

    # ── 5. Realized vs Unrealized PnL Bar ────────────────────────────
    data_row = _write_pnl_comparison(ws, overall_df,
                                     DATA_COL_START + 9, 1)
    _add_double_bar_chart(ws, 'Realized vs Unrealized PnL',
                          DATA_COL_START + 9, 1, data_row, anchor='J28')

    # ── 6. Stocks Nearest to SL Table ────────────────────────────────
    _write_nearest_sl_table(ws, portfolio_df, start_row=44)

    # ── 7. Tranche Distribution Bar ──────────────────────────────────
    data_row = _write_grouped_count(ws, portfolio_df, 'Latest_Tranche',
                                    DATA_COL_START + 12, 1, 'Tranche', 'Count')
    _add_bar_chart(ws, 'Tranche Distribution',
                   DATA_COL_START + 12, 1, data_row, anchor='J44')

    # ── 8. Holding Period Distribution Bar ───────────────────────────
    data_row = _write_holding_buckets(ws, portfolio_df,
                                      DATA_COL_START + 15, 1)
    _add_bar_chart(ws, 'Holding Period Distribution',
                   DATA_COL_START + 15, 1, data_row, anchor='A58')

    # ── Hide data columns ────────────────────────────────────────────
    for c in range(DATA_COL_START, DATA_COL_START + 18):
        ws.column_dimensions[get_column_letter(c)].hidden = True

    # Set reasonable column widths for visible area
    for c in range(1, 18):
        ws.column_dimensions[get_column_letter(c)].width = 14

    print("Dashboard sheet created.")


# ═══════════════════════════════════════════════════════════════════
#  KPI Table
# ═══════════════════════════════════════════════════════════════════

def _write_kpis(ws, portfolio_df, overall_df):
    """Writes summary KPI metrics into rows 1-10, columns A-D."""
    # Title row
    ws.merge_cells('A1:D1')
    title_cell = ws['A1']
    title_cell.value = '📊 Portfolio Dashboard'
    title_cell.font = Font(name='Calibri', bold=True, size=16, color='2F5496')
    title_cell.alignment = Alignment(horizontal='center')

    # Calculate KPIs
    total_invested = portfolio_df['Invested_Value'].sum() if 'Invested_Value' in portfolio_df.columns else 0
    total_current = portfolio_df['Current_Value'].sum() if 'Current_Value' in portfolio_df.columns else 0
    total_unrealized = portfolio_df['Unrealized_PnL'].sum() if 'Unrealized_PnL' in portfolio_df.columns else 0
    total_realized = overall_df['Realized_PnL'].sum() if 'Realized_PnL' in overall_df.columns else 0
    total_pnl = total_realized + total_unrealized
    pnl_pct = total_pnl / total_invested if total_invested > 0 else 0
    active_count = len(portfolio_df)
    total_traded = len(overall_df)

    kpis = [
        ('Total Invested Value', total_invested, INR_FMT),
        ('Total Current Value', total_current, INR_FMT),
        ('Unrealized PnL', total_unrealized, INR_FMT),
        ('Realized PnL', total_realized, INR_FMT),
        ('Combined PnL', total_pnl, INR_FMT),
        ('Combined PnL %', pnl_pct, PCT_FMT),
        ('Active Holdings', active_count, '0'),
        ('Total Stocks Traded', total_traded, '0'),
    ]

    # Header row
    for col, header in enumerate(['Metric', 'Value'], 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
        cell.border = THIN_BORDER

    # Data rows
    for i, (label, value, fmt) in enumerate(kpis, 4):
        label_cell = ws.cell(row=i, column=1, value=label)
        label_cell.font = KPI_LABEL_FONT
        label_cell.border = THIN_BORDER

        value_cell = ws.cell(row=i, column=2, value=value)
        value_cell.font = KPI_VALUE_FONT
        value_cell.number_format = fmt
        value_cell.alignment = Alignment(horizontal='right')
        value_cell.border = THIN_BORDER


# ═══════════════════════════════════════════════════════════════════
#  Data Writers (hidden columns for chart sources)
# ═══════════════════════════════════════════════════════════════════

def _write_grouped_data(ws, df, group_col, value_col, col_start, row_start,
                        label_header, value_header):
    """Groups df by group_col, sums value_col, writes to hidden area. Returns last row."""
    if group_col not in df.columns or value_col not in df.columns:
        ws.cell(row=row_start, column=col_start, value=label_header)
        ws.cell(row=row_start, column=col_start + 1, value=value_header)
        ws.cell(row=row_start + 1, column=col_start, value='N/A')
        ws.cell(row=row_start + 1, column=col_start + 1, value=0)
        return row_start + 1

    grouped = df.groupby(group_col)[value_col].sum().sort_values(ascending=False)
    # Filter out empty labels
    grouped = grouped[grouped.index != '']

    ws.cell(row=row_start, column=col_start, value=label_header)
    ws.cell(row=row_start, column=col_start + 1, value=value_header)

    row = row_start + 1
    for label, val in grouped.items():
        ws.cell(row=row, column=col_start, value=str(label))
        ws.cell(row=row, column=col_start + 1, value=round(val, 2))
        row += 1

    return row - 1  # last data row


def _write_grouped_count(ws, df, group_col, col_start, row_start,
                         label_header, value_header):
    """Groups df by group_col, counts rows, writes to hidden area."""
    if group_col not in df.columns:
        ws.cell(row=row_start, column=col_start, value=label_header)
        ws.cell(row=row_start, column=col_start + 1, value=value_header)
        ws.cell(row=row_start + 1, column=col_start, value='N/A')
        ws.cell(row=row_start + 1, column=col_start + 1, value=0)
        return row_start + 1

    counts = df[group_col].value_counts().sort_index()
    counts = counts[counts.index != '']

    ws.cell(row=row_start, column=col_start, value=label_header)
    ws.cell(row=row_start, column=col_start + 1, value=value_header)

    row = row_start + 1
    for label, cnt in counts.items():
        ws.cell(row=row, column=col_start, value=str(label))
        ws.cell(row=row, column=col_start + 1, value=int(cnt))
        row += 1

    return row - 1


def _write_top_bottom(ws, df, value_col, col_start, row_start):
    """Writes top 5 gainers and bottom 5 losers by value_col."""
    ws.cell(row=row_start, column=col_start, value='Symbol')
    ws.cell(row=row_start, column=col_start + 1, value='PnL (₹)')

    if value_col not in df.columns or 'Symbol' not in df.columns:
        ws.cell(row=row_start + 1, column=col_start, value='N/A')
        ws.cell(row=row_start + 1, column=col_start + 1, value=0)
        return row_start + 1

    sorted_df = df.nlargest(5, value_col)
    bottom_df = df.nsmallest(5, value_col)
    combined = pd.concat([sorted_df, bottom_df]).drop_duplicates(subset='Symbol')
    combined = combined.sort_values(value_col, ascending=False)

    row = row_start + 1
    for _, r in combined.iterrows():
        ws.cell(row=row, column=col_start, value=r['Symbol'])
        ws.cell(row=row, column=col_start + 1, value=round(r[value_col], 2))
        row += 1

    return row - 1


def _write_pnl_comparison(ws, overall_df, col_start, row_start):
    """Writes Realized vs Unrealized PnL summary by cap category."""
    ws.cell(row=row_start, column=col_start, value='Category')
    ws.cell(row=row_start, column=col_start + 1, value='Realized PnL')
    ws.cell(row=row_start, column=col_start + 2, value='Unrealized PnL')

    if 'Cap' not in overall_df.columns:
        ws.cell(row=row_start + 1, column=col_start, value='Total')
        ws.cell(row=row_start + 1, column=col_start + 1, value=0)
        ws.cell(row=row_start + 1, column=col_start + 2, value=0)
        return row_start + 1

    grouped = overall_df.groupby('Cap').agg(
        Realized=('Realized_PnL', 'sum'),
        Unrealized=('Unrealized_PnL', 'sum')
    ).sort_index()

    row = row_start + 1
    for cat, vals in grouped.iterrows():
        ws.cell(row=row, column=col_start, value=str(cat))
        ws.cell(row=row, column=col_start + 1, value=round(vals['Realized'], 2))
        ws.cell(row=row, column=col_start + 2, value=round(vals['Unrealized'], 2))
        row += 1

    return row - 1


def _write_holding_buckets(ws, df, col_start, row_start):
    """Writes holding period distribution into buckets."""
    ws.cell(row=row_start, column=col_start, value='Period')
    ws.cell(row=row_start, column=col_start + 1, value='Count')

    if 'Holding_Period' not in df.columns:
        for i, label in enumerate(['0-30 days', '31-90 days', '91-180 days', '180+ days'], 1):
            ws.cell(row=row_start + i, column=col_start, value=label)
            ws.cell(row=row_start + i, column=col_start + 1, value=0)
        return row_start + 4

    hp = df['Holding_Period']
    buckets = [
        ('0-30 days', int(((hp >= 0) & (hp <= 30)).sum())),
        ('31-90 days', int(((hp > 30) & (hp <= 90)).sum())),
        ('91-180 days', int(((hp > 90) & (hp <= 180)).sum())),
        ('180+ days', int((hp > 180).sum())),
    ]

    row = row_start + 1
    for label, cnt in buckets:
        ws.cell(row=row, column=col_start, value=label)
        ws.cell(row=row, column=col_start + 1, value=cnt)
        row += 1

    return row - 1


# ═══════════════════════════════════════════════════════════════════
#  Risk Table (Nearest to Stop Loss)
# ═══════════════════════════════════════════════════════════════════

def _write_nearest_sl_table(ws, portfolio_df, start_row):
    """Writes a visible table of stocks closest to their stop loss."""
    # Title
    ws.merge_cells(f'A{start_row}:F{start_row}')
    title = ws.cell(row=start_row, column=1, value='⚠️ Stocks Nearest to Stop Loss')
    title.font = Font(name='Calibri', bold=True, size=13, color='C00000')

    headers = ['Symbol', 'LTP', 'SL', 'Diff (₹)', 'Diff (%)', 'Tranche']
    col_keys = ['Symbol', 'LTP', 'SL', 'LTP_SL_Diff', 'LTP_SL_Diff_Pct', 'Latest_Tranche']
    formats = [None, INR_FMT, INR_FMT, INR_FMT, PCT_FMT, None]

    header_row = start_row + 1
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
        cell.border = THIN_BORDER

    # Sort by LTP_SL_Diff_Pct ascending (most at risk first), take top 8
    if 'LTP_SL_Diff_Pct' not in portfolio_df.columns:
        ws.cell(row=header_row + 1, column=1, value='No SL data available')
        return

    nearest = portfolio_df.nsmallest(8, 'LTP_SL_Diff_Pct')

    row = header_row + 1
    for _, stock in nearest.iterrows():
        for c, (key, fmt) in enumerate(zip(col_keys, formats), 1):
            val = stock.get(key, '')
            cell = ws.cell(row=row, column=c, value=val)
            cell.border = THIN_BORDER
            if fmt:
                cell.number_format = fmt
            # Highlight red if diff < 5%
            if key == 'LTP_SL_Diff_Pct' and isinstance(val, (int, float)) and val < 0.05:
                cell.font = Font(color='C00000', bold=True)
        row += 1


# ═══════════════════════════════════════════════════════════════════
#  Chart Builders
# ═══════════════════════════════════════════════════════════════════

def _add_pie_chart(ws, title, col_start, row_start, last_data_row, anchor='A12'):
    """Adds a pie chart referencing data in hidden columns."""
    if last_data_row <= row_start:
        return  # No data

    chart = PieChart()
    chart.title = title
    chart.style = 10
    chart.width = 18
    chart.height = 12

    labels = Reference(ws, min_col=col_start, min_row=row_start + 1, max_row=last_data_row)
    data = Reference(ws, min_col=col_start + 1, min_row=row_start, max_row=last_data_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    # Show percentage labels
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCatName = True
    chart.dataLabels.showVal = False

    ws.add_chart(chart, anchor)


def _add_bar_chart(ws, title, col_start, row_start, last_data_row, anchor='A28'):
    """Adds a bar chart referencing data in hidden columns."""
    if last_data_row <= row_start:
        return

    chart = BarChart()
    chart.type = 'col'
    chart.title = title
    chart.style = 10
    chart.width = 18
    chart.height = 12
    chart.y_axis.numFmt = '#,##0'

    labels = Reference(ws, min_col=col_start, min_row=row_start + 1, max_row=last_data_row)
    data = Reference(ws, min_col=col_start + 1, min_row=row_start, max_row=last_data_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.shape = 4

    ws.add_chart(chart, anchor)


def _add_double_bar_chart(ws, title, col_start, row_start, last_data_row, anchor='J28'):
    """Adds a grouped bar chart with two data series (Realized + Unrealized)."""
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
