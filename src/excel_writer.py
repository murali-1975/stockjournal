"""
Excel Writer Module
===================

Handles saving DataFrames to the master Excel workbook with proper formatting.

Features:
    - Writes/replaces specific sheets (Raw_Tradebook, Transaction,
      Current_Portfolio, Overall_Portfolio) without destroying user-created
      custom sheets or charts.
    - Applies INR currency formatting (₹) to monetary columns.
    - Applies percentage formatting to PnL percentage columns.
    - Converts large integer ID columns from string back to native Excel
      numbers with '0' format to prevent scientific notation display.

Uses openpyxl engine with mode='a' (append) and if_sheet_exists='replace'
when the workbook already exists.
"""

import os

import pandas as pd

from .dashboard import create_dashboard


def save_workbook(
    df: pd.DataFrame,
    grouped_df: pd.DataFrame,
    portfolio_df: pd.DataFrame,
    overall_df: pd.DataFrame,
    output_file: str,
    benchmark_returns: dict = None
) -> None:
    """
    Saves all computed DataFrames to the master Excel workbook.

    Writes four sheets:
        - Raw_Tradebook:     The full, merged raw trade data.
        - Transaction:       Grouped and labeled trade summary.
        - Current_Portfolio:  Active holdings with LTP, EMAs, SL.
        - Overall_Portfolio:  Complete PnL report for all symbols.

    If the workbook already exists, sheets are replaced in-place (mode='a')
    so that any user-created custom sheets are preserved.

    Args:
        df:            The merged raw trade DataFrame.
        grouped_df:    The grouped transaction DataFrame.
        portfolio_df:  The current portfolio DataFrame.
        overall_df:    The overall portfolio DataFrame.
        output_file:   Path to the output Excel file.
    """
    print(f"Saving transformed data to {output_file}...")
    try:
        mode = 'a' if os.path.exists(output_file) else 'w'
        if_sheet_exists = 'replace' if mode == 'a' else None

        with pd.ExcelWriter(output_file, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer:
            df.to_excel(writer, sheet_name='Raw_Tradebook', index=False)
            grouped_df.to_excel(writer, sheet_name='Transaction', index=False)
            portfolio_df.to_excel(writer, sheet_name='Current_Portfolio', index=False)
            overall_df.to_excel(writer, sheet_name='Overall_Portfolio', index=False)

            # --- Format ID columns as flat integers (no scientific notation) ---
            _format_id_columns(writer, df, 'Raw_Tradebook')

            # --- Apply INR currency formatting ---
            inr_format = '[$₹-en-IN] #,##0.00'

            _apply_number_format(
                writer, grouped_df, 'Transaction',
                columns=['Average_Price', 'Total_Value'],
                number_format=inr_format
            )

            _apply_number_format(
                writer, portfolio_df, 'Current_Portfolio',
                columns=['Average_Buy_Price', 'SL', 'LTP_SL_Diff', 'Invested_Value', 'LTP',
                         'Prev_Week_Close', 'EMA9', 'EMA10', 'EMA11', 'EMA21',
                         'Current_Value', 'Unrealized_PnL'],
                number_format=inr_format
            )

            _apply_number_format(
                writer, overall_df, 'Overall_Portfolio',
                columns=['Total_Buy_Value', 'Average_Buy_Price', 'Total_Sell_Value',
                         'Average_Sell_Price', 'Invested_Value', 'LTP',
                         'Current_Value', 'Realized_PnL', 'Unrealized_PnL', 'Total_PnL'],
                number_format=inr_format
            )

            _apply_number_format(
                writer, overall_df, 'Overall_Portfolio',
                columns=['Total_PnL_Percentage'],
                number_format='0.00%'
            )

            _apply_number_format(
                writer, portfolio_df, 'Current_Portfolio',
                columns=['LTP_SL_Diff_Pct'],
                number_format='0.00%'
            )

            # --- Auto-fit columns and freeze panes for readability ---
            for sheet in ['Raw_Tradebook', 'Transaction', 'Current_Portfolio', 'Overall_Portfolio']:
                if sheet in writer.sheets:
                    _auto_fit_columns_and_freeze(writer, sheet)

            # --- Create Dashboard with charts ---
            create_dashboard(writer.book, portfolio_df, overall_df, df, benchmark_returns)

        print("Done!")
    except PermissionError:
        print(f"Error saving {output_file}: Permission denied. Please close the Excel file if it is open and try again.")
    except Exception as e:
        print(f"Error saving {output_file}: {e}")


def _format_id_columns(writer, df: pd.DataFrame, sheet_name: str) -> None:
    """
    Converts ID columns from string values to native Excel integers
    and applies '0' number format to prevent scientific notation.

    Args:
        writer:     The active pd.ExcelWriter instance.
        df:         The DataFrame whose columns to scan.
        sheet_name: The target worksheet name.
    """
    worksheet = writer.sheets[sheet_name]
    for col_idx, col_name in enumerate(df.columns, 1):
        if 'ID' in str(col_name).upper():
            for row in range(2, len(df) + 2):
                cell = worksheet.cell(row=row, column=col_idx)
                try:
                    if pd.notna(cell.value):
                        cell.value = int(float(cell.value))
                except Exception:
                    pass
                cell.number_format = '0'


def _apply_number_format(
    writer,
    df: pd.DataFrame,
    sheet_name: str,
    columns: list,
    number_format: str
) -> None:
    """
    Applies a specified Excel number format to selected columns in a worksheet.

    Args:
        writer:        The active pd.ExcelWriter instance.
        df:            The DataFrame being written to the sheet.
        sheet_name:    The target worksheet name.
        columns:       List of column names to format.
        number_format: The Excel number format string (e.g., '[$₹-en-IN] #,##0.00').
    """
    worksheet = writer.sheets[sheet_name]
    for col_idx, col_name in enumerate(df.columns, 1):
        if col_name in columns:
            for row in range(2, len(df) + 2):
                worksheet.cell(row=row, column=col_idx).number_format = number_format


def _auto_fit_columns_and_freeze(writer, sheet_name: str) -> None:
    """
    Adjusts the column widths dynamically based on the maximum string length
    in each column (including the header) and freezes the top row.
    Also formats the header row with bold text, borders, and centering.

    Args:
        writer:     The active pd.ExcelWriter instance.
        sheet_name: The target worksheet name.
    """
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    worksheet = writer.sheets[sheet_name]
    
    # Freeze the top header row
    worksheet.freeze_panes = 'A2'

    # Define Header Styles
    header_font = Font(bold=True, color='FFFFFF')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name (e.g., 'A')
        
        # Apply style to the first cell (header) in the column
        header_cell = col[0]
        header_cell.font = header_font
        header_cell.alignment = header_alignment
        header_cell.border = thin_border
        header_cell.fill = header_fill

        for cell in col:
            # Apply border to all data cells in the populated range
            if 1 < cell.row <= worksheet.max_row:
                cell.border = thin_border
                
            try:
                # Calculate length of the string representation
                if cell.value is not None:
                    if cell.number_format and ('%' in cell.number_format or '₹' in cell.number_format):
                        # Give extra padding for formatted numbers (currency symbols, commas, decimals)
                        length = len(str(cell.value)) + 5
                    else:
                        length = len(str(cell.value))
                    
                    if length > max_length:
                        max_length = length
            except:
                pass
        
        # Set a slightly padded width, bound between a min and max width
        adjusted_width = min(max_length + 2, 40)
        # Ensure a minimum width so headers aren't totally squished if empty
        adjusted_width = max(adjusted_width, 10)
        
        worksheet.column_dimensions[column].width = adjusted_width

