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


def save_workbook(
    df: pd.DataFrame,
    grouped_df: pd.DataFrame,
    portfolio_df: pd.DataFrame,
    overall_df: pd.DataFrame,
    output_file: str
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
                columns=['Average_Buy_Price', 'SL', 'Invested_Value', 'LTP',
                         'EMA9', 'EMA10', 'EMA11', 'EMA21',
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

            # --- Apply percentage formatting ---
            _apply_number_format(
                writer, overall_df, 'Overall_Portfolio',
                columns=['Total_PnL_Percentage'],
                number_format='0.00%'
            )

        print("Done!")
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
