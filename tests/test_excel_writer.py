"""
Tests for src/excel_writer.py
==============================

Validates that the Excel writer:
    - Creates all expected sheets
    - Applies INR currency formatting to monetary columns
    - Applies percentage formatting to PnL % columns
    - Converts large ID numbers to native integers (no scientific notation)
"""

import os
import tempfile
import unittest

import pandas as pd
from openpyxl import load_workbook

from src.excel_writer import save_workbook


class TestSaveWorkbook(unittest.TestCase):
    """Test suite for the save_workbook function."""

    def setUp(self):
        """Create sample DataFrames and a temporary output path."""
        self.tmp_dir = tempfile.mkdtemp()
        self.output_path = os.path.join(self.tmp_dir, 'test_output.xlsx')

        self.raw_df = pd.DataFrame({
            'Trade Date': ['2025-01-01'],
            'Symbol': ['RELIANCE'],
            'Trade Type': ['buy'],
            'Quantity': [10],
            'Price': [2500.0],
            'Order ID': ['1200000047930385'],
            'Trade ID': ['405271862'],
        })

        self.grouped_df = pd.DataFrame({
            'Trade Date': ['2025-01-01'],
            'Symbol': ['RELIANCE'],
            'Trade Type': ['buy'],
            'Total_Quantity': [10],
            'Average_Price': [2500.0],
            'Total_Value': [25000.0],
            'Tranches/Cheat': ['Tranch 1'],
        })

        self.portfolio_df = pd.DataFrame({
            'Symbol': ['RELIANCE'],
            'Cap': ['Large Cap'],
            'TF_Sector': ['Technology'],
            'TF_Classification': ['Core'],
            'Latest_Tranche': ['Tranch 1'],
            'Current_Quantity': [10],
            'Average_Buy_Price': [2500.0],
            'SL': [2250.0],
            'Invested_Value': [25000.0],
            'LTP': [2600.0],
            'EMA9': [2590.0],
            'EMA10': [2585.0],
            'EMA11': [2580.0],
            'EMA21': [2550.0],
            'Current_Value': [26000.0],
            'Unrealized_PnL': [1000.0],
            'Holding_Period': [30],
        })

        self.overall_df = pd.DataFrame({
            'Symbol': ['RELIANCE'],
            'Cap': ['Large Cap'],
            'TF_Sector': ['Technology'],
            'TF_Classification': ['Core'],
            'Latest_Tranche': ['Tranch 1'],
            'Total_Buy_Quantity': [10],
            'Total_Buy_Value': [25000.0],
            'Average_Buy_Price': [2500.0],
            'Total_Sell_Quantity': [0],
            'Total_Sell_Value': [0.0],
            'Average_Sell_Price': [0.0],
            'Current_Quantity': [10],
            'Invested_Value': [25000.0],
            'LTP': [2600.0],
            'Current_Value': [26000.0],
            'Realized_PnL': [0.0],
            'Unrealized_PnL': [1000.0],
            'Total_PnL': [1000.0],
            'Total_PnL_Percentage': [0.04],
            'Holding_Period': [30],
        })

    def test_sheets_created(self):
        """All four expected sheets should be created in the workbook."""
        save_workbook(self.raw_df, self.grouped_df, self.portfolio_df, self.overall_df, self.output_path)
        wb = load_workbook(self.output_path)
        expected_sheets = {'Raw_Tradebook', 'Transaction', 'Current_Portfolio', 'Overall_Portfolio'}
        self.assertTrue(expected_sheets.issubset(set(wb.sheetnames)))
        wb.close()

    def test_id_column_is_integer(self):
        """Order ID cells should contain native integers, not strings."""
        save_workbook(self.raw_df, self.grouped_df, self.portfolio_df, self.overall_df, self.output_path)
        wb = load_workbook(self.output_path)
        ws = wb['Raw_Tradebook']

        # Find Order ID column index
        header_row = [cell.value for cell in ws[1]]
        oid_col = header_row.index('Order ID') + 1

        cell_value = ws.cell(row=2, column=oid_col).value
        self.assertIsInstance(cell_value, int)
        self.assertEqual(cell_value, 1200000047930385)
        wb.close()

    def test_currency_format_applied(self):
        """Monetary columns in Transaction sheet should have INR currency format."""
        save_workbook(self.raw_df, self.grouped_df, self.portfolio_df, self.overall_df, self.output_path)
        wb = load_workbook(self.output_path)
        ws = wb['Transaction']

        header_row = [cell.value for cell in ws[1]]
        tv_col = header_row.index('Total_Value') + 1

        fmt = ws.cell(row=2, column=tv_col).number_format
        self.assertIn('₹', fmt)
        wb.close()

    def test_percentage_format_applied(self):
        """Total_PnL_Percentage in Overall_Portfolio should have percentage format."""
        save_workbook(self.raw_df, self.grouped_df, self.portfolio_df, self.overall_df, self.output_path)
        wb = load_workbook(self.output_path)
        ws = wb['Overall_Portfolio']

        header_row = [cell.value for cell in ws[1]]
        pct_col = header_row.index('Total_PnL_Percentage') + 1

        fmt = ws.cell(row=2, column=pct_col).number_format
        self.assertIn('%', fmt)
        wb.close()

    def test_append_mode_preserves_custom_sheets(self):
        """Running save_workbook twice should not destroy custom user sheets."""
        # First run: create the workbook
        save_workbook(self.raw_df, self.grouped_df, self.portfolio_df, self.overall_df, self.output_path)

        # Manually add a custom sheet
        wb = load_workbook(self.output_path)
        wb.create_sheet('My_Notes')
        wb['My_Notes']['A1'] = 'User custom data'
        wb.save(self.output_path)
        wb.close()

        # Second run: should replace data sheets but keep My_Notes
        save_workbook(self.raw_df, self.grouped_df, self.portfolio_df, self.overall_df, self.output_path)

        wb = load_workbook(self.output_path)
        self.assertIn('My_Notes', wb.sheetnames)
        self.assertEqual(wb['My_Notes']['A1'].value, 'User custom data')
        wb.close()


if __name__ == '__main__':
    unittest.main()
