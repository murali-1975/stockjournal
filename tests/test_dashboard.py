"""
Tests for src/dashboard.py
==========================

Validates the generation and formatting of the Excel Dashboard Worksheet.
"""

import unittest
import pandas as pd
import openpyxl

from src.dashboard import create_dashboard


class TestDashboard(unittest.TestCase):
    """Test suite for the create_dashboard function."""

    def setUp(self):
        """Create sample Portfolio and Overall DataFrames for testing."""
        self.wb = openpyxl.Workbook()
        
        # Current Portfolio Mock
        self.portfolio_df = pd.DataFrame({
            'Symbol': ['TCS', 'INFY'],
            'Cap': ['Large Cap', 'Large Cap'],
            'TF_Sector': ['IT', 'IT'],
            'TF_Classification': ['Core', 'Satellite'],
            'Latest_Tranche': ['Tranch 1', 'Cheat 1'],
            'Current_Quantity': [10.0, 5.0],
            'Average_Buy_Price': [3000.0, 1500.0],
            'Invested_Value': [30000.0, 7500.0],
            'LTP': [3200.0, 1600.0],
            'EMA9': [3150.0, 1580.0],
            'EMA10': [3140.0, 1570.0],
            'EMA11': [3130.0, 1560.0],
            'EMA21': [3100.0, 1500.0],
            'Current_Value': [32000.0, 8000.0],
            'Unrealized_PnL': [2000.0, 500.0],
            'Holding_Period': [30, 45],
            'Split_Info': ['', ''],
            'Adj_Required': ['No', 'No'],
            'SL': [2900.0, 1400.0]
        })
        # Add derived columns needed for dashboard logic
        self.portfolio_df['LTP_SL_Diff'] = self.portfolio_df['LTP'] - self.portfolio_df['SL']
        self.portfolio_df['LTP_SL_Diff_Pct'] = self.portfolio_df['LTP_SL_Diff'] / self.portfolio_df['SL']

        # Overall Portfolio Mock
        self.overall_df = pd.DataFrame({
            'Symbol': ['TCS', 'INFY', 'WIPRO'],
            'Cap': ['Large Cap', 'Large Cap', 'Large Cap'],
            'Realized_PnL': [1000.0, -500.0, 200.0],
            'Unrealized_PnL': [2000.0, 500.0, 0.0],
            'Split_Info': ['', '', '2:1 Split on 2024-01-01'],
            'Adj_Required': ['No', 'No', 'Yes']
        })

    def test_create_dashboard(self):
        """Dashboard should render without exceptions and create the required sheet."""
        
        # Add a mock watchlist dataframe
        watchlist_df = pd.DataFrame({
            'Stock': ['INFY', 'TCS', 'NEWSTOCK'],
            'Color': ['BLUE', 'GREEN', 'RED'],
            'Date': ['01-05-2026', '02-05-2026', '03-05-2026'],
            'Price': [1700.0, 3300.0, 500.0],
            'Previous week Close': [1600.0, 3200.0, 450.0]
        })
        
        # Execute the main function with the watchlist
        create_dashboard(self.wb, self.portfolio_df, self.overall_df, watchlist_df=watchlist_df)
        
        # Verify the sheet was created
        self.assertIn('Dashboard', self.wb.sheetnames)
        
        ws = self.wb['Dashboard']
        
        # Verify title is correct
        title_cell_value = ws['A1'].value
        self.assertEqual(title_cell_value, '📊 Portfolio Dashboard')
        
        # Verify column widths are set
        self.assertEqual(ws.column_dimensions['A'].width, 24)
        
        # We don't thoroughly test every cell location because formatting can change often,
        # but the function shouldn't raise any errors with sample data.

    def test_create_dashboard_empty_data(self):
        """Dashboard should gracefully handle empty DataFrames."""
        # Execute with empty data (and None for watchlist_df)
        create_dashboard(self.wb, pd.DataFrame(), pd.DataFrame())
        
        self.assertIn('Dashboard', self.wb.sheetnames)
        ws = self.wb['Dashboard']
        self.assertEqual(ws['A1'].value, '📊 Portfolio Dashboard')

    def test_corporate_actions_filter(self):
        """Corporate actions should only include current portfolio stocks."""
        # Add a split to a current portfolio stock to see if it's picked up
        self.portfolio_df.loc[0, 'Split_Info'] = '1:1 Bonus'
        create_dashboard(self.wb, self.portfolio_df, self.overall_df)
        ws = self.wb['Dashboard']
        
        # Check all cells to ensure WIPRO (which has a split but is only in overall_df) is NOT in the dashboard
        found_wipro = False
        found_tcs_split = False
        for row in ws.iter_rows(values_only=True):
            if row:
                if 'WIPRO' in row:
                    found_wipro = True
                if '1:1 Bonus' in row:
                    found_tcs_split = True
        
        self.assertFalse(found_wipro, "WIPRO from overall_df should not appear in corporate actions.")
        self.assertTrue(found_tcs_split, "TCS split from current portfolio should appear.")

    def test_stop_loss_excludes_core(self):
        """Stop loss dashboard should exclude Core stocks."""
        create_dashboard(self.wb, self.portfolio_df, self.overall_df)
        ws = self.wb['Dashboard']
        
        # The stop loss section is titled '⚠️ Stocks Nearest to Stop Loss'
        sl_section_started = False
        core_stock_found = False
        satellite_stock_found = False
        
        for row in ws.iter_rows(values_only=True):
            if row and row[0] == '⚠️ Stocks Nearest to Stop Loss':
                sl_section_started = True
                
            if sl_section_started and row and row[0]:
                if row[0] == 'TCS': # TCS is Core
                    core_stock_found = True
                if row[0] == 'INFY': # INFY is Satellite
                    satellite_stock_found = True
                    
        self.assertFalse(core_stock_found, "TCS (Core) should not be in the Stop Loss table.")
        self.assertTrue(satellite_stock_found, "INFY (Satellite) should be in the Stop Loss table.")

    def test_top_bottom_split(self):
        """Top 5/Bottom 5 should be split into Core and Satellite tables."""
        create_dashboard(self.wb, self.portfolio_df, self.overall_df)
        ws = self.wb['Dashboard']
        
        satellite_header_found = False
        core_header_found = False
        
        in_satellite_section = False
        in_core_section = False
        
        for row in ws.iter_rows(values_only=True):
            # Check for headers
            if row and row[0] == 'Top 5 Gainers / Bottom 5 Losers (Satellite)':
                satellite_header_found = True
                in_satellite_section = True
                in_core_section = False
            elif row and row[0] == 'Top 5 Gainers / Bottom 5 Losers (Core)':
                core_header_found = True
                in_core_section = True
                in_satellite_section = False
            elif row and row[0] and isinstance(row[0], str) and ('Nearest to Stop Loss' in row[0] or 'Corporate Actions' in row[0] or 'Top 5 Cheat' in row[0]):
                in_satellite_section = False
                in_core_section = False
                
            # Check data rows
            if row and len(row) >= 2:
                classification = row[1]
                if classification == 'Satellite':
                    self.assertFalse(in_core_section, "Satellite stock found in Core section")
                elif classification == 'Core':
                    self.assertFalse(in_satellite_section, "Core stock found in Satellite section")
                    
        self.assertTrue(satellite_header_found, "Satellite Top/Bottom table header missing")
        self.assertTrue(core_header_found, "Core Top/Bottom table header missing")

if __name__ == '__main__':
    unittest.main()
