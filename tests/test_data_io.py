"""
Tests for src/data_io.py
========================

Validates data loading, master database reading, and the merge/deduplication
logic that prevents double-counting of trades.
"""

import os
import tempfile
import unittest

import pandas as pd

from src.data_io import load_data, merge_and_deduplicate, load_price_updates


def _create_test_excel(data: dict, path: str) -> str:
    """Helper to create a temporary Excel file from a dict of columns."""
    df = pd.DataFrame(data)
    df.to_excel(path, index=False)
    return path


class TestLoadData(unittest.TestCase):
    """Test suite for the load_data function."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.valid_data = {
            'Trade Date': ['2025-01-01', '2025-01-02'],
            'Symbol': ['RELIANCE', 'TCS'],
            'Trade Type': ['buy', 'sell'],
            'Quantity': [10, 5],
            'Price': [2500.0, 3500.0],
            'Order ID': ['1200000047930385', '1000000036265725'],
        }

    def test_valid_file(self):
        """A valid Excel file with required columns should load successfully."""
        path = os.path.join(self.tmp_dir, 'valid.xlsx')
        _create_test_excel(self.valid_data, path)
        df = load_data(path)
        self.assertIsNotNone(df)
        self.assertEqual(len(df), 2)
        self.assertIn('Symbol', df.columns)

    def test_missing_file(self):
        """A missing file should return None."""
        df = load_data('/nonexistent/file.xlsx')
        self.assertIsNone(df)

    def test_missing_columns(self):
        """A file missing required columns should return None."""
        path = os.path.join(self.tmp_dir, 'bad.xlsx')
        bad_data = {'Col1': [1], 'Col2': [2]}
        _create_test_excel(bad_data, path)
        df = load_data(path)
        self.assertIsNone(df)

    def test_id_columns_as_strings(self):
        """ID columns should be read as strings to prevent scientific notation."""
        path = os.path.join(self.tmp_dir, 'ids.xlsx')
        _create_test_excel(self.valid_data, path)
        df = load_data(path)
        self.assertIsNotNone(df)
        # The Order ID should be a string type (object or StringDtype depending on pandas version)
        self.assertTrue(
            pd.api.types.is_string_dtype(df['Order ID']),
            f"Expected string dtype, got {df['Order ID'].dtype}"
        )
        self.assertEqual(str(df['Order ID'].iloc[0]), '1200000047930385')


class TestMergeAndDeduplicate(unittest.TestCase):
    """Test suite for the merge_and_deduplicate function."""

    def _make_df(self, n_rows=3, symbol='RELIANCE', start_qty=10):
        """Helper to create a small trade DataFrame."""
        return pd.DataFrame({
            'Trade Date': [f'2025-01-0{i+1}' for i in range(n_rows)],
            'Symbol': [symbol] * n_rows,
            'Trade Type': ['buy'] * n_rows,
            'Quantity': [start_qty + i for i in range(n_rows)],
            'Price': [100.0] * n_rows,
        })

    def test_both_none(self):
        """If both master and new are None, should return None."""
        result = merge_and_deduplicate(None, None)
        self.assertIsNone(result)

    def test_only_new(self):
        """If only new_df is provided, it should be returned as-is."""
        new_df = self._make_df(2)
        result = merge_and_deduplicate(None, new_df)
        self.assertEqual(len(result), 2)

    def test_only_master(self):
        """If only master_df is provided, it should be returned as-is."""
        master_df = self._make_df(3)
        result = merge_and_deduplicate(master_df, None)
        self.assertEqual(len(result), 3)

    def test_merge_no_duplicates(self):
        """Merging two non-overlapping DataFrames should concatenate them."""
        master = self._make_df(2, symbol='RELIANCE')
        new = self._make_df(2, symbol='TCS')
        result = merge_and_deduplicate(master, new)
        self.assertEqual(len(result), 4)

    def test_merge_with_duplicates(self):
        """Exact duplicate rows should be dropped after merging."""
        master = self._make_df(3)
        new = self._make_df(3)  # Same as master → all duplicates
        result = merge_and_deduplicate(master, new)
        self.assertEqual(len(result), 3)  # Duplicates removed

    def test_merge_partial_duplicates(self):
        """Only exact duplicates should be removed; unique rows must be kept."""
        master = self._make_df(2, start_qty=10)
        new = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-01-03'],
            'Symbol': ['RELIANCE', 'RELIANCE'],
            'Trade Type': ['buy', 'buy'],
            'Quantity': [10, 50],  # First row duplicates master, second is new
            'Price': [100.0, 100.0],
        })
        result = merge_and_deduplicate(master, new)
        self.assertEqual(len(result), 3)  # 2 original + 1 new unique


class TestLoadPriceUpdates(unittest.TestCase):
    """Test suite for the load_price_updates function."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()

    def test_missing_sheet_or_file(self):
        """Should return an empty dict if the file or sheet is missing."""
        res = load_price_updates('/nonexistent/file.xlsx')
        self.assertEqual(res, {})

    def test_valid_price_updates_different_col_positions(self):
        """Should successfully parse prices even if Symbol and LTP columns are rearranged."""
        path = os.path.join(self.tmp_dir, 'prices.xlsx')
        
        # Test case 1: [LTP, Unrelated, Symbol]
        data1 = {
            'LTP': [2500.50, 3500.00],
            'Unrelated': ['A', 'B'],
            'Symbol': ['RELIANCE', 'tcs'] # Should handle case-insensitivity in Symbol value
        }
        # Write custom sheet using pandas excel writer
        with pd.ExcelWriter(path) as writer:
            pd.DataFrame(data1).to_excel(writer, sheet_name='Price_Update', index=False)
            
        res = load_price_updates(path)
        self.assertEqual(res, {'RELIANCE': 2500.50, 'TCS': 3500.00})

    def test_valid_price_updates_alternate_headers(self):
        """Should recognize alternative column names like 'Stock Symbol' or 'Last Traded Price'."""
        path = os.path.join(self.tmp_dir, 'alt_prices.xlsx')
        data = {
            'Stock Symbol': ['INFY', 'WIPRO'],
            'Last Traded Price': [1600.0, 'invalid_price'], # Should skip invalid prices
            'LTP': [1600.0, 500.0] # 'Last Traded Price' header matches first
        }
        with pd.ExcelWriter(path) as writer:
            pd.DataFrame(data).to_excel(writer, sheet_name='Price_Update', index=False)
            
        res = load_price_updates(path)
        # INFY gets float, WIPRO skipped due to string 'invalid_price'
        self.assertEqual(res, {'INFY': 1600.0})


if __name__ == '__main__':
    unittest.main()
