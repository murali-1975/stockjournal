"""
Tests for src/corporate_actions.py
==================================

Validates the Option B auto-adjustment logic for stock splits and bonuses.
"""

import unittest
from unittest.mock import patch, MagicMock, mock_open
import pandas as pd
import json

from src.corporate_actions import process_splits


class TestCorporateActions(unittest.TestCase):
    """Test suite for the process_splits function."""

    def setUp(self):
        """Create a standard raw trades DataFrame for testing."""
        self.df = pd.DataFrame({
            'Trade Date': ['2023-01-01', '2023-02-01', '2025-01-01'],
            'Symbol': ['APPLE', 'APPLE', 'GOOGLE'],
            'Trade Type': ['buy', 'buy', 'buy'],
            'Quantity': [10.0, 5.0, 2.0],
            'Price': [150.0, 160.0, 2500.0]
        })

    @patch('builtins.print')
    @patch('yfinance.Ticker')
    @patch('os.path.exists')
    def test_no_splits_found(self, mock_exists, mock_ticker, mock_print):
        """If no splits are found on Yahoo, it should return False and unmodified df."""
        mock_exists.return_value = False
        
        # Mock Ticker to return empty splits
        mock_instance = MagicMock()
        mock_instance.splits = pd.Series(dtype=float)
        mock_ticker.return_value = mock_instance

        modified, res_df = process_splits(self.df, base_dir='.')
        
        self.assertFalse(modified)
        self.assertTrue(self.df.equals(res_df))

    @patch('builtins.print')
    @patch('builtins.input', return_value='n')
    @patch('yfinance.Ticker')
    @patch('os.path.exists')
    def test_user_rejects_adjustment(self, mock_exists, mock_ticker, mock_input, mock_print):
        """If splits found but user types 'n', return False and unmodified df."""
        mock_exists.return_value = False
        
        mock_instance = MagicMock()
        # Fake split on 2024-01-01 (after Apple's first buy of 2023-01-01)
        mock_instance.splits = pd.Series([4.0], index=[pd.to_datetime('2024-01-01')])
        mock_ticker.return_value = mock_instance

        modified, res_df = process_splits(self.df, base_dir='.')
        
        self.assertFalse(modified)
        self.assertTrue(self.df.equals(res_df))

    @patch('builtins.print')
    @patch('builtins.open', new_callable=mock_open)
    @patch('builtins.input', return_value='y')
    @patch('yfinance.Ticker')
    @patch('os.path.exists')
    def test_apply_split_adjustment(self, mock_exists, mock_ticker, mock_input, mock_file, mock_print):
        """If user approves, quantity should multiply by ratio, price divide by ratio."""
        mock_exists.return_value = False  # JSON does not exist yet
        
        def fake_ticker(sym):
            mock_inst = MagicMock()
            if 'APPLE' in sym:
                # 4:1 Split on 2024-01-01
                mock_inst.splits = pd.Series([4.0], index=[pd.to_datetime('2024-01-01')])
            else:
                mock_inst.splits = pd.Series(dtype=float)
            return mock_inst
            
        mock_ticker.side_effect = fake_ticker

        modified, res_df = process_splits(self.df, base_dir='.')
        
        self.assertTrue(modified)
        
        # Apple qty 10 * 4 = 40, price 150 / 4 = 37.5
        self.assertEqual(res_df.loc[0, 'Quantity'], 40.0)
        self.assertEqual(res_df.loc[0, 'Price'], 37.5)
        
        # Google should remain 2.0 at 2500
        self.assertEqual(res_df.loc[2, 'Quantity'], 2.0)
        self.assertEqual(res_df.loc[2, 'Price'], 2500.0)
        
        # Check if python tried to open the JSON to save it
        mock_file.assert_called_with('.\\applied_splits.json', 'w', encoding='utf-8')

    @patch('builtins.print')
    @patch('builtins.open', new_callable=mock_open, read_data='{"APPLE": ["2024-01-01"]}')
    @patch('yfinance.Ticker')
    @patch('os.path.exists')
    def test_already_applied_split(self, mock_exists, mock_ticker, mock_file, mock_print):
        """If the split date is found in applied_splits.json, it should not prompt again."""
        mock_exists.return_value = True
        
        mock_instance = MagicMock()
        mock_instance.splits = pd.Series([4.0], index=[pd.to_datetime('2024-01-01')])
        mock_ticker.return_value = mock_instance

        modified, res_df = process_splits(self.df, base_dir='.')
        
        # Since it's in the simulated JSON, no pending action is triggered
        self.assertFalse(modified)
        self.assertTrue(self.df.equals(res_df))

    @patch('builtins.print')
    @patch('yfinance.Ticker')
    @patch('os.path.exists')
    def test_splits_before_first_buy_ignored(self, mock_exists, mock_ticker, mock_print):
        """Splits occurring before the user's first buy date should be ignored."""
        mock_exists.return_value = False
        
        mock_instance = MagicMock()
        # Split on 2022-01-01. First buy was 2023-01-01. Should be totally ignored.
        mock_instance.splits = pd.Series([2.0], index=[pd.to_datetime('2022-01-01')])
        mock_ticker.return_value = mock_instance

        modified, res_df = process_splits(self.df, base_dir='.')
        self.assertFalse(modified)


if __name__ == '__main__':
    unittest.main()
