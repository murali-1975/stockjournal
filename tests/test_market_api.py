"""
Tests for src/market_api.py
===========================

Validates the retrieval of Market Data (LTP, EMAs, Market Cap, Splits)
from Yahoo Finance. Mocks network responses to ensure robust, offline testing.
"""

import unittest
from unittest.mock import patch, MagicMock
import pandas as pd

from src.market_api import fetch_market_data_from_yahoo


class TestMarketAPI(unittest.TestCase):
    """Test suite for the fetch_market_data_from_yahoo function."""

    @patch('yfinance.Ticker')
    @patch('yfinance.download')
    def test_fetch_single_symbol(self, mock_download, mock_ticker):
        """Test fetching data for a single symbol."""
        # Mock yf.download returning a DataFrame with a 'Close' column
        mock_series = pd.Series([100.0, 110.0, 120.0])
        mock_df = pd.DataFrame({'Close': mock_series})
        mock_download.return_value = mock_df

        # Mock yf.Ticker info and splits
        mock_ticker_instance = MagicMock()
        mock_ticker_instance.info = {'marketCap': 500000000000}
        mock_ticker_instance.splits = pd.Series([2.0], index=[pd.to_datetime('2025-01-01')])
        mock_ticker.return_value = mock_ticker_instance

        symbols = ['RELIANCE']
        result = fetch_market_data_from_yahoo(symbols)

        # Assert correct structure and values
        self.assertIn('RELIANCE', result)
        data = result['RELIANCE']
        self.assertEqual(data['LTP'], 120.0)
        self.assertIn('EMA9', data)
        self.assertIn('EMA21', data)
        self.assertEqual(data['Market_Cap'], 500000000000)
        self.assertIn('Splits', data)
        self.assertEqual(len(data['Splits']), 1)
        self.assertEqual(data['Splits'].iloc[0], 2.0)

    @patch('yfinance.Ticker')
    @patch('yfinance.download')
    def test_fetch_multiple_symbols(self, mock_download, mock_ticker):
        """Test fetching data for multiple symbols where yfinance returns MultiIndex or columns."""
        # Mock yf.download returning a DataFrame with Close columns for each symbol
        mock_df = pd.DataFrame({
            ('Close', 'TCS.NS'): [3000.0, 3100.0, 3200.0],
            ('Close', 'INFY.NS'): [1400.0, 1450.0, 1500.0]
        })
        mock_download_df = pd.DataFrame(
            [[3000.0, 1400.0], [3100.0, 1450.0], [3200.0, 1500.0]],
            columns=['TCS.NS', 'INFY.NS']
        )
        
        # When accessing data['Close'], it should return the above dataframe
        mock_data = MagicMock()
        mock_data.__contains__.return_value = True
        mock_data.__getitem__.return_value = mock_download_df
        mock_download.return_value = mock_data

        # Mock yf.Ticker
        mock_ticker_instance = MagicMock()
        mock_ticker_instance.info = {'marketCap': 0}
        mock_ticker_instance.splits = None
        mock_ticker.return_value = mock_ticker_instance

        symbols = ['TCS', 'INFY']
        result = fetch_market_data_from_yahoo(symbols)

        self.assertIn('TCS', result)
        self.assertIn('INFY', result)
        self.assertEqual(result['TCS']['LTP'], 3200.0)
        self.assertEqual(result['INFY']['LTP'], 1500.0)

    @patch('yfinance.download')
    def test_fetch_empty_list(self, mock_download):
        """Fetching data for an empty list should return an empty dict without calling API."""
        result = fetch_market_data_from_yahoo([])
        self.assertEqual(result, {})
        mock_download.assert_not_called()

    @patch('yfinance.download')
    def test_fetch_api_error(self, mock_download):
        """If API raises an Exception, it should catch gracefully and return defaults."""
        mock_download.side_effect = Exception("API Timeout")
        result = fetch_market_data_from_yahoo(['STOCK'])
        
        self.assertIn('STOCK', result)
        self.assertEqual(result['STOCK']['LTP'], 0.0)

    @patch('src.market_api.datetime')
    @patch('yfinance.Ticker')
    @patch('yfinance.download')
    def test_fetch_calendar_relative_closes(self, mock_download, mock_ticker, mock_datetime):
        """Test that Prev_Week_Close and Prev_Month_Close are calculated calendar-relatively."""
        import datetime
        mock_now = MagicMock()
        # Mock today as Saturday, May 16, 2026
        mock_now.date.return_value = datetime.date(2026, 5, 16)
        mock_datetime.now.return_value = mock_now

        # Set up a series of daily prices: April 30 to May 15
        date_range = pd.date_range(start='2026-05-01', end='2026-05-15', freq='D')
        index_dates = [pd.Timestamp('2026-04-30')] + list(date_range)
        prices = [100.0] + [100.0 + idx + 1 for idx in range(len(date_range))]
        mock_series = pd.Series(prices, index=index_dates)
        mock_df = pd.DataFrame({'Close': mock_series})
        mock_download.return_value = mock_df

        # Mock Ticker
        mock_ticker_instance = MagicMock()
        mock_ticker_instance.info = {'marketCap': 0}
        mock_ticker_instance.splits = None
        mock_ticker.return_value = mock_ticker_instance

        # Call function
        result = fetch_market_data_from_yahoo(['STOCK'], classifications={'STOCK': 'Satellite'})

        # Verify results:
        # Today is May 16 (Sat). Target Friday = May 15. Previous week close = May 15 close.
        # Previous Month Close = April 30 close (100.0).
        self.assertIn('STOCK', result)
        data = result['STOCK']
        
        # April 30 price is 100.0
        # May 15 price is 100.0 + 15 = 115.0
        self.assertEqual(data['Prev_Week_Close'], 115.0)
        self.assertEqual(data['Prev_Month_Close'], 100.0)


if __name__ == '__main__':
    unittest.main()

