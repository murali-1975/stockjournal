"""
Tests for src/calculations.py
==============================

Validates:
    - Trade grouping and aggregation
    - Tranche/Cheat label assignment
    - Cheat-to-Tranche accumulation
    - Portfolio PnL calculations
    - Dynamic Stop Loss logic
"""

import unittest
from unittest.mock import patch

import pandas as pd
import numpy as np

from src.calculations import process_grouped_trades, _get_stop_loss, _classify_market_cap, _get_latest_tranche_cheat, calculate_portfolios


class TestProcessGroupedTrades(unittest.TestCase):
    """Test suite for the process_grouped_trades function."""

    def _make_trades(self, trades: list) -> pd.DataFrame:
        """
        Helper to build a raw trade DataFrame from a list of tuples.
        Each tuple: (date, symbol, trade_type, quantity, price)
        """
        return pd.DataFrame(trades, columns=['Trade Date', 'Symbol', 'Trade Type', 'Quantity', 'Price'])

    def test_basic_grouping_no_config(self):
        """Without config, trades should be grouped but no Tranche column added."""
        df = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 10, 2500),
            ('2025-01-01', 'RELIANCE', 'buy', 5, 2520),
        ])
        result = process_grouped_trades(df)
        self.assertEqual(len(result), 1)  # Same date + symbol + type → grouped
        self.assertEqual(result.iloc[0]['Total_Quantity'], 15)
        self.assertNotIn('Tranches/Cheat', result.columns)

    def test_tranch_label_assignment(self):
        """A buy matching the TRANCH size (within tolerance) should be labeled 'Tranch 1'."""
        df = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 40, 2500),  # Total = 100,000
        ])
        config = {'TRANCH': 100000, 'CHEAT': 75000, 'TRANCH_TOLERANCE': 0.10}
        result = process_grouped_trades(df, config)
        self.assertIn('Tranches/Cheat', result.columns)
        self.assertEqual(result.iloc[0]['Tranches/Cheat'], 'Tranch 1')

    def test_sequential_tranch_labels(self):
        """Multiple buys matching TRANCH size should produce Tranch 1, Tranch 2, etc."""
        df = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 40, 2500),  # 100,000 → Tranch 1
            ('2025-01-02', 'RELIANCE', 'buy', 40, 2500),  # 100,000 → Tranch 2
        ])
        config = {'TRANCH': 100000, 'CHEAT': 75000, 'TRANCH_TOLERANCE': 0.10}
        result = process_grouped_trades(df, config)
        labels = result.sort_values('Trade Date')['Tranches/Cheat'].tolist()
        self.assertEqual(labels, ['Tranch 1', 'Tranch 2'])

    def test_cheat_label_assignment(self):
        """A buy below the CHEAT threshold should be labeled 'Cheat 1'."""
        df = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 10, 5000),  # 50,000 < 75,000 → Cheat
        ])
        config = {'TRANCH': 100000, 'CHEAT': 75000, 'TRANCH_TOLERANCE': 0.10}
        result = process_grouped_trades(df, config)
        self.assertEqual(result.iloc[0]['Tranches/Cheat'], 'Cheat 1')

    def test_cheat_accumulation_to_tranch(self):
        """Two cheats summing to a full Tranche should bump the Tranche counter."""
        df = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 10, 5000),  # 50,000 → Cheat 1
            ('2025-01-02', 'RELIANCE', 'buy', 10, 5000),  # 50,000 → Cheat 2 (accumulated = 100,000 = 1 Tranch)
            ('2025-01-03', 'RELIANCE', 'buy', 40, 2500),  # 100,000 → should be Tranch 2 (not Tranch 1)
        ])
        config = {'TRANCH': 100000, 'CHEAT': 75000, 'TRANCH_TOLERANCE': 0.10}
        result = process_grouped_trades(df, config)
        labels = result.sort_values('Trade Date')['Tranches/Cheat'].tolist()
        self.assertEqual(labels[0], 'Cheat 1')
        self.assertEqual(labels[1], 'Tranch 1')
        self.assertEqual(labels[2], 'Tranch 2')

    def test_sell_labeled_na(self):
        """Sell transactions should always be labeled 'N/A'."""
        df = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 40, 2500),
            ('2025-01-02', 'RELIANCE', 'sell', 20, 2600),
        ])
        config = {'TRANCH': 100000, 'CHEAT': 75000, 'TRANCH_TOLERANCE': 0.10}
        result = process_grouped_trades(df, config)
        sell_row = result[result['Trade Type'] == 'sell']
        self.assertEqual(sell_row.iloc[0]['Tranches/Cheat'], 'N/A')

    def test_average_price_calculation(self):
        """Average price should be total_value / total_quantity, rounded to 2 decimals."""
        df = self._make_trades([
            ('2025-01-01', 'TCS', 'buy', 10, 3500),
            ('2025-01-01', 'TCS', 'buy', 10, 3600),
        ])
        result = process_grouped_trades(df)
        # Total qty=20, Total value=35000+36000=71000, Avg=3550.00
        self.assertEqual(result.iloc[0]['Average_Price'], 3550.0)

    def test_cheat_reinitialization_on_full_exit(self):
        """If a stock is fully sold out, the next buy should reset Tranche/Cheat counts to 1."""
        df = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 40, 2500),  # 100,000 → Tranch 1
            ('2025-01-02', 'RELIANCE', 'buy', 40, 2500),  # 100,000 → Tranch 2
            ('2025-01-03', 'RELIANCE', 'sell', 80, 2600), # 80 qty sold → fully exited (holdings = 0)
            ('2025-01-04', 'RELIANCE', 'buy', 40, 2500),  # 100,000 → should reset to Tranch 1 (not Tranch 3)
        ])
        config = {'TRANCH': 100000, 'CHEAT': 75000, 'TRANCH_TOLERANCE': 0.10}
        result = process_grouped_trades(df, config)
        labels = result.sort_values('Trade Date')['Tranches/Cheat'].tolist()
        self.assertEqual(labels, ['Tranch 1', 'Tranch 2', 'N/A', 'Tranch 1'])

    def test_post_tranch_1_marginal_tolerance(self):
        """Buys after Tranch 1 should follow tranche progression with +10% tolerance margin."""
        # Config has TRANCH=100000, CHEAT=75000, TRANCH_TOLERANCE=0.10
        config = {'TRANCH': 100000, 'CHEAT': 75000, 'TRANCH_TOLERANCE': 0.10}

        # Case 1: Tiny buy (₹30,000) after Tranch 1 should be Tranch 2, not a Cheat
        df1 = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 40, 2500),  # 100,000 -> Tranch 1
            ('2025-01-02', 'RELIANCE', 'buy', 10, 3000),  # 30,000 -> Should be Tranch 2
        ])
        result1 = process_grouped_trades(df1, config)
        labels1 = result1.sort_values('Trade Date')['Tranches/Cheat'].tolist()
        self.assertEqual(labels1, ['Tranch 1', 'Tranch 2'])

        # Case 2: Cumulative buy value of 210,000 (exactly +10% of Tranch 2 threshold 200,000) should be Tranch 2
        df2 = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 40, 2500),  # 100,000 -> Tranch 1
            ('2025-01-02', 'RELIANCE', 'buy', 30, 2500),  # 75,000 -> Tranch 2
            ('2025-01-03', 'RELIANCE', 'buy', 14, 2500),  # 35,000 -> Cumulative 210,000 (<= 200,000 + 10%) -> Tranch 2
        ])
        result2 = process_grouped_trades(df2, config)
        labels2 = result2.sort_values('Trade Date')['Tranches/Cheat'].tolist()
        self.assertEqual(labels2, ['Tranch 1', 'Tranch 2', 'Tranch 2'])

        # Case 3: Cumulative buy value of 211,000 (exceeds +10% of Tranch 2 threshold 200,000) should be Tranch 3
        df3 = self._make_trades([
            ('2025-01-01', 'RELIANCE', 'buy', 40, 2500),  # 100,000 -> Tranch 1
            ('2025-01-02', 'RELIANCE', 'buy', 30, 2500),  # 75,000 -> Tranch 2
            ('2025-01-03', 'RELIANCE', 'buy', 14, 2572),  # 36,008 -> Cumulative 211,008 (> 200,000 + 10%) -> Tranch 3
        ])
        result3 = process_grouped_trades(df3, config)
        labels3 = result3.sort_values('Trade Date')['Tranches/Cheat'].tolist()
        self.assertEqual(labels3, ['Tranch 1', 'Tranch 2', 'Tranch 3'])


class TestStopLoss(unittest.TestCase):
    """Test suite for the _get_stop_loss function."""

    def _make_grouped_df(self, labels: list) -> pd.DataFrame:
        """Helper to build a grouped DataFrame with specific Tranch labels."""
        n = len(labels)
        return pd.DataFrame({
            'Trade Date': [f'2025-01-0{i+1}' for i in range(n)],
            'Symbol': ['STOCK'] * n,
            'Trade Type': ['buy'] * n,
            'Total_Quantity': [40] * n,
            'Average_Price': [2500.0] * n,
            'Total_Value': [100000.0] * n,
            'Tranches/Cheat': labels,
        })

    def test_sl_tranch_1(self):
        """Tranch 1 SL should be -10% from average buy price."""
        grouped_df = self._make_grouped_df(['Tranch 1'])
        row = pd.Series({'Symbol': 'STOCK', 'Average_Buy_Price': 2500.0, 'EMA21': 2400.0})
        sl = _get_stop_loss(row, grouped_df)
        self.assertEqual(sl, 2250.0)  # 2500 * 0.9

    def test_sl_tranch_2(self):
        """Tranch 2 SL should be average price of Tranch 1 buys."""
        grouped_df = self._make_grouped_df(['Tranch 1', 'Tranch 2'])
        row = pd.Series({'Symbol': 'STOCK', 'Average_Buy_Price': 2500.0, 'EMA21': 2400.0})
        sl = _get_stop_loss(row, grouped_df)
        # Tranch 1 avg = 100000/40 = 2500.0
        self.assertEqual(sl, 2500.0)

    def test_sl_tranch_3(self):
        """Tranch 3 SL should be average price of Tranch 1+2+3 buys."""
        grouped_df = self._make_grouped_df(['Tranch 1', 'Tranch 2', 'Tranch 3'])
        row = pd.Series({'Symbol': 'STOCK', 'Average_Buy_Price': 2500.0, 'EMA21': 2400.0})
        sl = _get_stop_loss(row, grouped_df)
        # All same price → avg = 300000 / 120 = 2500.0
        self.assertEqual(sl, 2500.0)

    def test_sl_above_tranch_3(self):
        """Tranch >3 SL should be EMA 21."""
        grouped_df = self._make_grouped_df(['Tranch 1', 'Tranch 2', 'Tranch 3', 'Tranch 4'])
        row = pd.Series({'Symbol': 'STOCK', 'Average_Buy_Price': 2500.0, 'EMA21': 2400.0})
        sl = _get_stop_loss(row, grouped_df)
        self.assertEqual(sl, 2400.0)

    def test_sl_no_tranch_column(self):
        """Without Tranches/Cheat column, SL defaults to -10% from avg buy."""
        grouped_df = pd.DataFrame({'Symbol': ['STOCK'], 'Trade Type': ['buy']})
        row = pd.Series({'Symbol': 'STOCK', 'Average_Buy_Price': 1000.0, 'EMA21': 950.0})
        sl = _get_stop_loss(row, grouped_df)
        self.assertEqual(sl, 900.0)

    def test_sl_tranch_2_satellite(self):
        """Satellite Tranch 2 SL should follow the risk-controlled exit formula."""
        grouped_df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-01-02'],
            'Symbol': ['STOCK', 'STOCK'],
            'Trade Type': ['buy', 'buy'],
            'Total_Quantity': [1000.0, 500.0],
            'Average_Price': [100.0, 130.0],
            'Total_Value': [100000.0, 65000.0],
            'Tranches/Cheat': ['Tranch 1', 'Tranch 2'],
        })
        
        # Scenario A: Satellite stock
        row_sat = pd.Series({'Symbol': 'STOCK', 'TF_Classification': 'Satellite', 'Average_Buy_Price': 110.0, 'EMA21': 95.0})
        sl_sat = _get_stop_loss(row_sat, grouped_df)
        # Expected: (0.90 * 100000 + 65000) / 1500 = 155000 / 1500 = 103.33
        self.assertEqual(sl_sat, 103.33)
        
        # Scenario B: Core stock
        row_core = pd.Series({'Symbol': 'STOCK', 'TF_Classification': 'Core', 'Average_Buy_Price': 110.0, 'EMA21': 95.0})
        sl_core = _get_stop_loss(row_core, grouped_df)
        # Expected: Tranch 1 average price = 100000 / 1000 = 100.0
        self.assertEqual(sl_core, 100.0)


class TestCalculatePortfolios(unittest.TestCase):
    """Test suite for the calculate_portfolios function (with mocked Yahoo API)."""

    @patch('src.calculations.fetch_market_data_from_yahoo')
    def test_basic_portfolio(self, mock_yahoo):
        """A simple buy-only scenario should produce correct portfolio values."""
        mock_yahoo.return_value = {
            'RELIANCE': {'LTP': 2600.0, 'EMA9': 2590.0, 'EMA10': 2585.0, 'EMA11': 2580.0, 'EMA21': 2550.0, 'Market_Cap': 1_800_000_000_000}
        }
        df = pd.DataFrame({
            'Trade Date': ['2025-01-01'],
            'Symbol': ['RELIANCE'],
            'Trade Type': ['buy'],
            'Quantity': [10],
            'Price': [2500.0],
        })
        grouped_df = pd.DataFrame({
            'Trade Date': ['2025-01-01'],
            'Symbol': ['RELIANCE'],
            'Trade Type': ['buy'],
            'Total_Quantity': [10],
            'Average_Price': [2500.0],
            'Total_Value': [25000.0],
            'Tranches/Cheat': ['Tranch 1'],
        })

        portfolio_df, overall_df = calculate_portfolios(df, grouped_df)

        # Current Portfolio checks
        self.assertEqual(len(portfolio_df), 1)
        self.assertEqual(portfolio_df.iloc[0]['Current_Quantity'], 10)
        self.assertEqual(portfolio_df.iloc[0]['Average_Buy_Price'], 2500.0)
        self.assertEqual(portfolio_df.iloc[0]['LTP'], 2600.0)
        self.assertIn('SL', portfolio_df.columns)
        self.assertIn('EMA21', portfolio_df.columns)
        self.assertIn('Cap', portfolio_df.columns)
        self.assertEqual(portfolio_df.iloc[0]['Cap'], 'Large Cap')
        self.assertIn('Latest_Tranche', portfolio_df.columns)
        self.assertEqual(portfolio_df.iloc[0]['Latest_Tranche'], 'Tranch 1')
        self.assertIn('Holding_Period', portfolio_df.columns)
        self.assertIsInstance(portfolio_df.iloc[0]['Holding_Period'], (int, np.integer))

        # Overall Portfolio checks
        self.assertNotIn('EMA9', overall_df.columns)  # EMAs removed from overall
        self.assertIn('Cap', overall_df.columns)
        self.assertIn('Unrealized_PnL', overall_df.columns)
        self.assertEqual(overall_df.iloc[0]['Unrealized_PnL'], 1000.0)  # (2600-2500)*10
        self.assertIn('Holding_Period', overall_df.columns)

    @patch('src.calculations.fetch_market_data_from_yahoo')
    def test_realized_pnl(self, mock_yahoo):
        """Selling shares should produce correct Realized PnL."""
        mock_yahoo.return_value = {
            'TCS': {'LTP': 3500.0, 'EMA9': 0, 'EMA10': 0, 'EMA11': 0, 'EMA21': 0, 'Market_Cap': 500_000_000_000}
        }
        df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-01-05'],
            'Symbol': ['TCS', 'TCS'],
            'Trade Type': ['buy', 'sell'],
            'Quantity': [10, 5],
            'Price': [3000.0, 3200.0],
        })
        grouped_df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-01-05'],
            'Symbol': ['TCS', 'TCS'],
            'Trade Type': ['buy', 'sell'],
            'Total_Quantity': [10, 5],
            'Average_Price': [3000.0, 3200.0],
            'Total_Value': [30000.0, 16000.0],
            'Tranches/Cheat': ['Tranch 1', 'N/A'],
        })

        _, overall_df = calculate_portfolios(df, grouped_df)

        # Realized PnL = Sell Value - (Sell Qty * Avg Buy Price) = 16000 - (5*3000) = 1000
        self.assertEqual(overall_df.iloc[0]['Realized_PnL'], 1000.0)
        # Current Qty should be 10 - 5 = 5
        self.assertEqual(overall_df.iloc[0]['Current_Quantity'], 5)

    @patch('src.calculations.fetch_market_data_from_yahoo')
    def test_fully_sold_position(self, mock_yahoo):
        """A fully sold position should not appear in Current Portfolio but should be in Overall."""
        mock_yahoo.return_value = {
            'INFY': {'LTP': 1500.0, 'EMA9': 0, 'EMA10': 0, 'EMA11': 0, 'EMA21': 0, 'Market_Cap': 200_000_000_000}
        }
        df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-01-05'],
            'Symbol': ['INFY', 'INFY'],
            'Trade Type': ['buy', 'sell'],
            'Quantity': [10, 10],
            'Price': [1400.0, 1500.0],
        })
        grouped_df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-01-05'],
            'Symbol': ['INFY', 'INFY'],
            'Trade Type': ['buy', 'sell'],
            'Total_Quantity': [10, 10],
            'Average_Price': [1400.0, 1500.0],
            'Total_Value': [14000.0, 15000.0],
            'Tranches/Cheat': ['Tranch 1', 'N/A'],
        })

        portfolio_df, overall_df = calculate_portfolios(df, grouped_df)

        # Fully sold → should NOT be in current portfolio
        self.assertEqual(len(portfolio_df), 0)
        # But SHOULD be in overall portfolio
        self.assertEqual(len(overall_df), 1)
        self.assertEqual(overall_df.iloc[0]['Current_Quantity'], 0)

    @patch('src.calculations.fetch_market_data_from_yahoo')
    def test_custom_price_updates_override(self, mock_yahoo):
        """Custom price updates should override live prices, falling back to live prices if missing."""
        mock_yahoo.return_value = {
            'RELIANCE': {'LTP': 2600.0, 'EMA9': 2590.0, 'EMA10': 2585.0, 'EMA11': 2580.0, 'EMA21': 2550.0, 'Market_Cap': 1_800_000_000_000},
            'TCS': {'LTP': 3500.0, 'EMA9': 3490.0, 'EMA10': 3485.0, 'EMA11': 3480.0, 'EMA21': 3450.0, 'Market_Cap': 500_000_000_000}
        }
        df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-01-01'],
            'Symbol': ['RELIANCE', 'TCS'],
            'Trade Type': ['buy', 'buy'],
            'Quantity': [10, 10],
            'Price': [2500.0, 3400.0],
        })
        grouped_df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-01-01'],
            'Symbol': ['RELIANCE', 'TCS'],
            'Trade Type': ['buy', 'buy'],
            'Total_Quantity': [10, 10],
            'Average_Price': [2500.0, 3400.0],
            'Total_Value': [25000.0, 34000.0],
            'Tranches/Cheat': ['Tranch 1', 'Tranch 1'],
        })

        price_updates = {'RELIANCE': 2750.50} # Overridden price. TCS is missing (should fallback to 3500.0)

        portfolio_df, overall_df = calculate_portfolios(df, grouped_df, price_updates=price_updates)

        # Check RELIANCE has overridden price
        reliance_row = portfolio_df[portfolio_df['Symbol'] == 'RELIANCE'].iloc[0]
        self.assertEqual(reliance_row['LTP'], 2750.50)

        # Check TCS fallback to Yahoo Finance LTP
        tcs_row = portfolio_df[portfolio_df['Symbol'] == 'TCS'].iloc[0]
        self.assertEqual(tcs_row['LTP'], 3500.0)

    @patch('src.calculations.fetch_market_data_from_yahoo')
    def test_chronological_holding_period_reset(self, mock_yahoo):
        """If a stock is bought, fully exited, and then bought again, the holding period should start from the fresh buy date."""
        mock_yahoo.return_value = {
            'STOCK': {'LTP': 100.0, 'EMA9': 0, 'EMA10': 0, 'EMA11': 0, 'EMA21': 0, 'Market_Cap': 10_000_000}
        }
        df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-02-01', '2025-03-01'],
            'Symbol': ['STOCK', 'STOCK', 'STOCK'],
            'Trade Type': ['buy', 'sell', 'buy'],
            'Quantity': [10, 10, 5],
            'Price': [90.0, 95.0, 92.0]
        })
        grouped_df = pd.DataFrame({
            'Trade Date': ['2025-01-01', '2025-02-01', '2025-03-01'],
            'Symbol': ['STOCK', 'STOCK', 'STOCK'],
            'Trade Type': ['buy', 'sell', 'buy'],
            'Total_Quantity': [10, 10, 5],
            'Average_Price': [90.0, 95.0, 92.0],
            'Total_Value': [900.0, 950.0, 460.0],
            'Tranches/Cheat': ['Tranch 1', 'N/A', 'Tranch 1']
        })

        portfolio_df, overall_df = calculate_portfolios(df, grouped_df)

        # 1. Verification of overall portfolio
        self.assertEqual(len(overall_df), 1)
        # Holding period should be calculated from 2025-03-01 to today (not from 2025-01-01)
        today = pd.Timestamp.now().normalize()
        expected_hp = (today - pd.to_datetime('2025-03-01')).days
        self.assertEqual(overall_df.iloc[0]['Holding_Period'], expected_hp)
        
        # 2. Verification of current portfolio (held position)
        self.assertEqual(len(portfolio_df), 1)
        self.assertEqual(portfolio_df.iloc[0]['Current_Quantity'], 5)
        # Holding period should be calculated from 2025-03-01 to today (not from 2025-01-01)
        today = pd.Timestamp.now().normalize()
        expected_hp = (today - pd.to_datetime('2025-03-01')).days
        self.assertEqual(portfolio_df.iloc[0]['Holding_Period'], expected_hp)


class TestClassifyMarketCap(unittest.TestCase):
    """Test suite for the _classify_market_cap function."""

    def setUp(self):
        self.config = {
            'SMALL_CAP': {'type': 'below', 'value': 347_000_000_000},
            'LARGE_CAP': {'type': 'above', 'value': 1_050_000_000_000},
        }

    def test_large_cap(self):
        """Market cap >= 1,05,000 Cr should be Large Cap."""
        result = _classify_market_cap(1_800_000_000_000, self.config)
        self.assertEqual(result, 'Large Cap')

    def test_mid_cap(self):
        """Market cap between small and large thresholds should be Mid Cap."""
        result = _classify_market_cap(500_000_000_000, self.config)
        self.assertEqual(result, 'Mid Cap')

    def test_small_cap(self):
        """Market cap below 34,700 Cr should be Small Cap."""
        result = _classify_market_cap(100_000_000_000, self.config)
        self.assertEqual(result, 'Small Cap')

    def test_zero_market_cap(self):
        """Zero market cap should return empty string."""
        result = _classify_market_cap(0, self.config)
        self.assertEqual(result, '')

    def test_boundary_large_cap(self):
        """Market cap exactly at large cap threshold should be Large Cap."""
        result = _classify_market_cap(1_050_000_000_000, self.config)
        self.assertEqual(result, 'Large Cap')

    def test_default_thresholds(self):
        """Without config, SEBI defaults should apply."""
        result = _classify_market_cap(1_800_000_000_000, {})
        self.assertEqual(result, 'Large Cap')


class TestLatestTrancheCheat(unittest.TestCase):
    """Test suite for _get_latest_tranche_cheat function."""

    def test_single_tranch(self):
        """Single Tranch 1 should return 'Tranch 1'."""
        gdf = pd.DataFrame({
            'Symbol': ['STOCK'], 'Trade Type': ['buy'], 'Tranches/Cheat': ['Tranch 1']
        })
        self.assertEqual(_get_latest_tranche_cheat('STOCK', gdf), 'Tranch 1')

    def test_multiple_tranches(self):
        """Should return the highest numbered Tranche."""
        gdf = pd.DataFrame({
            'Symbol': ['STOCK'] * 3,
            'Trade Type': ['buy'] * 3,
            'Tranches/Cheat': ['Tranch 1', 'Tranch 2', 'Tranch 3']
        })
        self.assertEqual(_get_latest_tranche_cheat('STOCK', gdf), 'Tranch 3')

    def test_cheats_only(self):
        """If only cheats, return the highest cheat."""
        gdf = pd.DataFrame({
            'Symbol': ['STOCK'] * 2,
            'Trade Type': ['buy'] * 2,
            'Tranches/Cheat': ['Cheat 1', 'Cheat 2']
        })
        self.assertEqual(_get_latest_tranche_cheat('STOCK', gdf), 'Cheat 2')

    def test_mixed_tranch_and_cheat(self):
        """Tranch 2 should beat Cheat 1 (higher number)."""
        gdf = pd.DataFrame({
            'Symbol': ['STOCK'] * 3,
            'Trade Type': ['buy'] * 3,
            'Tranches/Cheat': ['Tranch 1', 'Cheat 1', 'Tranch 2']
        })
        self.assertEqual(_get_latest_tranche_cheat('STOCK', gdf), 'Tranch 2')

    def test_no_tranch_column(self):
        """Without Tranches/Cheat column, should return empty string."""
        gdf = pd.DataFrame({'Symbol': ['STOCK'], 'Trade Type': ['buy']})
        self.assertEqual(_get_latest_tranche_cheat('STOCK', gdf), '')

    def test_sell_rows_ignored(self):
        """Sell rows should not contribute to the latest label."""
        gdf = pd.DataFrame({
            'Symbol': ['STOCK'] * 2,
            'Trade Type': ['buy', 'sell'],
            'Total_Quantity': [100, 50],
            'Tranches/Cheat': ['Tranch 1', 'N/A']
        })
        self.assertEqual(_get_latest_tranche_cheat('STOCK', gdf), 'Tranch 1')


class TestXIRRCalculations(unittest.TestCase):
    """Test suite for the calculate_xirr solver and Return_Pct/XIRR in portfolios."""

    def test_calculate_xirr_success(self):
        """Should correctly calculate XIRR for standard cash flows."""
        cfs = [
            (pd.to_datetime('2020-01-01'), -1000.0),
            (pd.to_datetime('2021-01-01'), 1100.0)
        ]
        from src.calculations import calculate_xirr
        rate = calculate_xirr(cfs)
        self.assertAlmostEqual(rate, 0.0997135, places=5)

    def test_calculate_xirr_empty_or_same_sign(self):
        """Should return 0.0 for empty or single-signed cash flows."""
        from src.calculations import calculate_xirr
        self.assertEqual(calculate_xirr([]), 0.0)
        self.assertEqual(calculate_xirr([(pd.to_datetime('2020-01-01'), -1000.0)]), 0.0)
        self.assertEqual(calculate_xirr([(pd.to_datetime('2020-01-01'), 1000.0)]), 0.0)


if __name__ == '__main__':
    unittest.main()
