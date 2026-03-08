"""
Tests for src/config.py
=======================

Validates that the configuration parser correctly interprets:
    - Direct numeric values
    - Percentage-of-portfolio expressions
    - Tolerance values (+/- format)
    - Cheat threshold comparisons (< operator)
    - Missing/invalid config files
"""

import os
import tempfile
import unittest

from src.config import load_config


class TestLoadConfig(unittest.TestCase):
    """Test suite for the load_config function."""

    def _write_temp_config(self, content: str) -> str:
        """Helper to write a temporary config file and return its path."""
        fd, path = tempfile.mkstemp(suffix='.cfg')
        with os.fdopen(fd, 'w') as f:
            f.write(content)
        return path

    def test_direct_numeric_values(self):
        """Direct numeric values like TOTAL_PORTFOLIO = 5000000 should parse to float."""
        path = self._write_temp_config("TOTAL_PORTFOLIO = 5000000\n")
        try:
            config = load_config(path)
            self.assertEqual(config['TOTAL_PORTFOLIO'], 5000000.0)
        finally:
            os.unlink(path)

    def test_percentage_of_portfolio(self):
        """'2% of TOTAL_PORTFOLIO' should resolve to 2% of the total portfolio value."""
        content = "TOTAL_PORTFOLIO = 5000000\nTRANCH = 2% of TOTAL_PORTFOLIO\n"
        path = self._write_temp_config(content)
        try:
            config = load_config(path)
            self.assertEqual(config['TRANCH'], 100000.0)
        finally:
            os.unlink(path)

    def test_tolerance_parsing(self):
        """'+/-10%' should parse to 0.10."""
        content = "TRANCH_TOLERANCE=+/-10%\n"
        path = self._write_temp_config(content)
        try:
            config = load_config(path)
            self.assertEqual(config['TRANCH_TOLERANCE'], 0.10)
        finally:
            os.unlink(path)

    def test_cheat_threshold(self):
        """'<75000' should parse to 75000.0."""
        content = "CHEAT=<75000\n"
        path = self._write_temp_config(content)
        try:
            config = load_config(path)
            self.assertEqual(config['CHEAT'], 75000.0)
        finally:
            os.unlink(path)

    def test_missing_config_file(self):
        """A missing config file should return an empty dict without crashing."""
        config = load_config('/nonexistent/path/config.cfg')
        self.assertEqual(config, {})

    def test_full_config(self):
        """Full realistic config should parse all values correctly."""
        content = (
            "TOTAL_PORTFOLIO = 5000000\n"
            "TRANCH = 2% of TOTAL_PORTFOLIO\n"
            "MAX_STOCKS = 20\n"
            "TRANCH_TOLERANCE=+/-10%\n"
            "CHEAT=<75000\n"
            "MAX_POSITION_SIZE = 5% of TOTAL_PORTFOLIO\n"
        )
        path = self._write_temp_config(content)
        try:
            config = load_config(path)
            self.assertEqual(config['TOTAL_PORTFOLIO'], 5000000.0)
            self.assertEqual(config['TRANCH'], 100000.0)
            self.assertEqual(config['MAX_STOCKS'], 20.0)
            self.assertEqual(config['TRANCH_TOLERANCE'], 0.10)
            self.assertEqual(config['CHEAT'], 75000.0)
            self.assertEqual(config['MAX_POSITION_SIZE'], 250000.0)
        finally:
            os.unlink(path)


if __name__ == '__main__':
    unittest.main()
