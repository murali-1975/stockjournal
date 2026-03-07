"""
Stock Journal - Trade Processing Engine
========================================

A modular Python package for processing stock trade data from Excel tradebooks.
Supports incremental trade appending, Tranche/Cheat classification, market data
fetching from Yahoo Finance, Stop Loss calculations, and comprehensive PnL reporting.

Modules:
    config       - Configuration file parsing (input.cfg)
    data_io      - Excel reading/writing and data deduplication
    market_api   - Yahoo Finance LTP and EMA data fetching
    calculations - Trade grouping, portfolio math, and Stop Loss logic
    excel_writer - openpyxl formatting and workbook export
"""
