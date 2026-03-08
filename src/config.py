"""
Configuration Parser
====================

Reads and parses the `input.cfg` configuration file used to define portfolio
sizing rules, Tranche thresholds, Cheat limits, and tolerance values.

The config file uses a simple `KEY = VALUE` format with support for:
    - Direct numeric values (e.g., TOTAL_PORTFOLIO = 5000000)
    - Percentage-of-portfolio expressions (e.g., TRANCH = 2% of TOTAL_PORTFOLIO)
    - Tolerance percentages (e.g., TRANCH_TOLERANCE = +/-10%)
    - Comparison operators (e.g., CHEAT = <75000)

Example input.cfg:
    TOTAL_PORTFOLIO = 5000000
    TRANCH = 2% of TOTAL_PORTFOLIO
    TRANCH_TOLERANCE = +/-10%
    CHEAT = <75000
"""


def load_config(config_file: str) -> dict:
    """
    Parses a configuration file to extract key-value pairs.

    Performs a two-pass parse:
        1. First pass: extracts raw numeric values.
        2. Second pass: resolves percentage-of-portfolio expressions,
           tolerance values, and comparison operators.

    Args:
        config_file: Absolute or relative path to the `.cfg` file.

    Returns:
        A dictionary mapping configuration keys to their resolved float values.
        Returns an empty dict if the file is not found.

    Example:
        >>> config = load_config('input.cfg')
        >>> config['TRANCH']
        100000.0
    """
    config = {}
    raw_lines = {}
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            for line in f:
                if '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    raw_lines[key] = value

                    # First pass: try direct numeric conversion
                    numeric_part = value.split()[0]
                    if '%' not in numeric_part:
                        try:
                            config[key] = float(numeric_part)
                        except ValueError:
                            # Preserve non-numeric values as strings (e.g., file paths)
                            config[key] = value

            # Second pass: evaluate percentages of TOTAL_PORTFOLIO and specific string rules
            total_portfolio = config.get('TOTAL_PORTFOLIO', 0)
            for key, value in raw_lines.items():
                if isinstance(value, str):
                    if '%' in value and 'TOTAL_PORTFOLIO' in value:
                        try:
                            percent_str = value.split('%')[0].strip()
                            percent_val = float(percent_str) / 100
                            config[key] = percent_val * total_portfolio
                        except ValueError:
                            pass
                    elif key == 'TRANCH_TOLERANCE':
                        if '%' in value:
                            val_str = value.replace('+/-', '').replace('%', '').strip()
                            try:
                                config[key] = float(val_str) / 100
                            except ValueError:
                                pass
                    elif key == 'CHEAT':
                        if '<' in value:
                            val_str = value.replace('<', '').strip()
                            try:
                                config[key] = float(val_str)
                            except ValueError:
                                pass
                        elif '%' not in value:
                            try:
                                config[key] = float(value)
                            except ValueError:
                                pass
                    elif key in ('SMALL_CAP', 'MEDIUM_CAP', 'LARGE_CAP'):
                        # Parse market cap thresholds in ₹ Crore format
                        # Examples: "Below ₹34,700 Cr", "Above ₹1,05,000 Cr",
                        #           "Between ₹34,700 Cr and ₹1,05,000 Cr"
                        config[key] = _parse_cap_threshold(value)
    except FileNotFoundError:
        print(f"Config file {config_file} not found. Skipping tranche logic.")

    return config


def _parse_cap_threshold(value: str) -> dict:
    """
    Parses a market cap threshold string into a structured dict.

    Supported formats:
        - "Below ₹34,700 Cr"       → {'type': 'below', 'value': 347000000000}
        - "Above ₹1,05,000 Cr"     → {'type': 'above', 'value': 1050000000000}
        - "Between ₹34,700 Cr and ₹1,05,000 Cr" → {'type': 'between', 'low': ..., 'high': ...}

    Values are converted from Crores to absolute numbers (1 Cr = 10,000,000).

    Args:
        value: The raw config string for the cap threshold.

    Returns:
        A dict describing the threshold type and value(s).
    """
    import re

    ONE_CRORE = 10_000_000

    def _extract_crore_number(s: str) -> float:
        """Extracts a number from a string like '₹34,700' or '₹1,05,000'."""
        # Remove ₹ symbol and commas, then convert to float
        cleaned = s.replace('₹', '').replace(',', '').strip()
        return float(cleaned) * ONE_CRORE

    val_lower = value.lower().strip()

    if val_lower.startswith('below'):
        numbers = re.findall(r'₹[\d,]+', value)
        if numbers:
            return {'type': 'below', 'value': _extract_crore_number(numbers[0])}
    elif val_lower.startswith('above'):
        numbers = re.findall(r'₹[\d,]+', value)
        if numbers:
            return {'type': 'above', 'value': _extract_crore_number(numbers[0])}
    elif val_lower.startswith('between'):
        numbers = re.findall(r'₹[\d,]+', value)
        if len(numbers) >= 2:
            return {
                'type': 'between',
                'low': _extract_crore_number(numbers[0]),
                'high': _extract_crore_number(numbers[1])
            }

    return {'type': 'unknown', 'value': 0}

