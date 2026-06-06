"""
Stock Journal - Utility Functions
=================================

General helpers for UI and command-line execution.
"""
import sys

def print_progress_bar(iteration: int, total: int, prefix: str = '', suffix: str = '', length: int = 30):
    """
    Prints a simple text-based progress bar to stdout.
    
    Args:
        iteration: Current progress step (1-indexed or 0-indexed).
        total:     Total number of steps.
        prefix:    Message before the progress bar.
        suffix:    Message after the progress bar.
        length:    Character length of the bar.
    """
    if total <= 0:
        return
    percent = f"{100 * (iteration / float(total)):.1f}%"
    filled_length = int(length * iteration // total)
    bar = '=' * filled_length + '-' * (length - filled_length)
    import sys
    try:
        sys.stdout.write(f'\r{prefix} |{bar}| {percent} {suffix}')
        sys.stdout.flush()
    except UnicodeEncodeError:
        # Fallback in case suffix contains unencodable characters
        clean_suffix = suffix.encode('ascii', errors='ignore').decode('ascii')
        sys.stdout.write(f'\r{prefix} |{bar}| {percent} {clean_suffix}')
        sys.stdout.flush()
    if iteration >= total:
        sys.stdout.write('\n')
        sys.stdout.flush()
