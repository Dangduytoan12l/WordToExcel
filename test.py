import re

def split_options(text: str) -> list:
    """Splits options that are on the same line into a list."""
    return re.split(r'\s+(?=[a-dA-D]\.)', text, flags=re.IGNORECASE)

options = split_options("c. (1), (2), (6), (8).     d. (3), (4), (7), (8).")
for option in options:
    print(option.strip())
