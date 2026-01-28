# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an Amazon product page parser that extracts structured data (ASIN, title, price, rating, bullet points, technical details, images) from Amazon URLs. The project uses Python with two parsing engines:

1. **Static parser** (`AmazonProductParser`) - Uses requests + BeautifulSoup for fast HTML parsing
2. **Playwright parser** (`PlaywrightAmazonProductParser`) - Uses headless browser to expand dynamic/collapsible sections before parsing

The Playwright parser includes a fallback mechanism to static parsing if rendering fails.

## Commands

### Installation
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
playwright install  # Required for Playwright engine
```

### Running tests
```bash
pytest
```

### Single product (CLI)
```bash
# Static parser
python -m src.amazon_product_parser "https://www.amazon.com/dp/B003G2ZKRC" --pretty

# Playwright standard mode
python -m src.amazon_product_parser --engine playwright "https://www.amazon.com/dp/B003G2ZKRC" --pretty

# Playwright fast mode (default, ~10s per product)
python -m src.amazon_product_parser --engine playwright-fast "https://www.amazon.com/dp/B003G2ZKRC" --pretty
```

### Excel batch processing
```bash
python -m src.amazon_product_parser --excel ./input.xlsx --output ./output.xlsx
```

### GUI
```bash
python -m src.gui
```

## Architecture

### Core Components

**Data Model (`AmazonProduct` dataclass)**
- Location: `src/amazon_product_parser.py:97-126`
- Container for parsed product data with `to_dict()` serialization

**Parser Classes**
- `AmazonProductParser` (lines 129-485) - Base static HTML parser using BeautifulSoup
- `PlaywrightAmazonProductParser` (lines 487-710) - Subclass that renders with headless browser first

**Factory Function** (`create_parser` at line 712)
- Accepts engine: `"playwright"`, `"playwright-fast"` (default), or `"standard"`
- `"playwright-fast"` is optimized with: headless=True, domcontentloaded wait, 8s timeout

**Batch Processor** (`process_excel` at line 877)
- Reads URLs/ASINs from Excel column
- Appends structured data to adjacent columns
- Uses openpyxl for Excel I/O

### Playwright Parser Details

The Playwright parser:
1. Opens URL in headless browser (chromium/firefox/webkit)
2. Clicks expandable sections using predefined selectors (`DEFAULT_EXPANDER_SELECTORS`)
3. Extracts HTML and delegates to parent class for parsing
4. Falls back to static parser on failure if `fallback_to_static=True` (default)

Performance settings (fast mode):
- `wait_until="domcontentloaded"` - Don't wait for all resources
- `navigation_timeout=8.0` - Shorter timeout
- `settle_timeout=0` - No extra settling time

### Code Organization

```
src/
├── __init__.py              # Package exports
├── amazon_product_parser.py # All parsing logic (~1000 lines)
└── gui.py                   # Tkinter GUI in Chinese
```

### Import Pattern

```python
from src.amazon_product_parser import AmazonProductParser, create_parser

# For static parsing
parser = AmazonProductParser()

# For Playwright (recommended for complete data)
parser = create_parser("playwright-fast")
```

### Language Notes

The codebase contains Chinese comments and logging messages (e.g., `perf_logger.info("[解析] 开始..."`). The GUI (`src/gui.py`) is entirely in Chinese.

## Testing

Tests use pytest with HTML fixtures in `tests/data/`. The test file is at `tests/test_amazon_product_parser.py`.
