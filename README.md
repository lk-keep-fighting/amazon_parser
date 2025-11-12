# Amazon Product Parser

This repository distils the original automation project down to a single, well
scoped responsibility: turning an Amazon product page into structured data.  The
default parser operates on static HTML using `requests` and `BeautifulSoup`, and
an optional Playwright-backed renderer is available for pages that require
interactive expansion of hidden details.

## Features

- Extracts the ASIN, title, pricing information, currency, rating and review
  count from product pages.
- Captures bullet points and technical details from both the details list and
  specification tables.
- Falls back to embedded JSON-LD blocks to enrich incomplete pages.
- Converts the result into a serialisable Python dictionary.
- Optional Playwright engine expands dynamic sections before parsing when required.
- Supports batch processing by reading an Excel workbook and appending structured results.
- Provides a small CLI for ad-hoc usage.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

If you plan to use the Playwright engine, install the browser binaries as well:

```bash
playwright install
```

## Usage

### In Python code

```python
from src.amazon_product_parser import AmazonProductParser, create_parser

# Static HTML parser (default)
parser = AmazonProductParser()
product = parser.parse("https://www.amazon.com/dp/B012345678")
print(product.to_dict())

# Playwright engine for pages that hide details behind expandable sections
dynamic_parser = create_parser("playwright")
dynamic_product = dynamic_parser.parse("https://www.amazon.com/dp/B012345678")
print(dynamic_product.to_dict())
```

If you already have the HTML, skip the network request:

```python
html = "... fetched elsewhere ..."
product = parser.parse_html(html, url="https://www.amazon.com/dp/B012345678")
```

### Command line interface

```bash
python -m src.amazon_product_parser "https://www.amazon.com/dp/B003G2ZKRC" --pretty
```

Add `--engine playwright` when you need the page to be rendered in a headless
browser before parsing:

```bash
python -m src.amazon_product_parser --engine playwright "https://www.amazon.com/dp/B003G2ZKRC" --pretty
```

### Graphical interface

```bash
python -m src.gui
```

This opens a Tkinter-based application that mirrors the CLI features. Use the
"Single Product" tab to parse a single URL and the "Excel Batch" tab to process
an entire workbook with the same options available on the command line. The
"Parser engine" toggle lets you switch between the static parser and the
Playwright-backed renderer.

### Batch processing from Excel

The CLI can parse every Amazon URL stored in the first column of an Excel
workbook and append structured data to adjacent columns:

```bash
python -m src.amazon_product_parser --excel ./input.xlsx --output ./output.xlsx
```

By default the parser reads column A from the active sheet. Use `--sheet` to
select a different sheet and `--column` to target another column. The `--engine`
flag can be combined with `--excel` to run the batch processor with the
Playwright renderer when needed.

## Running tests

```bash
pytest
```

## License

MIT
