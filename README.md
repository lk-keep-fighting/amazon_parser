# Amazon Product Parser

This repository distils the original automation project down to a single, well
scoped responsibility: turning an Amazon product page into structured data.  The
parser operates on static HTML using `requests` and `BeautifulSoup`, so it can
run anywhere without a heavyweight browser dependency.

## Features

- Extracts the ASIN, title, pricing information, currency, rating and review
  count from product pages.
- Captures bullet points and technical details from both the details list and
  specification tables.
- Falls back to embedded JSON-LD blocks to enrich incomplete pages.
- Converts the result into a serialisable Python dictionary.
- Supports batch processing by reading an Excel workbook and appending structured results.
- Provides a small CLI for ad-hoc usage.
- Offers a lightweight web UI to preview imported workbooks and formatted parse results.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

### In Python code

```python
from src.amazon_product_parser import AmazonProductParser

parser = AmazonProductParser()
product = parser.parse("https://www.amazon.com/dp/B012345678")
print(product.to_dict())
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

### Web UI preview

A small Flask application is bundled to make it easier to inspect the contents of
Excel workbooks that contain parsing results.  Launch it with:

```bash
python -m src.ui_app
```

Then open <http://127.0.0.1:5000/> in your browser and upload a workbook (for
example the sample `output.xlsx` in the project root).  The interface lists the
imported product identifiers alongside every parsed field, formats bullet lists
as readable items, and pretty-prints JSON columns automatically.

### Batch processing from Excel

The CLI can parse every Amazon URL stored in the first column of an Excel
workbook and append structured data to adjacent columns:

```bash
python -m src.amazon_product_parser --excel ./input.xlsx --output ./output.xlsx
```

By default the parser reads column A from the active sheet. Use `--sheet` to
select a different sheet and `--column` to target another column.

## Running tests

```bash
pytest
```

## License

MIT
