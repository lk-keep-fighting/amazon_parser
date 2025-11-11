"""Amazon product parser.

This module provides a lightweight parser that converts Amazon product pages
into a structured :class:`AmazonProduct` data object.  It focuses on static HTML
parsing via :mod:`requests` and :mod:`bs4`, keeping the implementation simple
and dependency free from browser automation frameworks.

The parser exposes two main entry points:

``AmazonProductParser.parse(url)``
    Fetches the target URL and returns structured product information.

``AmazonProductParser.parse_html(html, url=None)``
    Parses an in-memory HTML string.  Useful for testing or when the caller
    already has the page content.

The resulting :class:`AmazonProduct` instance can be easily serialised with the
``to_dict`` helper.  A small CLI is available via ``python -m
src.amazon_product_parser <amazon-url>``.
"""

from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Iterable, List, Optional, Set, Tuple
from urllib.parse import parse_qs, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# A conservative desktop user-agent keeps Amazon happy while remaining predictable
DEFAULT_HEADERS: Dict[str, str] = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/118.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# Regex helpers used across different extraction routines
PRICE_REGEX = re.compile(
    r"(?P<symbol>US\$|CA\$|AU\$|A\$|C\$|£|€|¥|₹|\$)\s*(?P<amount>\d[\d,\.]*)"
)
AMOUNT_REGEX = re.compile(r"\d[\d,\.]*")
FLOAT_REGEX = re.compile(r"([0-9]+(?:\.[0-9]+)?)")

CURRENCY_MAP: Dict[str, str] = {
    "US$": "USD",
    "CA$": "CAD",
    "C$": "CAD",
    "AU$": "AUD",
    "A$": "AUD",
    "£": "GBP",
    "€": "EUR",
    "¥": "JPY",
    "₹": "INR",
    "$": "USD",
}
CURRENCY_WORD_MAP: Dict[str, str] = {
    "usd": "USD",
    "cad": "CAD",
    "aud": "AUD",
    "eur": "EUR",
    "gbp": "GBP",
    "inr": "INR",
    "jpy": "JPY",
}


@dataclass
class AmazonProduct:
    """Container for structured product information."""

    url: str
    asin: Optional[str] = None
    title: Optional[str] = None
    price: Optional[float] = None
    currency: Optional[str] = None
    rating: Optional[float] = None
    review_count: Optional[int] = None
    bullet_points: List[str] = field(default_factory=list)
    details: Dict[str, str] = field(default_factory=dict)
    image: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Return a JSON-serialisable representation of the product."""

        return {
            "url": self.url,
            "asin": self.asin,
            "title": self.title,
            "price": self.price,
            "currency": self.currency,
            "rating": self.rating,
            "review_count": self.review_count,
            "bullet_points": list(self.bullet_points),
            "details": dict(self.details),
            "image": self.image,
        }


class AmazonProductParser:
    """Parse Amazon product pages into structured information."""

    def __init__(
        self,
        *,
        session: Optional[requests.Session] = None,
        timeout: float = 15.0,
    ) -> None:
        self.session = session or requests.Session()
        self.timeout = timeout
        self.headers = DEFAULT_HEADERS.copy()

        # Populate the underlying requests session with deterministic headers so
        # callers benefit from the same behaviour whether or not they inject a custom session.
        if hasattr(self.session, "headers") and isinstance(self.session.headers, dict):
            for key, value in DEFAULT_HEADERS.items():
                self.session.headers.setdefault(key, value)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def parse(self, url: str) -> AmazonProduct:
        """Fetch ``url`` and parse the resulting Amazon product page."""

        response = self.session.get(url, headers=self.headers, timeout=self.timeout)
        response.raise_for_status()
        return self.parse_html(response.text, url=url)

    def parse_html(self, html: str, url: Optional[str] = None) -> AmazonProduct:
        """Parse pre-fetched HTML into an :class:`AmazonProduct`."""

        soup = BeautifulSoup(html, "html.parser")

        product = AmazonProduct(url=url or "")
        product.details = self._extract_details(soup)
        product.asin = self._extract_asin(url, soup, product.details)
        product.title = self._extract_title(soup)
        product.price, product.currency = self._extract_price(soup)
        product.rating = self._extract_rating(soup)
        product.review_count = self._extract_review_count(soup)
        product.bullet_points = self._extract_bullets(soup)
        product.image = self._extract_main_image(soup)

        # Enrich with structured data blocks when available.
        self._enrich_from_structured_data(product, soup)

        # Ensure ASIN consistency with the structured data.
        if not product.asin:
            product.asin = self._asin_from_details(product.details)

        if product.asin:
            detail_keys_lower = {key.lower() for key in product.details}
            if "asin" not in detail_keys_lower:
                product.details["ASIN"] = product.asin

        return product

    # ------------------------------------------------------------------
    # Extraction helpers
    # ------------------------------------------------------------------
    def _extract_title(self, soup: BeautifulSoup) -> Optional[str]:
        element = soup.select_one("#productTitle")
        return _clean_text(element.get_text()) if element else None

    def _extract_price(self, soup: BeautifulSoup) -> Tuple[Optional[float], Optional[str]]:
        selectors = [
            "#corePriceDisplay_desktop_feature_div span.a-offscreen",
            "#tp-price-block_total_price_ww span.a-offscreen",
            "#newBuyBoxPrice",
            "#priceblock_ourprice",
            "#priceblock_dealprice",
            "#priceblock_saleprice",
            "span.a-price span.a-offscreen",
            "span[data-a-color='price']",
        ]
        for selector in selectors:
            element = soup.select_one(selector)
            if element:
                price, currency = _parse_price_text(element.get_text())
                if price is not None or currency is not None:
                    return price, currency
        return None, None

    def _extract_rating(self, soup: BeautifulSoup) -> Optional[float]:
        candidates = [
            "span[data-hook='rating-out-of-text']",
            "#acrPopover",
            "#acrPopover span.a-icon-alt",
            "span.reviewCountTextLinkedHistogram[data-hook='rating-out-of-text']",
        ]
        for selector in candidates:
            element = soup.select_one(selector)
            if not element:
                continue
            source = element.get("title") or element.get_text()
            value = _extract_float(source)
            if value is not None:
                return value
        return None

    def _extract_review_count(self, soup: BeautifulSoup) -> Optional[int]:
        candidates = [
            "#acrCustomerReviewText",
            "#acrCustomerReviewTextWithoutGlobalLink",
            "span[data-hook='total-review-count']",
            "span[data-hook='rating-count']",
            "span[data-automation-id='review-count']",
        ]
        for selector in candidates:
            element = soup.select_one(selector)
            if not element:
                continue
            value = _extract_int(element.get_text())
            if value is not None:
                return value
        return None

    def _extract_bullets(self, soup: BeautifulSoup) -> List[str]:
        bullets: List[str] = []
        for span in soup.select("#feature-bullets li span"):
            text = _clean_text(span.get_text())
            if text:
                bullets.append(text)
        return bullets

    def _extract_details(self, soup: BeautifulSoup) -> Dict[str, str]:
        details: Dict[str, str] = {}
        seen_keys: Set[str] = set()

        def add_detail(key: str, value: str) -> None:
            key_clean = _clean_text(key).rstrip(":")
            value_clean = _clean_text(value)
            if not key_clean or not value_clean:
                return
            key_lower = key_clean.lower()
            if key_lower in seen_keys:
                return
            seen_keys.add(key_lower)
            details[key_clean] = value_clean

        # Detail bullets block (common on desktop layout)
        for item in soup.select("#detailBullets_feature_div li"):
            text = _clean_text(item.get_text(" ", strip=True))
            if not text or ":" not in text:
                continue
            key, value = text.split(":", 1)
            add_detail(key, value)

        # Technical detail tables
        table_selectors = [
            "table#productDetails_techSpec_section_1",
            "table#productDetails_detailBullets_sections1",
            "table.prodDetTable",
            "table.a-normal.a-spacing-micro",
        ]
        for selector in table_selectors:
            for row in soup.select(f"{selector} tr"):
                header = row.select_one("th, td.a-color-secondary")
                value = row.select_one("td")
                if not header or not value:
                    continue
                add_detail(header.get_text(), value.get_text(" ", strip=True))

        return details

    def _extract_main_image(self, soup: BeautifulSoup) -> Optional[str]:
        image = soup.select_one("#landingImage")
        if image:
            source = image.get("data-old-hires")
            if not source:
                dynamic = image.get("data-a-dynamic-image")
                if dynamic:
                    try:
                        dynamic_data = json.loads(dynamic)
                    except json.JSONDecodeError:
                        dynamic_data = None
                    if isinstance(dynamic_data, dict) and dynamic_data:
                        source = next(iter(dynamic_data.keys()), None)
            if not source:
                source = image.get("src")
            if source:
                return source.strip()
        meta = soup.select_one("meta[property='og:image']")
        if meta and meta.get("content"):
            return meta["content"].strip()
        return None

    def _extract_asin(
        self,
        url: Optional[str],
        soup: BeautifulSoup,
        details: Dict[str, str],
    ) -> Optional[str]:
        asin = self._asin_from_url(url) if url else None
        if asin:
            return asin

        asin_input = soup.select_one("input#ASIN")
        if asin_input and asin_input.get("value"):
            candidate = asin_input["value"].strip()
            if _looks_like_asin(candidate):
                return candidate.upper()

        data_element = soup.select_one("[data-asin]")
        if data_element:
            candidate = data_element.get("data-asin")
            if candidate and _looks_like_asin(candidate):
                return candidate.upper()

        return self._asin_from_details(details)

    # ------------------------------------------------------------------
    # Structured data helpers
    # ------------------------------------------------------------------
    def _enrich_from_structured_data(
        self,
        product: AmazonProduct,
        soup: BeautifulSoup,
    ) -> None:
        for script in soup.select("script[type='application/ld+json']"):
            text = script.string
            if not text:
                continue
            try:
                payload = json.loads(text)
            except json.JSONDecodeError:
                # Some Amazon pages embed multiple JSON objects without wrapping
                # them in a list.  Fall back to a simple heuristic split.
                fragments = _split_json_documents(text)
                for fragment in fragments:
                    try:
                        payload = json.loads(fragment)
                    except json.JSONDecodeError:
                        continue
                    self._apply_structured_data(product, payload)
                continue
            self._apply_structured_data(product, payload)

    def _apply_structured_data(self, product: AmazonProduct, payload: Any) -> None:
        if isinstance(payload, list):
            for entry in payload:
                self._apply_structured_data(product, entry)
            return

        if not isinstance(payload, dict):
            return

        raw_type = payload.get("@type")
        if isinstance(raw_type, list):
            types = [str(entry) for entry in raw_type]
        elif raw_type is None:
            types = []
        else:
            types = [str(raw_type)]
        if not any(t.lower() == "product" for t in types):
            return

        product.title = product.title or payload.get("name")
        sku = payload.get("asin") or payload.get("sku")
        if sku and _looks_like_asin(str(sku)):
            product.asin = (product.asin or str(sku)).upper()

        image = payload.get("image")
        if isinstance(image, str):
            product.image = product.image or image
        elif isinstance(image, list) and image:
            product.image = product.image or str(image[0])

        offers = payload.get("offers")
        if isinstance(offers, dict):
            price_text = offers.get("price")
            price_currency = offers.get("priceCurrency")
            if price_text and product.price is None:
                try:
                    product.price = float(str(price_text))
                except ValueError:
                    pass
            if price_currency and product.currency is None:
                product.currency = str(price_currency).upper()

        aggregate = payload.get("aggregateRating")
        if isinstance(aggregate, dict):
            if product.rating is None:
                rating_value = aggregate.get("ratingValue")
                if rating_value is not None:
                    product.rating = _extract_float(str(rating_value))
            if product.review_count is None:
                review_count = aggregate.get("reviewCount") or aggregate.get("ratingCount")
                if review_count is not None:
                    count = _extract_int(str(review_count))
                    if count is not None:
                        product.review_count = count

        brand = payload.get("brand")
        if isinstance(brand, dict):
            brand_name = brand.get("name")
            if brand_name:
                existing_keys = {key.lower() for key in product.details}
                if "brand" not in existing_keys:
                    product.details["Brand"] = _clean_text(str(brand_name))

    # ------------------------------------------------------------------
    # ASIN helper logic
    # ------------------------------------------------------------------
    def _asin_from_url(self, url: Optional[str]) -> Optional[str]:
        if not url:
            return None
        parsed = urlparse(url)
        path_match = re.search(r"/(?:dp|gp/product|gp/aw/d)/([A-Z0-9]{10})", parsed.path, re.IGNORECASE)
        if path_match:
            return path_match.group(1).upper()

        # General fallback: first 10 character alphanumeric group
        fallback_match = re.search(r"/([A-Z0-9]{10})(?:[/?]|$)", parsed.path, re.IGNORECASE)
        if fallback_match:
            return fallback_match.group(1).upper()

        query = parse_qs(parsed.query)
        for key in ("asin", "ASIN"):
            if key in query and query[key]:
                candidate = query[key][0]
                if _looks_like_asin(candidate):
                    return candidate.upper()
        return None

    def _asin_from_details(self, details: Dict[str, str]) -> Optional[str]:
        for key, value in details.items():
            if "asin" in key.lower() and _looks_like_asin(value):
                return value.upper()
        return None


# ----------------------------------------------------------------------
# Utility helpers
# ----------------------------------------------------------------------


def _clean_text(text: Optional[str]) -> str:
    if not text:
        return ""
    return " ".join(text.split())


def _parse_price_text(text: Optional[str]) -> Tuple[Optional[float], Optional[str]]:
    if not text:
        return None, None

    cleaned = _clean_text(text)
    match = PRICE_REGEX.search(cleaned)
    if match:
        symbol = match.group("symbol")
        amount = match.group("amount")
        currency = CURRENCY_MAP.get(symbol)
        return _coerce_float(amount), currency

    amount_match = AMOUNT_REGEX.search(cleaned)
    if not amount_match:
        return None, None

    amount_txt = amount_match.group(0)
    currency = _infer_currency(cleaned)
    return _coerce_float(amount_txt), currency


def _infer_currency(text: str) -> Optional[str]:
    lower = text.lower()
    for word, currency in CURRENCY_WORD_MAP.items():
        if word in lower:
            return currency
    return None


def _coerce_float(text: str) -> Optional[float]:
    try:
        return float(text.replace(",", ""))
    except (ValueError, AttributeError):
        return None


def _extract_float(text: Optional[str]) -> Optional[float]:
    if not text:
        return None
    match = FLOAT_REGEX.search(text)
    if not match:
        return None
    return _coerce_float(match.group(1))


def _extract_int(text: Optional[str]) -> Optional[int]:
    if not text:
        return None
    match = AMOUNT_REGEX.search(text)
    if not match:
        return None
    try:
        # Some strings include decimal parts – coerce to int after float conversion.
        return int(float(match.group(0).replace(",", "")))
    except ValueError:
        return None


def _looks_like_asin(value: str) -> bool:
    candidate = value.strip().upper()
    return bool(re.fullmatch(r"[A-Z0-9]{10}", candidate))


def _looks_like_url(value: Any) -> bool:
    if value is None:
        return False
    text = str(value).strip().lower()
    return text.startswith("http://") or text.startswith("https://")


def _split_json_documents(text: str) -> List[str]:
    """Split multiple JSON documents embedded in the same script tag."""

    fragments: List[str] = []
    buffer: List[str] = []
    depth = 0
    in_string = False
    escape = False

    for char in text:
        buffer.append(char)
        if escape:
            escape = False
            continue
        if char == "\\":
            escape = True
            continue
        if char == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                fragments.append("".join(buffer).strip())
                buffer = []
    return fragments


# ----------------------------------------------------------------------
# Excel batch processing
# ----------------------------------------------------------------------


def process_excel(
    input_path: str,
    *,
    output_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    column: int = 1,
    parser: Optional[Any] = None,
) -> str:
    """Parse Amazon URLs found in the first column of an Excel workbook.

    Args:
        input_path: Path to the source workbook.
        output_path: Optional destination path. Defaults to in-place updates.
        sheet_name: Target sheet name. Defaults to the workbook's active sheet.
        column: 1-based index of the column containing Amazon URLs.
        parser: Optional parser instance used for URL parsing (dependency injection friendly).

    Returns:
        The path to the workbook that contains the updated data.
    """

    if column < 1:
        raise ValueError("column must be greater than or equal to 1")

    parser_instance = parser or AmazonProductParser()

    source_path = str(input_path)
    workbook = load_workbook(filename=source_path)
    worksheet = workbook[sheet_name] if sheet_name else workbook.active

    result_columns: List[Tuple[str, Callable[[Optional[AmazonProduct], str], Any]]] = [
        ("ASIN", lambda product, error: product.asin if product else None),
        ("Title", lambda product, error: product.title if product else None),
        (
            "Price",
            lambda product, error: product.price if product and product.price is not None else None,
        ),
        ("Currency", lambda product, error: product.currency if product else None),
        (
            "Rating",
            lambda product, error: product.rating if product and product.rating is not None else None,
        ),
        (
            "Review Count",
            lambda product, error: product.review_count
            if product and product.review_count is not None
            else None,
        ),
        ("Image", lambda product, error: product.image if product else None),
        (
            "Bullet Points",
            lambda product, error: "\n".join(product.bullet_points)
            if product and product.bullet_points
            else None,
        ),
        (
            "Details JSON",
            lambda product, error: json.dumps(product.details, ensure_ascii=False, sort_keys=True)
            if product and product.details
            else None,
        ),
        ("Error", lambda product, error: error or None),
    ]

    for offset, (header, _) in enumerate(result_columns, start=1):
        worksheet.cell(row=1, column=column + offset, value=header)

    first_cell_value = worksheet.cell(row=1, column=column).value
    first_cell_text = str(first_cell_value).strip() if first_cell_value is not None else ""
    has_header = bool(first_cell_text) and not (
        _looks_like_url(first_cell_text) or _looks_like_asin(first_cell_text)
    )
    data_start_row = 2 if has_header else 1
    max_row = worksheet.max_row or 0

    for row_index in range(data_start_row, max_row + 1):
        cell_value = worksheet.cell(row=row_index, column=column).value
        if cell_value is None:
            continue

        cell_text = str(cell_value).strip()
        if not cell_text:
            continue

        if _looks_like_url(cell_text):
            url_text = cell_text
        elif _looks_like_asin(cell_text):
            url_text = f"https://www.amazon.com/dp/{cell_text.upper()}"
        else:
            continue

        product: Optional[AmazonProduct]
        error_message = ""
        try:
            product = parser_instance.parse(url_text)
        except Exception as exc:  # pragma: no cover - network/parse failures
            product = None
            error_message = str(exc)

        for offset, (_, accessor) in enumerate(result_columns, start=1):
            value = accessor(product, error_message)
            worksheet.cell(row=row_index, column=column + offset, value=value)

    destination = str(output_path) if output_path else source_path
    workbook.save(destination)
    workbook.close()
    return destination


# ----------------------------------------------------------------------
# Command line interface
# ----------------------------------------------------------------------


def cli(argv: Optional[Iterable[str]] = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Parse Amazon product pages into structured data")
    parser.add_argument(
        "target",
        nargs="?",
        help="Amazon product URL or, when used with --excel, the path to an Excel workbook",
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print the resulting JSON output for single URL parsing",
    )
    parser.add_argument(
        "--excel",
        action="store_true",
        help="Interpret the target argument as an Excel workbook to parse in batch",
    )
    parser.add_argument(
        "--sheet",
        help="Sheet name to process when reading from Excel (defaults to the active sheet)",
    )
    parser.add_argument(
        "--output",
        help="Destination workbook path when processing Excel files (defaults to in-place updates)",
    )
    parser.add_argument(
        "--column",
        type=int,
        default=1,
        help="1-based column index that contains Amazon URLs when processing Excel files",
    )
    args = parser.parse_args(args=list(argv) if argv is not None else None)

    product_parser = AmazonProductParser()

    if args.excel:
        if not args.target:
            parser.error("Excel mode requires a path to the workbook")
        destination = process_excel(
            args.target,
            output_path=args.output,
            sheet_name=args.sheet,
            column=args.column,
            parser=product_parser,
        )
        print(f"Excel updated: {destination}")
        return 0

    if not args.target:
        parser.error("Please provide an Amazon product URL or use --excel for batch mode")

    try:
        product = product_parser.parse(args.target)
    except requests.RequestException as exc:  # pragma: no cover - network failure path
        print(f"Failed to download product page: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - unexpected failure path
        print(f"Failed to parse product page: {exc}", file=sys.stderr)
        return 1

    payload = product.to_dict()
    json.dump(payload, sys.stdout, ensure_ascii=False, indent=2 if args.pretty else None)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI invocation
    raise SystemExit(cli())
