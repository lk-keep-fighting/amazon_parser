"""Amazon product parser.

This module provides a Playwright-powered parser that converts Amazon product pages
into a structured :class:`AmazonProduct` data object. It uses a headless browser
to render pages and expand dynamic content sections.

The parser exposes two main entry points:

``PlaywrightAmazonProductParser.parse(url)``
    Fetches the target URL and returns structured product information.

``PlaywrightAmazonProductParser.parse_html(html, url=None)``
    Parses an in-memory HTML string.  Useful for testing or when the caller
    already has the page content.

The resulting :class:`AmazonProduct` instance can be easily serialised with the
``to_dict`` helper.  A small CLI is available via ``python -m
src.amazon_product_parser <amazon-url>``.  Use the fast mode for better
performance while maintaining full functionality.
"""

from __future__ import annotations

import contextlib
import json
import re
import sys
import time
import logging
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Iterable, List, Optional, Set, Tuple
from urllib.parse import parse_qs, urlparse

import requests

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S.%f'
)
perf_logger = logging.getLogger('perf')

from bs4 import BeautifulSoup
from openpyxl import load_workbook

try:  # pragma: no cover - optional dependency
    from playwright.sync_api import (
        Error as PlaywrightError,
        TimeoutError as PlaywrightTimeoutError,
        sync_playwright,
    )
except ImportError:  # pragma: no cover - optional dependency
    PlaywrightError = Exception  # type: ignore
    PlaywrightTimeoutError = Exception  # type: ignore
    sync_playwright = None  # type: ignore

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
        perf_logger.info(f"{'='*60}")
        perf_logger.info(f"开始解析 HTML - URL: {url[:80] if url else 'None'}...")
        perf_logger.info(f"HTML 大小: {len(html):,} 字符")

        with PerfTimer("BeautifulSoup 解析"):
            soup = BeautifulSoup(html, "html.parser")

        product = AmazonProduct(url=url or "")

        with PerfTimer("提取 details"):
            product.details = self._extract_details(soup)
            perf_logger.info(f"  -> 提取到 {len(product.details)} 个详情字段")

        with PerfTimer("提取 ASIN"):
            product.asin = self._extract_asin(url, soup, product.details)

        with PerfTimer("提取 title"):
            product.title = self._extract_title(soup)

        with PerfTimer("提取 price"):
            product.price, product.currency = self._extract_price(soup)

        with PerfTimer("提取 rating"):
            product.rating = self._extract_rating(soup)

        with PerfTimer("提取 review_count"):
            product.review_count = self._extract_review_count(soup)

        with PerfTimer("提取 bullet_points"):
            product.bullet_points = self._extract_bullets(soup)
            perf_logger.info(f"  -> 提取到 {len(product.bullet_points)} 个要点")

        with PerfTimer("提取 image"):
            product.image = self._extract_main_image(soup)

        # Enrich with structured data blocks when available.
        with PerfTimer("从 structured data 丰富数据"):
            self._enrich_from_structured_data(product, soup)

        # Ensure ASIN consistency with the structured data.
        if not product.asin:
            with PerfTimer("从 details 提取 ASIN"):
                product.asin = self._asin_from_details(product.details)

        if product.asin:
            detail_keys_lower = {key.lower() for key in product.details}
            if "asin" not in detail_keys_lower:
                product.details["ASIN"] = product.asin

        perf_logger.info(f"解析完成 - ASIN: {product.asin}, Title: {product.title[:50] if product.title else 'None'}...")
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


class PlaywrightAmazonProductParser(AmazonProductParser):
    """Render the product page with Playwright before parsing its HTML.

    This parser delegates all extraction logic to :class:`AmazonProductParser`
    after rendering the page in a headless browser. It is useful for Amazon
    layouts that hide specification tables or detail sections behind expander
    widgets.
    """

    DEFAULT_EXPANDER_SELECTORS: Tuple[str, ...] = (
        "[data-action='a-expander-toggle']",
        ".a-expander-header",
        ".a-expander-toggle",
        ".a-expander-prompt",
        ".a-expander-content-fade .a-expander-prompt",
        "button[aria-expanded='false']",
        ".a-declarative[data-action='a-expander-toggle']",
    )

    def __init__(
        self,
        *,
        browser: str = "chromium",
        headless: bool = True,  # 默认使用无头模式，更快
        wait_until: Optional[str] = "domcontentloaded",  # 只等待 DOM 加载完成，不等待所有资源
        navigation_timeout: float = 15.0,  # 减少导航超时时间
        settle_timeout: float = 100,  # 减少等待时间
        extra_click_selectors: Optional[Iterable[str]] = None,
        fallback_to_static: bool = True,
    ) -> None:
        super().__init__(session=None)
        self.browser = browser
        self.headless = headless
        self.wait_until = wait_until
        self.navigation_timeout_ms = self._coerce_timeout_ms(navigation_timeout)
        self.settle_timeout_ms = self._coerce_timeout_ms(settle_timeout) if settle_timeout else 0.0
        self.fallback_to_static = fallback_to_static
        selectors = list(self.DEFAULT_EXPANDER_SELECTORS)
        if extra_click_selectors:
            for selector in extra_click_selectors:
                if selector not in selectors:
                    selectors.append(selector)
        self.expander_selectors: Tuple[str, ...] = tuple(selectors)

    def parse(self, url: str) -> AmazonProduct:
        total_start = time.perf_counter()
        perf_logger.info(f"{'='*60}")
        perf_logger.info(f"[parse] 开始解析 URL: {url[:80]}...")
        perf_logger.info(f"  模式: headless={self.headless}, wait_until={self.wait_until}")
        perf_logger.info(f"  超时设置: navigation={self.navigation_timeout_ms}ms, settle={self.settle_timeout_ms}ms")

        try:
            html = self._render_url(url)
        except Exception as exc:
            perf_logger.error(f"[parse] 渲染失败: {exc}")
            if not self.fallback_to_static:
                raise
            perf_logger.info("[parse] 尝试使用静态解析回退...")
            try:
                return super().parse(url)
            except Exception:
                raise exc
        return self.parse_html(html, url=url)

    @staticmethod
    def _coerce_timeout_ms(value: float) -> float:
        if value <= 0:
            raise ValueError("Timeout values must be greater than zero.")
        # Treat small values as seconds for convenience.
        return value * 1000.0 if value <= 120 else value

    def _render_url(self, url: str) -> str:
        if not url:
            raise ValueError("url must be a non-empty string")
        if sync_playwright is None:
            raise RuntimeError(
                "PlaywrightAmazonProductParser requires the 'playwright' package. "
                "Install it via 'pip install playwright' and run 'playwright install'."
            )

        browser = None
        context = None
        total_start = time.perf_counter()
        try:
            perf_logger.info("[_render_url] 启动 Playwright...")
            with sync_playwright() as playwright:
                with PerfTimer("启动浏览器"):
                    browser_factory = getattr(playwright, self.browser, None)
                    if browser_factory is None:
                        raise ValueError(
                            f"Unsupported Playwright browser '{self.browser}'. "
                            "Valid options are 'chromium', 'firefox', or 'webkit'."
                        )
                    browser = browser_factory.launch(
                        headless=self.headless,
                        args=[
                            '--no-sandbox',
                            '--disable-setuid-sandbox',
                            '--disable-dev-shm-usage',
                            '--disable-accelerated-2d-canvas',
                            '--no-first-run',
                            '--no-zygote',
                            '--disable-gpu',
                            '--disable-background-timer-throttling',
                            '--disable-backgrounding-occluded-windows',
                            '--disable-renderer-backgrounding',
                            '--disable-features=TranslateUI',
                            '--disable-ipc-flooding-protection',
                        ]
                    )

                with PerfTimer("创建浏览器上下文"):
                    context = browser.new_context(
                        user_agent=self.headers.get("User-Agent"),
                        locale="en-US",
                        extra_http_headers=self.headers,
                        # 禁用图片和一些资源加载以提高速度
                        bypass_csp=True,
                        java_script_enabled=True,
                        ignore_https_errors=True,
                    )
                    # 阻止不必要的资源加载
                    context.route("**/*.{png,jpg,jpeg,gif,svg,css,woff,woff2,ico,webp}", lambda route: route.abort())
                    context.route("**/analytics/**", lambda route: route.abort())
                    context.route("**/ads/**", lambda route: route.abort())
                    context.route("**/tracking/**", lambda route: route.abort())

                    context.set_default_navigation_timeout(self.navigation_timeout_ms)
                    context.set_default_timeout(self.navigation_timeout_ms)

                with PerfTimer("创建页面"):
                    page = context.new_page()
                    page.set_default_navigation_timeout(self.navigation_timeout_ms)
                    page.set_default_timeout(self.navigation_timeout_ms)

                goto_kwargs: Dict[str, Any] = {}
                if self.wait_until:
                    goto_kwargs["wait_until"] = self.wait_until

                with PerfTimer(f"页面导航 (wait_until={self.wait_until})"):
                    try:
                        page.goto(url, **goto_kwargs)
                    except PlaywrightTimeoutError:
                        perf_logger.warning("  -> 导航超时，使用已加载的内容继续")

                # 等待网络空闲状态，但设置较短的超时时间
                with PerfTimer("等待 networkidle (3000ms 超时)"):
                    try:
                        page.wait_for_load_state("networkidle", timeout=3000)  # 只等待 3 秒
                    except PlaywrightTimeoutError:
                        perf_logger.warning("  -> networkidle 超时（这是正常的，继续处理）")

                # 始终执行内容展开，这是重要功能
                with PerfTimer("展开动态内容"):
                    self._expand_dynamic_sections(page)

                if self.settle_timeout_ms:
                    perf_logger.info(f"[SETTLE] 等待 {self.settle_timeout_ms}ms...")
                    page.wait_for_timeout(self.settle_timeout_ms)

                with PerfTimer("获取页面内容 (page.content())"):
                    html = page.content()
                    perf_logger.info(f"  -> 获取到 HTML: {len(html):,} 字符")

                total_elapsed = (time.perf_counter() - total_start) * 1000
                perf_logger.info(f"[_render_url] 总耗时: {total_elapsed:.2f}ms")
                return html
        except PlaywrightTimeoutError as exc:
            raise RuntimeError(f"Timed out while rendering {url!r} with Playwright.") from exc
        except PlaywrightError as exc:
            raise RuntimeError(f"Playwright failed while rendering {url!r}: {exc}") from exc
        finally:
            close_start = time.perf_counter()
            if context is not None:
                with contextlib.suppress(Exception):
                    context.close()
            if browser is not None:
                with contextlib.suppress(Exception):
                    browser.close()
            close_elapsed = (time.perf_counter() - close_start) * 1000
            if close_elapsed > 100:
                perf_logger.info(f"[CLEANUP] 关闭浏览器耗时: {close_elapsed:.2f}ms")

    def _expand_dynamic_sections(self, page: Any) -> None:
        # 快速模式：先用 JavaScript 强制展开所有内容（最快）
        # 只有在需要时才点击展开器

        # 先尝试强制展开（这通常足够了）
        with PerfTimer("  [展开] 强制展开内容 (JavaScript)"):
            try:
                expanded_count = page.evaluate(
                    """
                    () => {
                        let count = 0;
                        document.querySelectorAll('.a-expander-content, .a-expander').forEach((el) => {
                            el.style.removeProperty('max-height');
                            el.style.removeProperty('height');
                            el.classList.remove('a-expander-collapsed-height');
                            if (el.hasAttribute('aria-hidden')) {
                                el.setAttribute('aria-hidden', 'false');
                            }
                            count++;
                        });
                        // 同时展开 aria-expanded=false 的元素
                        document.querySelectorAll('[aria-expanded="false"]').forEach((el) => {
                            el.setAttribute('aria-expanded', 'true');
                            count++;
                        });
                        return count;
                    }
                    """
                )
                perf_logger.info(f"  -> JavaScript 展开了 {expanded_count} 个元素")
            except PlaywrightError:
                pass

        # 滚动到页面底部（触发懒加载）
        with PerfTimer("  [展开] 滚动到页面底部"):
            try:
                page.evaluate("() => window.scrollTo(0, document.body.scrollHeight)")
                # 快速模式不需要等待，因为后面已经用 JavaScript 展开了
            except PlaywrightError:
                pass


def create_parser(engine: str = "playwright") -> PlaywrightAmazonProductParser:
    """Return a parser instance for the requested engine name."""

    normalized = (engine or "playwright").strip().lower()
    if normalized in ("", "playwright", "standard"):
        return PlaywrightAmazonProductParser()
    if normalized == "playwright-fast":
        # 快速模式：更短的超时时间，但保留内容展开功能
        perf_logger.info("[create_parser] 创建快速模式解析器")
        return PlaywrightAmazonProductParser(
            headless=True,  # 使用无头模式更快
            wait_until="domcontentloaded",
            navigation_timeout=8.0,
            settle_timeout=0,  # 不需要额外等待
        )
    raise ValueError(
        f"Unsupported parser engine '{engine}'. Expected 'playwright' or 'playwright-fast'."
    )


# ----------------------------------------------------------------------
# Utility helpers
# ----------------------------------------------------------------------


class PerfTimer:
    """简单的性能计时器，用于记录各步骤耗时"""

    def __init__(self, name: str, logger: logging.Logger = perf_logger):
        self.name = name
        self.logger = logger
        self.start_time = None
        self.end_time = None

    def __enter__(self):
        self.start_time = time.perf_counter()
        self.logger.info(f"[START] {self.name}")
        return self

    def __exit__(self, *args):
        self.end_time = time.perf_counter()
        elapsed = (self.end_time - self.start_time) * 1000  # 转换为毫秒
        self.logger.info(f"[DONE] {self.name} - 耗时: {elapsed:.2f}ms")

    @property
    def elapsed_ms(self) -> float:
        if self.start_time is None:
            return 0.0
        end = self.end_time if self.end_time else time.perf_counter()
        return (end - self.start_time) * 1000


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
    perf_logger.info(f"{'='*60}")
    perf_logger.info(f"[Excel] 开始处理 Excel 文件: {input_path}")

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

    perf_logger.info(f"[Excel] 工作表: {worksheet.title}, 数据范围: 第 {data_start_row}-{max_row} 行")
    processed_count = 0
    success_count = 0
    error_count = 0
    excel_start = time.perf_counter()

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

        processed_count += 1
        row_start = time.perf_counter()
        perf_logger.info(f"[Excel [{processed_count}/{max_row - data_start_row + 1}]] 处理行 {row_index}: {url_text[:80]}...")

        product: Optional[AmazonProduct]
        error_message = ""
        try:
            product = parser_instance.parse(url_text)
            success_count += 1
        except Exception as exc:  # pragma: no cover - network/parse failures
            product = None
            error_message = str(exc)
            error_count += 1
            perf_logger.error(f"  [Excel] 错误: {exc}")

        row_elapsed = (time.perf_counter() - row_start) * 1000
        perf_logger.info(f"  [Excel] 行 {row_index} 完成 - 耗时: {row_elapsed:.2f}ms")

        for offset, (_, accessor) in enumerate(result_columns, start=1):
            value = accessor(product, error_message)
            worksheet.cell(row=row_index, column=column + offset, value=value)

    destination = str(output_path) if output_path else source_path
    perf_logger.info(f"[Excel] 保存工作簿到: {destination}")
    workbook.save(destination)
    workbook.close()

    total_elapsed = (time.perf_counter() - excel_start) / 60  # 转换为分钟
    perf_logger.info(f"[Excel] 处理完成! 处理: {processed_count} 行, 成功: {success_count}, 失败: {error_count}")
    perf_logger.info(f"[Excel] 总耗时: {total_elapsed:.2f} 分钟 (平均 {total_elapsed * 60 / processed_count:.2f} 秒/行)")
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
    parser.add_argument(
        "--engine",
        choices=("playwright", "playwright-fast"),
        default="playwright-fast",
        help=(
            "解析引擎选择。'playwright' 为标准模式（完整解析），"
            "'playwright-fast' 为快速模式（推荐，速度更快）。"
        ),
    )
    args = parser.parse_args(args=list(argv) if argv is not None else None)

    product_parser = create_parser(args.engine)

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
