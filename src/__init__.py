"""Amazon product parser package."""

from .amazon_product_parser import AmazonProduct, AmazonProductParser, process_excel
from .ui_app import app, create_app

__all__ = ["AmazonProduct", "AmazonProductParser", "process_excel", "create_app", "app"]
__version__ = "0.1.0"
__author__ = "Amazon Parser Contributors"
