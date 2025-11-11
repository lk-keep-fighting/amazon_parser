"""Web UI for visualising imported Amazon product data.

This module exposes a small Flask application that renders an HTML interface for
previewing the contents of Excel workbooks produced during batch parsing.  The
interface is designed to highlight both the raw imported product identifiers and
any structured data extracted by :mod:`src.amazon_product_parser`.  JSON payloads
are formatted for readability and bullet lists are rendered as list items so the
resulting view is easier to scan than a plain spreadsheet.
"""

from __future__ import annotations

import json
import math
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence

from flask import Flask, render_template, request
from openpyxl import load_workbook
from werkzeug.datastructures import FileStorage

# Accept common Excel extensions used by openpyxl
_ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
_DEFAULT_IMPORT_HEADER = "导入数据"
_BASE_DIR = Path(__file__).resolve().parent


def create_app() -> Flask:
    """Create and configure the Flask application."""

    app = Flask(
        __name__,
        template_folder=str(_BASE_DIR / "templates"),
        static_folder=str(_BASE_DIR / "static"),
    )
    app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MiB upload limit

    @app.route("/", methods=["GET", "POST"])
    def index() -> str:
        context: Dict[str, Any] = {
            "file_name": None,
            "imported_rows": None,
            "parsed_rows": None,
            "row_count": 0,
            "error": None,
        }

        if request.method == "POST":
            file = request.files.get("workbook")
            if not file or not file.filename:
                context["error"] = "请选择需要上传的 Excel 文件。"
            else:
                try:
                    preview = _build_preview_from_upload(file)
                except ValueError as exc:
                    context["error"] = str(exc)
                else:
                    context.update(preview)
                    context["file_name"] = file.filename

        return render_template("ui/index.html", **context)

    return app


def _build_preview_from_upload(file: FileStorage) -> Mapping[str, Any]:
    """Parse the uploaded workbook into structures suitable for rendering."""

    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in _ALLOWED_EXTENSIONS:
        raise ValueError("仅支持扩展名为 .xlsx 的 Excel 文件。")

    payload = file.read()
    # Reset the stream position so repeated reads (e.g. by Flask) behave predictably.
    try:
        file.seek(0)
    except Exception:  # pragma: no cover - some streams may not support seek
        pass

    if not payload:
        raise ValueError("上传的文件为空，请重新选择。")

    return _build_preview_from_bytes(payload)


def _build_preview_from_bytes(data: bytes) -> Mapping[str, Any]:
    try:
        workbook = load_workbook(filename=BytesIO(data), data_only=True)
    except Exception as exc:  # pragma: no cover - invalid or corrupted files
        raise ValueError("无法读取 Excel 文件，请确认格式正确。") from exc

    worksheet = workbook.active
    rows = list(worksheet.iter_rows(values_only=True))
    max_columns = worksheet.max_column or max(len(row) for row in rows)
    workbook.close()

    if not rows:
        raise ValueError("工作簿中没有可展示的数据。")

    has_header = _has_header_row(rows[0])
    raw_header = rows[0] if has_header else None
    data_rows = rows[1:] if has_header else rows

    headers = _normalise_headers(raw_header, max_columns)

    imported_rows: List[Dict[str, Any]] = []
    parsed_rows: List[Dict[str, Any]] = []

    start_index = 2 if has_header else 1
    for row_number, raw_row in enumerate(data_rows, start=start_index):
        values = list(raw_row) if raw_row else []
        if len(values) < max_columns:
            values.extend([None] * (max_columns - len(values)))

        source_value = values[0] if values else None
        source_display = _format_source_value(source_value)
        if source_display["type"] != "empty":
            imported_rows.append(
                {
                    "row_number": row_number,
                    "display": source_display,
                }
            )

        cells: List[Dict[str, Any]] = []
        for column_index in range(1, max_columns):
            header = headers[column_index] if column_index < len(headers) else f"列 {column_index + 1}"
            value = values[column_index] if column_index < len(values) else None
            display = _format_cell_value(value)
            if display["type"] == "empty":
                continue
            cells.append({"header": header, "display": display})

        if cells:
            parsed_rows.append(
                {
                    "row_number": row_number,
                    "source": source_display,
                    "cells": cells,
                }
            )

    return {
        "imported_rows": imported_rows,
        "parsed_rows": parsed_rows,
        "row_count": len(data_rows),
    }


def _has_header_row(row: Sequence[Any]) -> bool:
    """Heuristically determine whether the first row contains header labels."""

    if not row:
        return False

    text_candidates = [str(cell).strip() for cell in row if isinstance(cell, str) and str(cell).strip()]
    if not text_candidates:
        return False

    # If any value in the header row looks like actual data, treat it as data.
    if any(_looks_like_url(text) or _looks_like_asin(text) for text in text_candidates):
        return False
    return True


def _normalise_headers(raw_header: Optional[Sequence[Any]], width: int) -> List[str]:
    headers: List[str] = []
    for index in range(width):
        if raw_header and index < len(raw_header):
            value = raw_header[index]
            label = str(value).strip() if value is not None else ""
        else:
            label = ""

        if not label:
            label = _DEFAULT_IMPORT_HEADER if index == 0 else f"列 {index + 1}"
        headers.append(label)
    return headers


def _format_source_value(value: Any) -> Dict[str, Any]:
    text = _stringify(value)
    if not text:
        return {"type": "empty", "text": ""}
    if _looks_like_url(text):
        return {"type": "link", "text": text, "url": text}
    if _looks_like_asin(text):
        asin = text.upper()
        return {
            "type": "asin",
            "text": asin,
            "url": f"https://www.amazon.com/dp/{asin}",
        }
    return {"type": "text", "text": text}


def _format_cell_value(value: Any) -> Dict[str, Any]:
    if value is None:
        return {"type": "empty", "text": ""}

    if isinstance(value, (int, float)):
        return {"type": "number", "text": _stringify(value)}

    if isinstance(value, (list, dict)):
        pretty = json.dumps(value, ensure_ascii=False, indent=2)
        return {"type": "json", "json": pretty}

    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return {"type": "empty", "text": ""}
        if _looks_like_json(stripped):
            try:
                parsed = json.loads(stripped)
            except json.JSONDecodeError:
                pass
            else:
                pretty = json.dumps(parsed, ensure_ascii=False, indent=2)
                return {"type": "json", "json": pretty}
        if "\n" in value:
            lines = [line.strip() for line in value.splitlines() if line.strip()]
            if lines:
                return {"type": "list", "lines": lines}
        if _looks_like_url(stripped):
            return {"type": "link", "text": stripped, "url": stripped}
        return {"type": "text", "text": value}

    return {"type": "text", "text": _stringify(value)}


def _stringify(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if not math.isfinite(value):
            return ""
        if value.is_integer():
            return str(int(value))
    return str(value)


def _looks_like_url(value: str) -> bool:
    lower = value.strip().lower()
    return lower.startswith("http://") or lower.startswith("https://")


def _looks_like_asin(value: str) -> bool:
    text = value.strip().upper()
    return len(text) == 10 and text.isalnum()


def _looks_like_json(value: str) -> bool:
    stripped = value.strip()
    return (stripped.startswith("{") and stripped.endswith("}")) or (
        stripped.startswith("[") and stripped.endswith("]")
    )


# Expose a default application instance for convenience with `flask run`.
app = create_app()


__all__ = ["create_app", "app"]


if __name__ == "__main__":  # pragma: no cover - manual execution helper
    app.run(host="127.0.0.1", port=5000, debug=False)
