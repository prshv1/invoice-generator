"""
Tests for Image Processor module.

Tests are split into:
- Unit tests (no API calls, no heavy deps) — always run
- Integration tests (require API key) — marked with @pytest.mark.integration
- OCR tests (require easyocr + model download) — marked with @pytest.mark.ocr
"""

import json
import os
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from image_processor import (
    OUTPUT_COLUMNS,
    _records_to_dataframe,
    extract_from_image,
    ocr_text_to_dataframe,
    repair_json,
    validate_extraction,
)


# ─── Fixtures ────────────────────────────────────────────────────────────────

@pytest.fixture
def sample_records():
    """Valid delivery records as would come from LLM."""
    return [
        {
            "Date": "16/02/2025",
            "Challan No.": 4521,
            "Vehicle No.": 7938,
            "Site": "Alpha",
            "Material": "10 mm",
            "Quantity": 35.61,
            "Rate": 380,
            "Per": "Tonne",
        },
        {
            "Date": "18/02/2025",
            "Challan No.": 4576,
            "Vehicle No.": 7938,
            "Site": "Beta",
            "Material": "20 mm",
            "Quantity": 40.60,
            "Rate": 280,
            "Per": "Tonne",
        },
    ]


@pytest.fixture
def sample_dataframe():
    """Valid DataFrame for validation testing."""
    return pd.DataFrame({
        "Date": ["16/02/2025", "18/02/2025"],
        "Challan No.": [4521, 4576],
        "Vehicle No.": [7938, 7938],
        "Site": ["Alpha", "Beta"],
        "Material": ["10 mm", "20 mm"],
        "Quantity": [35.61, 40.60],
        "Rate": [380, 280],
        "Per": ["Tonne", "Tonne"],
    })


@pytest.fixture
def test_image(tmp_path):
    """Create a simple test image with text using Pillow."""
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (800, 400), "white")
    draw = ImageDraw.Draw(image)

    # Simple table-like text
    lines = [
        "Date       Challan  Vehicle  Site      Material  Qty    Rate  Per",
        "16/02/2025  4521     7938    Alpha     10 mm     35.61  380   Tonne",
        "18/02/2025  4576     7938    Beta      20 mm     40.60  280   Tonne",
    ]
    for i, line in enumerate(lines):
        draw.text((20, 30 + i * 40), line, fill="black")

    path = tmp_path / "test_challan.jpg"
    image.save(path)
    return path


# ─── JSON Repair Tests ──────────────────────────────────────────────────────

class TestRepairJSON:
    def test_clean_json(self):
        text = '[{"Date": "16/02/2025", "Quantity": 35.61}]'
        result = repair_json(text)
        assert len(result) == 1
        assert result[0]["Date"] == "16/02/2025"

    def test_markdown_fences(self):
        text = '```json\n[{"Date": "16/02/2025"}]\n```'
        result = repair_json(text)
        assert len(result) == 1

    def test_surrounding_commentary(self):
        text = 'Here are the records:\n[{"Date": "16/02/2025"}]\nHope this helps!'
        result = repair_json(text)
        assert len(result) == 1

    def test_trailing_comma(self):
        text = '[{"Date": "16/02/2025"},]'
        result = repair_json(text)
        assert len(result) == 1

    def test_single_object(self):
        """Single object (not array) should be wrapped in a list."""
        text = '{"Date": "16/02/2025", "Quantity": 35.61}'
        result = repair_json(text)
        assert len(result) == 1

    def test_multiple_objects_without_array(self):
        """Multiple objects without enclosing array."""
        text = '{"Date": "16/02/2025"}\n{"Date": "18/02/2025"}'
        result = repair_json(text)
        assert len(result) == 2

    def test_empty_array(self):
        text = "[]"
        result = repair_json(text)
        assert result == []

    def test_garbage_raises_error(self):
        with pytest.raises(ValueError, match="Could not parse JSON"):
            repair_json("This is not JSON at all")

    def test_nested_markdown(self):
        text = "```json\n```json\n[{\"Date\": \"16/02/2025\"}]\n```\n```"
        result = repair_json(text)
        assert len(result) == 1


# ─── Records to DataFrame Tests ─────────────────────────────────────────────

class TestRecordsToDataframe:
    def test_valid_records(self, sample_records):
        df = _records_to_dataframe(sample_records)
        assert len(df) == 2
        assert list(df.columns) == OUTPUT_COLUMNS

    def test_fills_missing_per(self, sample_records):
        sample_records[0].pop("Per")
        df = _records_to_dataframe(sample_records)
        assert df["Per"].iloc[0] == "Tonne"

    def test_filters_total_rows(self, sample_records):
        sample_records.append({
            "Date": "", "Challan No.": None, "Vehicle No.": None,
            "Site": "Total", "Material": "", "Quantity": 76.21,
            "Rate": 0, "Per": "",
        })
        df = _records_to_dataframe(sample_records)
        assert len(df) == 2  # Total row filtered out

    def test_removes_duplicates(self, sample_records):
        sample_records.append(sample_records[0].copy())  # Exact duplicate
        df = _records_to_dataframe(sample_records)
        assert len(df) == 2

    def test_handles_string_numbers(self, sample_records):
        sample_records[0]["Quantity"] = "35.61"
        sample_records[0]["Rate"] = "380"
        df = _records_to_dataframe(sample_records)
        assert df["Quantity"].iloc[0] == 35.61
        assert df["Rate"].iloc[0] == 380.0

    def test_empty_records_raises(self):
        with pytest.raises(ValueError, match="No records"):
            _records_to_dataframe([])

    def test_column_normalization(self):
        """Handles slightly different column names from LLM."""
        records = [{"date": "16/02/2025", "challan no.": 4521,
                     "vehicle no.": 7938, "site": "Alpha",
                     "material": "10 mm", "quantity": 35.61,
                     "rate": 380, "per": "Tonne"}]
        df = _records_to_dataframe(records)
        assert "Date" in df.columns or "date" in df.columns
        assert len(df) == 1

    def test_hindi_total_filtered(self, sample_records):
        """Hindi 'कुल' (total) should be filtered."""
        sample_records.append({
            "Date": "", "Challan No.": None, "Vehicle No.": None,
            "Site": "कुल", "Material": "", "Quantity": 76.21,
            "Rate": 0, "Per": "",
        })
        df = _records_to_dataframe(sample_records)
        assert len(df) == 2


# ─── Validation Tests ───────────────────────────────────────────────────────

class TestValidation:
    def test_clean_data_no_warnings(self, sample_dataframe):
        warnings = validate_extraction(sample_dataframe)
        assert len(warnings) == 0

    def test_warns_zero_quantity(self, sample_dataframe):
        sample_dataframe.loc[0, "Quantity"] = 0
        warnings = validate_extraction(sample_dataframe)
        assert any("Quantity = 0" in w for w in warnings)

    def test_warns_zero_rate(self, sample_dataframe):
        sample_dataframe.loc[0, "Rate"] = 0
        warnings = validate_extraction(sample_dataframe)
        assert any("Rate = 0" in w for w in warnings)

    def test_warns_huge_amount(self, sample_dataframe):
        sample_dataframe.loc[0, "Quantity"] = 10000
        sample_dataframe.loc[0, "Rate"] = 500
        warnings = validate_extraction(sample_dataframe)
        assert any("₹10,00,000" in w for w in warnings)

    def test_warns_empty_site(self, sample_dataframe):
        sample_dataframe.loc[0, "Site"] = ""
        warnings = validate_extraction(sample_dataframe)
        assert any("empty Site" in w for w in warnings)

    def test_warns_missing_quantity(self, sample_dataframe):
        sample_dataframe.loc[0, "Quantity"] = None
        warnings = validate_extraction(sample_dataframe)
        assert any("missing Quantity" in w for w in warnings)


# ─── OCR Heuristic Parsing Tests ────────────────────────────────────────────

class TestOCRParsing:
    def test_pipe_separated_text(self):
        text = (
            "Date | Challan | Vehicle | Site | Material | Qty | Rate | Per\n"
            "16/02/2025 | 4521 | 7938 | Alpha | 10 mm | 35.61 | 380 | Tonne\n"
            "18/02/2025 | 4576 | 7938 | Beta | 20 mm | 40.60 | 280 | Tonne\n"
        )
        df = ocr_text_to_dataframe(text)
        assert len(df) == 2

    def test_filters_header_row(self):
        text = (
            "Date | Challan | Vehicle | Site | Material | Quantity | Rate | Per\n"
            "16/02/2025 | 4521 | 7938 | Alpha | 10 mm | 35.61 | 380 | Tonne\n"
        )
        df = ocr_text_to_dataframe(text)
        assert len(df) == 1
        # Header should not appear as data

    def test_filters_total_row(self):
        text = (
            "16/02/2025 | 4521 | 7938 | Alpha | 10 mm | 35.61 | 380 | Tonne\n"
            "Total | | | | | 35.61 | | \n"
        )
        df = ocr_text_to_dataframe(text)
        assert len(df) == 1

    def test_no_records_raises(self):
        with pytest.raises(ValueError, match="Could not extract"):
            ocr_text_to_dataframe("No structured data here at all")

    def test_date_required_for_row(self):
        text = "Alpha | 10 mm | 35.61 | 380 | Tonne\n"
        with pytest.raises(ValueError):
            ocr_text_to_dataframe(text)


# ─── Image Preprocessing Tests ──────────────────────────────────────────────

class TestPreprocessing:
    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            extract_from_image("/nonexistent/image.jpg", force_ocr=True)

    def test_unsupported_format(self, tmp_path):
        bad_file = tmp_path / "test.docx"
        bad_file.write_text("not an image")
        with pytest.raises(ValueError, match="Unsupported image format"):
            from image_processor import preprocess_image
            preprocess_image(bad_file)

    def test_preprocessing_returns_bytes(self, test_image):
        from image_processor import preprocess_image
        result = preprocess_image(test_image)
        assert isinstance(result, bytes)
        assert len(result) > 0
        # Should be valid JPEG
        assert result[:2] == b'\xff\xd8'  # JPEG magic bytes

    def test_base64_encoding(self, test_image):
        from image_processor import load_raw_image_as_base64
        b64 = load_raw_image_as_base64(test_image)
        assert isinstance(b64, str)
        # Should decode without error
        decoded = base64.b64decode(b64)
        assert len(decoded) > 0


# ─── Integration Tests (require API key) ────────────────────────────────────

@pytest.mark.integration
class TestLLMIntegration:
    """These tests make real API calls. Run with: pytest -m integration"""

    def test_extract_from_test_image(self, test_image):
        api_key = os.environ.get("OPENROUTER_API_KEY")
        if not api_key:
            pytest.skip("OPENROUTER_API_KEY not set")

        df = extract_from_image(test_image, api_key=api_key)
        assert len(df) > 0
        assert "Date" in df.columns
        assert "Quantity" in df.columns


import base64
