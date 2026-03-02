"""
Tests for Image Processor module.

Test categories:
- Unit tests (no API calls, no heavy deps) — always run
- Integration tests (require API key) — marked with @pytest.mark.integration
"""

import base64
import json
import os
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import pytest

import sys
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from invoice_generator.image_processor import (
    BatchResult,
    OCRDetection,
    OUTPUT_COLUMNS,
    _group_detections_into_lines,
    _records_to_dataframe,
    extract_from_image,
    ocr_text_to_dataframe,
    repair_json,
    validate_extraction,
)


# ─── Fixtures ────────────────────────────────────────────────────────────────

@pytest.fixture
def sample_records():
    """Valid delivery records as would come from an LLM extraction."""
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
        "Rate": [380.0, 280.0],
        "Per": ["Tonne", "Tonne"],
    })


@pytest.fixture
def test_image(tmp_path):
    """Create a simple test image with tabular text using Pillow."""
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (800, 400), "white")
    draw = ImageDraw.Draw(image)

    lines = [
        "Date       Challan  Vehicle  Site      Material  Qty    Rate  Per",
        "16/02/2025  4521     7938    Alpha     10 mm     35.61  380   Tonne",
        "18/02/2025  4576     7938    Beta      20 mm     40.60  280   Tonne",
    ]
    for line_index, line_text in enumerate(lines):
        draw.text((20, 30 + line_index * 40), line_text, fill="black")

    image_path = tmp_path / "test_challan.jpg"
    image.save(image_path)
    return image_path


# ─── JSON Repair Tests ──────────────────────────────────────────────────────

class TestRepairJSON:
    def test_clean_json_array(self):
        text = '[{"Date": "16/02/2025", "Quantity": 35.61}]'
        result = repair_json(text)
        assert len(result) == 1
        assert result[0]["Date"] == "16/02/2025"

    def test_strips_markdown_fences(self):
        text = '```json\n[{"Date": "16/02/2025"}]\n```'
        result = repair_json(text)
        assert len(result) == 1

    def test_extracts_array_from_commentary(self):
        text = 'Here are the records:\n[{"Date": "16/02/2025"}]\nHope this helps!'
        result = repair_json(text)
        assert len(result) == 1

    def test_fixes_trailing_comma(self):
        text = '[{"Date": "16/02/2025"},]'
        result = repair_json(text)
        assert len(result) == 1

    def test_wraps_single_object_in_list(self):
        text = '{"Date": "16/02/2025", "Quantity": 35.61}'
        result = repair_json(text)
        assert len(result) == 1

    def test_extracts_multiple_separate_objects(self):
        text = '{"Date": "16/02/2025"}\n{"Date": "18/02/2025"}'
        result = repair_json(text)
        assert len(result) == 2

    def test_empty_array_returns_empty_list(self):
        assert repair_json("[]") == []

    def test_raises_on_unparseable_text(self):
        with pytest.raises(ValueError, match="Could not parse JSON"):
            repair_json("This is not JSON at all")

    def test_handles_nested_markdown_fences(self):
        text = "```json\n```json\n[{\"Date\": \"16/02/2025\"}]\n```\n```"
        result = repair_json(text)
        assert len(result) == 1


# ─── Records to DataFrame Conversion ────────────────────────────────────────

class TestRecordsToDataframe:
    def test_valid_records_produce_correct_dataframe(self, sample_records):
        dataframe = _records_to_dataframe(sample_records)
        assert len(dataframe) == 2
        assert list(dataframe.columns) == OUTPUT_COLUMNS

    def test_fills_missing_per_with_tonne(self, sample_records):
        del sample_records[0]["Per"]
        dataframe = _records_to_dataframe(sample_records)
        assert dataframe["Per"].iloc[0] == "Tonne"

    def test_filters_out_english_total_rows(self, sample_records):
        sample_records.append({
            "Date": "", "Challan No.": None, "Vehicle No.": None,
            "Site": "Total", "Material": "", "Quantity": 76.21,
            "Rate": 0, "Per": "",
        })
        dataframe = _records_to_dataframe(sample_records)
        assert len(dataframe) == 2

    def test_filters_out_hindi_total_rows(self, sample_records):
        sample_records.append({
            "Date": "", "Challan No.": None, "Vehicle No.": None,
            "Site": "कुल", "Material": "", "Quantity": 76.21,
            "Rate": 0, "Per": "",
        })
        dataframe = _records_to_dataframe(sample_records)
        assert len(dataframe) == 2

    def test_deduplicates_by_challan_date_material(self, sample_records):
        sample_records.append(sample_records[0].copy())
        dataframe = _records_to_dataframe(sample_records)
        assert len(dataframe) == 2

    def test_coerces_string_numbers(self, sample_records):
        sample_records[0]["Quantity"] = "35.61"
        sample_records[0]["Rate"] = "380"
        dataframe = _records_to_dataframe(sample_records)
        assert dataframe["Quantity"].iloc[0] == 35.61
        assert dataframe["Rate"].iloc[0] == 380.0

    def test_raises_on_empty_records(self):
        with pytest.raises(ValueError, match="No records"):
            _records_to_dataframe([])

    def test_normalizes_lowercase_column_names(self):
        records = [{
            "date": "16/02/2025", "challan no.": 4521,
            "vehicle no.": 7938, "site": "Alpha",
            "material": "10 mm", "quantity": 35.61,
            "rate": 380, "per": "Tonne",
        }]
        dataframe = _records_to_dataframe(records)
        assert len(dataframe) == 1
        # Columns should be normalized to canonical names
        for column in OUTPUT_COLUMNS:
            assert column in dataframe.columns


# ─── Extraction Validation Tests ────────────────────────────────────────────

class TestValidation:
    def test_clean_data_produces_no_warnings(self, sample_dataframe):
        warnings = validate_extraction(sample_dataframe)
        assert len(warnings) == 0

    def test_warns_on_zero_quantity(self, sample_dataframe):
        sample_dataframe.loc[0, "Quantity"] = 0
        warnings = validate_extraction(sample_dataframe)
        assert any("Quantity = 0" in warning for warning in warnings)

    def test_warns_on_zero_rate(self, sample_dataframe):
        sample_dataframe.loc[0, "Rate"] = 0
        warnings = validate_extraction(sample_dataframe)
        assert any("Rate = 0" in warning for warning in warnings)

    def test_warns_on_huge_amount(self, sample_dataframe):
        sample_dataframe.loc[0, "Quantity"] = 10000
        sample_dataframe.loc[0, "Rate"] = 500
        warnings = validate_extraction(sample_dataframe)
        assert any("₹10,00,000" in warning for warning in warnings)

    def test_warns_on_empty_site(self, sample_dataframe):
        sample_dataframe.loc[0, "Site"] = ""
        warnings = validate_extraction(sample_dataframe)
        assert any("empty Site" in warning for warning in warnings)

    def test_warns_on_missing_quantity(self, sample_dataframe):
        sample_dataframe.loc[0, "Quantity"] = None
        warnings = validate_extraction(sample_dataframe)
        assert any("missing Quantity" in warning for warning in warnings)

    def test_warns_on_empty_material(self, sample_dataframe):
        sample_dataframe.loc[0, "Material"] = ""
        warnings = validate_extraction(sample_dataframe)
        assert any("empty Material" in warning for warning in warnings)


# ─── OCR Heuristic Parsing Tests ────────────────────────────────────────────

class TestOCRTextParsing:
    def test_parses_pipe_separated_text(self):
        text = (
            "Date | Challan | Vehicle | Site | Material | Qty | Rate | Per\n"
            "16/02/2025 | 4521 | 7938 | Alpha | 10 mm | 35.61 | 380 | Tonne\n"
            "18/02/2025 | 4576 | 7938 | Beta | 20 mm | 40.60 | 280 | Tonne\n"
        )
        dataframe = ocr_text_to_dataframe(text)
        assert len(dataframe) == 2

    def test_filters_header_row(self):
        text = (
            "Date | Challan | Vehicle | Site | Material | Quantity | Rate | Per\n"
            "16/02/2025 | 4521 | 7938 | Alpha | 10 mm | 35.61 | 380 | Tonne\n"
        )
        dataframe = ocr_text_to_dataframe(text)
        assert len(dataframe) == 1

    def test_filters_total_row(self):
        text = (
            "16/02/2025 | 4521 | 7938 | Alpha | 10 mm | 35.61 | 380 | Tonne\n"
            "Total | | | | | 35.61 | | \n"
        )
        dataframe = ocr_text_to_dataframe(text)
        assert len(dataframe) == 1

    def test_raises_on_no_extractable_records(self):
        with pytest.raises(ValueError, match="Could not extract"):
            ocr_text_to_dataframe("No structured data here at all")

    def test_requires_date_in_row(self):
        text = "Alpha | 10 mm | 35.61 | 380 | Tonne\n"
        with pytest.raises(ValueError):
            ocr_text_to_dataframe(text)


# ─── OCR Detection Dataclass Tests ──────────────────────────────────────────

class TestOCRDetection:
    def test_center_coordinates(self):
        detection = OCRDetection(
            bounding_box=[[0, 0], [100, 0], [100, 50], [0, 50]],
            text="test",
            confidence=0.95,
        )
        assert detection.center_x == 50.0
        assert detection.center_y == 25.0

    def test_grouping_into_lines(self):
        detections = [
            OCRDetection([[0, 0], [50, 0], [50, 20], [0, 20]], "A", 0.9),
            OCRDetection([[60, 0], [110, 0], [110, 20], [60, 20]], "B", 0.9),
            OCRDetection([[0, 40], [50, 40], [50, 60], [0, 60]], "C", 0.9),
            OCRDetection([[60, 40], [110, 40], [110, 60], [60, 60]], "D", 0.9),
        ]
        lines = _group_detections_into_lines(detections)
        assert len(lines) == 2
        assert [d.text for d in lines[0]] == ["A", "B"]
        assert [d.text for d in lines[1]] == ["C", "D"]

    def test_grouping_sorts_left_to_right(self):
        detections = [
            OCRDetection([[60, 0], [110, 0], [110, 20], [60, 20]], "B", 0.9),
            OCRDetection([[0, 0], [50, 0], [50, 20], [0, 20]], "A", 0.9),
        ]
        lines = _group_detections_into_lines(detections)
        assert len(lines) == 1
        assert [d.text for d in lines[0]] == ["A", "B"]

    def test_grouping_empty_returns_empty(self):
        assert _group_detections_into_lines([]) == []


# ─── Image Preprocessing Tests ──────────────────────────────────────────────

class TestPreprocessing:
    def test_file_not_found_raises(self):
        with pytest.raises(FileNotFoundError):
            extract_from_image("/nonexistent/image.jpg", force_ocr=True)

    def test_unsupported_format_raises(self, tmp_path):
        bad_file = tmp_path / "test.docx"
        bad_file.write_text("not an image")
        with pytest.raises(ValueError, match="Unsupported image format"):
            from invoice_generator.image_processor import preprocess_image
            preprocess_image(bad_file)

    def test_preprocessing_returns_valid_jpeg(self, test_image):
        from invoice_generator.image_processor import preprocess_image
        jpeg_bytes = preprocess_image(test_image)
        assert isinstance(jpeg_bytes, bytes)
        assert len(jpeg_bytes) > 0
        assert jpeg_bytes[:2] == b'\xff\xd8'  # JPEG magic bytes

    def test_base64_encoding_produces_valid_string(self, test_image):
        from invoice_generator.image_processor import encode_image_for_llm
        base64_string = encode_image_for_llm(test_image)
        assert isinstance(base64_string, str)
        decoded_bytes = base64.b64decode(base64_string)
        assert len(decoded_bytes) > 0
        assert decoded_bytes[:2] == b'\xff\xd8'  # Should be valid JPEG


# ─── BatchResult Tests ──────────────────────────────────────────────────────

class TestBatchResult:
    def test_default_values(self):
        result = BatchResult(input_path="img.jpg", invoice_number=100)
        assert result.excel_path is None
        assert result.pdf_path is None
        assert result.error is None
        assert result.record_count == 0

    def test_with_all_fields(self):
        result = BatchResult(
            input_path="img.jpg",
            invoice_number=100,
            excel_path="Invoice_100.xlsx",
            pdf_path="Invoice_100.pdf",
            extracted_data_path="Data_100.xlsx",
            record_count=5,
            error=None,
        )
        assert result.record_count == 5


# ─── Integration Tests (require API key) ────────────────────────────────────

@pytest.mark.integration
class TestLLMIntegration:
    """These tests make real API calls. Run with: pytest -m integration"""

    def test_extract_from_test_image(self, test_image):
        api_key = os.environ.get("OPENROUTER_API_KEY")
        if not api_key:
            pytest.skip("OPENROUTER_API_KEY not set")

        dataframe = extract_from_image(test_image, api_key=api_key)
        assert len(dataframe) > 0
        assert "Date" in dataframe.columns
        assert "Quantity" in dataframe.columns
