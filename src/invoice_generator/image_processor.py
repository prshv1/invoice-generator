#!/usr/bin/env python3
"""
Image Processor — Extract delivery data from challan/bill images.

Converts photos of handwritten or printed delivery challans into structured
Excel data that feeds into the Invoice Generator pipeline.

Architecture (3-tier fallback):
    1. LLM Vision (primary) — sends image to free OpenRouter vision model
    2. OCR + LLM Text (fallback) — EasyOCR reads characters, LLM structures them
    3. OCR + Heuristics (offline) — pure-offline spatial parsing, no API needed

CLI Usage:
    python3 image_processor.py --image challan.jpg --output data.xlsx
    python3 image_processor.py --image challan.jpg -i 178 --pdf
    python3 image_processor.py --batch ./photos/ --start 178 --pdf --output-dir ./invoices/

Module Usage:
    from image_processor import extract_from_image, batch_process_images
    df = extract_from_image("challan.jpg")
    results = batch_process_images("./photos/", start_num=178, pdf=True)
"""

__version__ = "0.3.0"

__all__ = [
    "extract_from_image",
    "extract_from_images",
    "batch_process_images",
    "images_to_invoice",
    "validate_extraction",
    "repair_json",
]

import argparse
import base64
import io
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Union

import pandas as pd
from dotenv import load_dotenv

# Load .env file for API keys
load_dotenv()


# ─── Constants ───────────────────────────────────────────────────────────────

OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"

# Model cascade: try each in order until one succeeds.
# Verified available via OpenRouter /api/v1/models endpoint (March 2026).
LLM_MODELS = [
    "google/gemma-3-27b-it:free",                    # Best quality free vision model
    "mistralai/mistral-small-3.1-24b-instruct:free",  # Strong backup
    "google/gemma-3-12b-it:free",                     # Lighter, faster
    "nvidia/nemotron-nano-12b-v2-vl:free",            # Supports video too
    "google/gemma-3-4b-it:free",                      # Smallest, last resort
]

# Image processing limits
MAX_IMAGE_WIDTH_PX = 2000      # Resize images wider than this before processing
JPEG_COMPRESSION_QUALITY = 85  # Quality for JPEG encoding (bandwidth vs. clarity)
API_TIMEOUT_SECONDS = 60       # Max wait for an LLM API response
MAX_RETRIES_PER_MODEL = 3      # Attempts per model before cascading to next
OCR_CONFIDENCE_THRESHOLD = 0.80  # Below this, OCR results are flagged as unreliable
Y_TOLERANCE_PX = 15            # Pixel tolerance for grouping OCR words into rows

SUPPORTED_IMAGE_EXTENSIONS = frozenset({
    ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp",
})

# The expected output schema — matches the Invoice Generator's input format
OUTPUT_COLUMNS = [
    "Date", "Challan No.", "Vehicle No.", "Site",
    "Material", "Quantity", "Rate", "Per",
]

# EXIF orientation codes → degrees of rotation needed to fix
EXIF_ROTATION_MAP = {
    3: 180,   # Upside down
    6: 270,   # Rotated 90° clockwise
    8: 90,    # Rotated 90° counter-clockwise
}

# Keywords that identify summary/total rows to filter out
TOTAL_ROW_KEYWORDS = frozenset({
    "total", "sum", "grand", "subtotal", "jodh", "कुल",
})

# Keywords that identify header rows to skip during OCR parsing
HEADER_KEYWORDS = frozenset({
    "date", "challan", "vehicle", "site", "material",
    "quantity", "rate", "per", "amount", "total", "sum", "grand",
})


# ─── LLM Prompts ────────────────────────────────────────────────────────────

VISION_EXTRACTION_PROMPT = """You are an expert at reading Indian delivery challans, bills, and handwritten records.

Extract ALL delivery records from this image into a JSON array.
Each record MUST have these exact fields:

{
  "Date": "dd/mm/yyyy",
  "Challan No.": number_or_null,
  "Vehicle No.": number_or_null,
  "Site": "location name",
  "Material": "material type (e.g. 10 mm, 20 mm, C. Sand, Cement)",
  "Quantity": decimal_number,
  "Rate": decimal_number,
  "Per": "unit (usually Tonne)"
}

Rules:
- Use dd/mm/yyyy format for ALL dates (Indian standard)
- If a value is unclear but you can make a reasonable guess, do so
- Do NOT include total/summary rows (rows labeled "Total", "Sum", etc.)
- Do NOT invent records that aren't in the image
- If the Per/unit column is missing, default to "Tonne"
- Return ONLY the raw JSON array — no markdown fences, no commentary"""

TEXT_STRUCTURING_PROMPT = """You are an expert at structuring Indian delivery challan data.

Below is raw OCR text extracted from a delivery challan image. The text may be
messy, with words out of order or partially misread.

Structure this into a JSON array where each delivery record has these exact fields:
{
  "Date": "dd/mm/yyyy",
  "Challan No.": number_or_null,
  "Vehicle No.": number_or_null,
  "Site": "location name",
  "Material": "material type",
  "Quantity": decimal_number,
  "Rate": decimal_number,
  "Per": "unit (usually Tonne)"
}

Rules:
- Use dd/mm/yyyy format for dates
- Do NOT include total/summary rows
- Return ONLY the raw JSON array

OCR Text:
"""


# ─── Data Structures ────────────────────────────────────────────────────────

@dataclass
class OCRDetection:
    """A single word detected by OCR, with its position and confidence."""
    bounding_box: list
    text: str
    confidence: float

    @property
    def center_x(self) -> float:
        return sum(point[0] for point in self.bounding_box) / 4

    @property
    def center_y(self) -> float:
        return sum(point[1] for point in self.bounding_box) / 4


@dataclass
class BatchResult:
    """Result of processing one image in a batch."""
    input_path: str
    invoice_number: int
    excel_path: Optional[str] = None
    pdf_path: Optional[str] = None
    extracted_data_path: Optional[str] = None
    record_count: int = 0
    error: Optional[str] = None


# ─── Image Preprocessing ────────────────────────────────────────────────────

def _fix_exif_rotation(pil_image):
    """
    Auto-rotate an image based on EXIF orientation metadata.

    Phone cameras embed orientation in EXIF tags rather than actually rotating
    the pixel data. Without this fix, landscape photos appear sideways.
    """
    from PIL import ExifTags

    try:
        exif_data = pil_image._getexif()
        if not exif_data:
            return pil_image

        orientation_key = next(
            (key for key, name in ExifTags.TAGS.items() if name == "Orientation"),
            None,
        )
        if orientation_key and orientation_key in exif_data:
            rotation_degrees = EXIF_ROTATION_MAP.get(exif_data[orientation_key])
            if rotation_degrees:
                return pil_image.rotate(rotation_degrees, expand=True)
    except (AttributeError, StopIteration):
        pass  # No EXIF data available

    return pil_image


def _resize_if_needed(pil_image, max_width: int = MAX_IMAGE_WIDTH_PX):
    """Resize image proportionally if it exceeds max_width."""
    width, height = pil_image.size
    if width <= max_width:
        return pil_image

    scale_factor = max_width / width
    new_dimensions = (max_width, int(height * scale_factor))
    from PIL import Image
    return pil_image.resize(new_dimensions, Image.LANCZOS)


def preprocess_image(image_path: Union[str, Path]) -> bytes:
    """
    Load, preprocess, and return an image as JPEG bytes optimized for OCR.

    Pipeline:
    1. Load image and fix EXIF rotation
    2. Resize to max width (faster OCR, less memory)
    3. Convert to grayscale
    4. Denoise (removes paper texture speckles)
    5. Adaptive threshold (improves contrast under uneven lighting)

    Returns:
        JPEG-encoded bytes of the preprocessed image.

    Raises:
        FileNotFoundError: If the image file doesn't exist.
        ValueError: If the file format is unsupported.
    """
    import cv2
    import numpy as np
    from PIL import Image

    file_path = Path(image_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Image not found: {file_path}")

    if file_path.suffix.lower() not in SUPPORTED_IMAGE_EXTENSIONS:
        raise ValueError(
            f"Unsupported image format: {file_path.suffix}. "
            f"Supported: {', '.join(sorted(SUPPORTED_IMAGE_EXTENSIONS))}"
        )

    pil_image = Image.open(file_path)
    pil_image = _fix_exif_rotation(pil_image)

    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")

    pil_image = _resize_if_needed(pil_image)

    # Convert to OpenCV format for image processing
    cv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    grayscale = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)
    denoised = cv2.fastNlMeansDenoising(grayscale, h=10)
    thresholded = cv2.adaptiveThreshold(
        denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, blockSize=11, C=2,
    )

    _, jpeg_bytes = cv2.imencode(
        ".jpg", thresholded,
        [cv2.IMWRITE_JPEG_QUALITY, JPEG_COMPRESSION_QUALITY],
    )
    return jpeg_bytes.tobytes()


def encode_image_for_llm(image_path: Union[str, Path]) -> str:
    """
    Encode an image as base64 JPEG for the LLM vision API.

    Unlike preprocess_image(), this does minimal processing — only EXIF
    rotation and resize. LLMs interpret images better without heavy
    binarization or thresholding.

    Returns:
        Base64-encoded JPEG string.
    """
    from PIL import Image

    file_path = Path(image_path)
    pil_image = Image.open(file_path)
    pil_image = _fix_exif_rotation(pil_image)

    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")

    pil_image = _resize_if_needed(pil_image)

    buffer = io.BytesIO()
    pil_image.save(buffer, format="JPEG", quality=JPEG_COMPRESSION_QUALITY)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


# ─── JSON Repair ─────────────────────────────────────────────────────────────

def repair_json(raw_text: str) -> list[dict]:
    """
    Parse JSON from potentially messy LLM output.

    LLMs often return JSON wrapped in markdown fences, surrounded by
    commentary, or with trailing commas. This function handles all of that.

    Repair strategies (tried in order):
    1. Direct parse after stripping markdown fences
    2. Extract JSON array from surrounding text via regex
    3. Fix trailing commas and re-parse
    4. Extract individual JSON objects and collect them

    Returns:
        List of parsed dicts.

    Raises:
        ValueError: If the text cannot be parsed as JSON by any strategy.
    """
    text = raw_text.strip()

    # Strip markdown code fences (```json ... ```)
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s*```$', '', text, flags=re.MULTILINE)
    text = text.strip()

    # Strategy 1: Direct parse
    try:
        result = json.loads(text)
        if isinstance(result, list):
            return result
        if isinstance(result, dict):
            return [result]
    except json.JSONDecodeError:
        pass

    # Strategy 2: Extract JSON array from surrounding commentary
    array_match = re.search(r'\[[\s\S]*\]', text)
    if array_match:
        try:
            result = json.loads(array_match.group())
            if isinstance(result, list):
                return result
        except json.JSONDecodeError:
            pass

    # Strategy 3: Fix trailing commas (,] → ] and ,} → })
    comma_fixed = re.sub(r',\s*([}\]])', r'\1', text)
    try:
        result = json.loads(comma_fixed)
        if isinstance(result, list):
            return result
    except json.JSONDecodeError:
        pass

    # Strategy 4: Extract individual JSON objects
    individual_objects = re.findall(r'\{[^{}]+\}', text)
    if individual_objects:
        parsed_objects = []
        for object_string in individual_objects:
            try:
                parsed_objects.append(json.loads(object_string))
            except json.JSONDecodeError:
                continue
        if parsed_objects:
            return parsed_objects

    raise ValueError(f"Could not parse JSON from LLM response:\n{raw_text[:500]}")


# ─── LLM API Layer ──────────────────────────────────────────────────────────

def _call_openrouter(
    messages: list[dict],
    model: str,
    api_key: str,
) -> str:
    """
    Make a single chat completion call to the OpenRouter API.

    Args:
        messages: Chat messages in OpenAI format.
        model: Model identifier (e.g., "google/gemma-3-27b-it:free").
        api_key: OpenRouter API key.

    Returns:
        The model's response text.

    Raises:
        requests.HTTPError: On 4xx/5xx responses.
        requests.Timeout: If the request exceeds API_TIMEOUT_SECONDS.
        ValueError: If the response contains no choices.
    """
    import requests

    response = requests.post(
        OPENROUTER_API_URL,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com/prshv1/invoice-generator",
            "X-Title": "Invoice Generator",
        },
        json={
            "model": model,
            "messages": messages,
            "temperature": 0.1,  # Low temperature for factual extraction
            "max_tokens": 4096,
        },
        timeout=API_TIMEOUT_SECONDS,
    )

    response.raise_for_status()
    response_data = response.json()

    if "choices" not in response_data or len(response_data["choices"]) == 0:
        raise ValueError(f"Empty response from model {model}: {response_data}")

    return response_data["choices"][0]["message"]["content"]


def _call_with_cascade(
    messages: list[dict],
    api_key: str,
    models: Optional[list[str]] = None,
    on_json_error_suffix: Optional[str] = None,
) -> list[dict]:
    """
    Call the LLM API, cascading through models on failure.

    Handles rate limits (429) with exponential backoff, model downtime
    (502/503) by skipping to the next model, and JSON parse errors with
    an optional retry suffix.

    Args:
        messages: Chat messages to send.
        api_key: OpenRouter API key.
        models: Model IDs to try in order. Defaults to LLM_MODELS.
        on_json_error_suffix: Text appended to prompt on JSON parse retry.

    Returns:
        Parsed list of record dicts from the LLM response.

    Raises:
        RuntimeError: If all models and retries are exhausted.
    """
    import requests

    if models is None:
        models = LLM_MODELS.copy()

    last_error = None

    for model in models:
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                print(
                    f"  🤖 Trying {model} (attempt {attempt + 1})...",
                    file=sys.stderr,
                )
                response_text = _call_openrouter(messages, model, api_key)
                return repair_json(response_text)

            except requests.exceptions.HTTPError as http_error:
                status_code = (
                    http_error.response.status_code if http_error.response else None
                )

                if status_code == 429:
                    wait_seconds = 2 ** attempt
                    print(
                        f"  ⏳ Rate limited, waiting {wait_seconds}s...",
                        file=sys.stderr,
                    )
                    time.sleep(wait_seconds)
                    last_error = http_error
                    continue

                if status_code in (502, 503):
                    print(
                        f"  ⚠️  {model} unavailable ({status_code}), trying next...",
                        file=sys.stderr,
                    )
                    last_error = http_error
                    break  # Skip to next model

                last_error = http_error
                break

            except requests.exceptions.Timeout:
                print(f"  ⏰ Timeout on {model}, retrying...", file=sys.stderr)
                last_error = TimeoutError(f"Timeout calling {model}")
                continue

            except (ValueError, json.JSONDecodeError) as parse_error:
                print(f"  ⚠️  JSON parse error: {parse_error}", file=sys.stderr)
                last_error = parse_error

                # On JSON error, retry with a stricter prompt suffix
                if attempt < MAX_RETRIES_PER_MODEL - 1 and on_json_error_suffix:
                    content = messages[0].get("content")
                    if isinstance(content, list):
                        content[0]["text"] += on_json_error_suffix
                    elif isinstance(content, str):
                        messages[0]["content"] += on_json_error_suffix
                continue

            except Exception as unexpected_error:
                last_error = unexpected_error
                break

    raise RuntimeError(
        f"All LLM models failed. Last error: {last_error}\n"
        f"Models tried: {', '.join(models)}"
    )


# ─── LLM Extraction (Primary Path) ──────────────────────────────────────────

def extract_with_llm_vision(
    image_path: Union[str, Path],
    api_key: str,
    models: Optional[list[str]] = None,
) -> pd.DataFrame:
    """
    Extract delivery records by sending the image directly to an LLM vision model.

    This is the primary extraction method — highest accuracy, especially
    for handwritten text where OCR struggles.

    Args:
        image_path: Path to the challan image.
        api_key: OpenRouter API key.
        models: Model IDs to try in cascade order.

    Returns:
        DataFrame with columns matching OUTPUT_COLUMNS.
    """
    image_base64 = encode_image_for_llm(image_path)

    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": VISION_EXTRACTION_PROMPT},
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/jpeg;base64,{image_base64}",
                    },
                },
            ],
        }
    ]

    records = _call_with_cascade(
        messages, api_key, models,
        on_json_error_suffix=(
            "\n\nIMPORTANT: Return ONLY valid JSON. "
            "No markdown, no explanation, just the array."
        ),
    )

    dataframe = _records_to_dataframe(records)
    print(
        f"  ✅ Extracted {len(dataframe)} records via LLM vision",
        file=sys.stderr,
    )
    return dataframe


def extract_with_llm_text(
    ocr_text: str,
    api_key: str,
    models: Optional[list[str]] = None,
) -> pd.DataFrame:
    """
    Send raw OCR text to an LLM for structuring into table format.

    Middle-ground fallback: OCR handles character recognition (free, offline),
    LLM handles understanding the table structure. Cheaper than sending images
    since text tokens cost less than image tokens.
    """
    messages = [
        {
            "role": "user",
            "content": TEXT_STRUCTURING_PROMPT + ocr_text,
        }
    ]

    records = _call_with_cascade(messages, api_key, models)
    return _records_to_dataframe(records)


# ─── OCR Extraction (Fallback Path) ─────────────────────────────────────────

# Module-level cache for the EasyOCR reader (first load downloads ~500MB)
_ocr_reader_cache = None


def _get_ocr_reader():
    """
    Get the EasyOCR reader instance, creating it on first use.

    The reader is cached at module level because instantiation downloads
    model weights (~500MB) and takes several seconds.
    """
    global _ocr_reader_cache

    if _ocr_reader_cache is None:
        import easyocr
        print(
            "  📦 Loading OCR model (first run downloads ~500MB)...",
            file=sys.stderr,
        )
        _ocr_reader_cache = easyocr.Reader(["en", "hi"], verbose=False)

    return _ocr_reader_cache


def extract_with_ocr(image_path: Union[str, Path]) -> tuple[str, float]:
    """
    Extract text from an image using EasyOCR with spatial layout preservation.

    OCR detections are grouped into lines based on vertical position (Y-coordinate)
    and sorted left-to-right within each line, producing pipe-separated text
    that approximates the original table layout.

    Args:
        image_path: Path to the challan image.

    Returns:
        Tuple of (structured_text, median_confidence), where structured_text
        has one line per detected row with fields separated by " | ", and
        median_confidence is the overall OCR quality score (0.0–1.0).
    """
    import cv2
    import numpy as np

    reader = _get_ocr_reader()
    preprocessed_bytes = preprocess_image(image_path)

    # Decode JPEG bytes into an image array for EasyOCR
    image_array = np.frombuffer(preprocessed_bytes, np.uint8)
    grayscale_image = cv2.imdecode(image_array, cv2.IMREAD_GRAYSCALE)

    # EasyOCR returns: [(bounding_box, text, confidence), ...]
    raw_detections = reader.readtext(grayscale_image, detail=1)

    if not raw_detections:
        return "", 0.0

    # Convert raw tuples to typed dataclass instances
    detections = [
        OCRDetection(bounding_box=bbox, text=text, confidence=conf)
        for bbox, text, conf in raw_detections
    ]

    # Compute median confidence across all detections
    median_confidence = float(
        np.median([detection.confidence for detection in detections])
    )

    # Group detections into lines and build structured text
    lines = _group_detections_into_lines(detections)
    structured_text = "\n".join(
        " | ".join(detection.text for detection in line)
        for line in lines
    )

    return structured_text, median_confidence


def _group_detections_into_lines(
    detections: list[OCRDetection],
    y_tolerance: int = Y_TOLERANCE_PX,
) -> list[list[OCRDetection]]:
    """
    Group OCR detections into lines based on vertical proximity.

    Words whose center Y-coordinates are within y_tolerance pixels of each
    other are placed on the same line. Each line is then sorted left-to-right
    by center X-coordinate.

    This spatial grouping reconstructs the table row structure that OCR
    destroys when it returns an unordered list of word detections.
    """
    if not detections:
        return []

    # Sort all detections by vertical position
    sorted_detections = sorted(detections, key=lambda d: d.center_y)

    lines: list[list[OCRDetection]] = [[sorted_detections[0]]]

    for detection in sorted_detections[1:]:
        previous_detection = lines[-1][-1]

        if abs(detection.center_y - previous_detection.center_y) <= y_tolerance:
            # Same line — append
            lines[-1].append(detection)
        else:
            # New line
            lines.append([detection])

    # Sort each line left-to-right
    for line in lines:
        line.sort(key=lambda d: d.center_x)

    return lines


def ocr_text_to_dataframe(ocr_text: str) -> pd.DataFrame:
    """
    Parse structured OCR text into a DataFrame using heuristic rules.

    This is the pure-offline fallback for when no LLM API is available.
    It identifies field types by pattern matching:
    - Dates: dd/mm/yyyy or dd-mm-yyyy patterns
    - Numbers: sequences of digits (mapped to challan, vehicle, qty, rate)
    - Text: site names, material types

    Rules:
    - A row must contain a date to be considered valid
    - Header and totals rows are filtered out
    - Rows with fewer than 4 detected fields are skipped

    Returns:
        DataFrame with OUTPUT_COLUMNS.

    Raises:
        ValueError: If no valid records could be extracted.
    """
    text_lines = [
        line.strip()
        for line in ocr_text.strip().split("\n")
        if line.strip()
    ]

    date_pattern = re.compile(r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}')
    pure_number_pattern = re.compile(r'^\d+\.?\d*$')

    records = []

    for line in text_lines:
        # Split by pipe separator (our OCR output format) or tabs
        fields = [field.strip() for field in re.split(r'[|\t]', line) if field.strip()]

        # Skip header and summary rows
        line_lower = " ".join(fields).lower()
        if any(keyword in line_lower for keyword in HEADER_KEYWORDS):
            continue

        # Need at least 4 fields for a meaningful data row
        if len(fields) < 4:
            continue

        # A valid data row must contain a date
        detected_date = None
        for field in fields:
            date_match = date_pattern.search(field)
            if date_match:
                detected_date = date_match.group()
                break

        if detected_date is None:
            continue

        # Classify remaining fields as numeric or text
        numeric_values = []
        text_values = []
        for field in fields:
            if field == detected_date:
                continue
            cleaned = field.replace(",", "").strip()
            if pure_number_pattern.match(cleaned):
                numeric_values.append(float(cleaned))
            else:
                text_values.append(field)

        # Map values to columns by position and type:
        # Numerics (in order): challan_no, vehicle_no, ..., quantity, rate
        # Text (in order): site, material, per
        record = {
            "Date": detected_date,
            "Challan No.": numeric_values[0] if len(numeric_values) > 0 else None,
            "Vehicle No.": numeric_values[1] if len(numeric_values) > 1 else None,
            "Site": text_values[0] if len(text_values) > 0 else "",
            "Material": text_values[1] if len(text_values) > 1 else "",
            "Quantity": (
                numeric_values[-2] if len(numeric_values) > 3
                else numeric_values[2] if len(numeric_values) > 2
                else 0
            ),
            "Rate": numeric_values[-1] if len(numeric_values) > 2 else 0,
            "Per": text_values[2] if len(text_values) > 2 else "Tonne",
        }

        records.append(record)

    if not records:
        raise ValueError(
            "Could not extract any records from OCR text using heuristics. "
            "The image may be too blurry or not contain tabular delivery data."
        )

    return pd.DataFrame(records, columns=OUTPUT_COLUMNS)


# ─── Data Validation ─────────────────────────────────────────────────────────

def _records_to_dataframe(records: list[dict]) -> pd.DataFrame:
    """
    Convert a list of raw extraction records into a clean, validated DataFrame.

    This is the shared validation pipeline used by all extraction strategies
    (LLM vision, LLM text, OCR heuristics). It ensures consistent output
    regardless of extraction method.

    Validation steps:
    1. Normalize column names (case-insensitive matching)
    2. Ensure all required columns exist (fill defaults where missing)
    3. Coerce numeric columns (Quantity, Rate, Challan No., Vehicle No.)
    4. Fill missing Per/unit with "Tonne"
    5. Filter out totals/summary rows (English + Hindi keywords)
    6. Remove rows where both Quantity and Rate are missing
    7. Deduplicate by Challan No. + Date + Material

    Raises:
        ValueError: If no records provided or all filtered out.
    """
    if not records:
        raise ValueError("No records to convert.")

    dataframe = pd.DataFrame(records)

    # Normalize column names (LLMs sometimes return lowercase or underscored)
    column_mapping = {}
    for expected_column in OUTPUT_COLUMNS:
        for actual_column in dataframe.columns:
            if actual_column.lower().replace("_", " ").strip() == expected_column.lower():
                column_mapping[actual_column] = expected_column
                break
    dataframe = dataframe.rename(columns=column_mapping)

    # Ensure all required columns exist
    for column in OUTPUT_COLUMNS:
        if column not in dataframe.columns:
            default = None if column in ("Challan No.", "Vehicle No.") else ""
            dataframe[column] = default

    # Keep only expected columns, in the canonical order
    dataframe = dataframe[[col for col in OUTPUT_COLUMNS if col in dataframe.columns]]

    # Fill missing Per/unit with default
    if "Per" in dataframe.columns:
        dataframe["Per"] = dataframe["Per"].fillna("Tonne").replace("", "Tonne")

    # Coerce numeric columns
    for column in ["Quantity", "Rate", "Challan No.", "Vehicle No."]:
        if column in dataframe.columns:
            dataframe[column] = pd.to_numeric(dataframe[column], errors="coerce")

    # Filter out totals/summary rows
    if "Site" in dataframe.columns:
        is_data_row = dataframe.apply(
            lambda row: not any(
                keyword in str(cell_value).lower()
                for cell_value in row.values
                for keyword in TOTAL_ROW_KEYWORDS
            ),
            axis=1,
        )
        dataframe = dataframe[is_data_row]

    # Drop rows where both Quantity and Rate are missing
    dataframe = dataframe.dropna(subset=["Quantity", "Rate"], how="all")

    # Deduplicate rows with the same Challan No. + Date + Material
    if dataframe["Challan No."].notna().any():
        row_count_before = len(dataframe)
        dataframe = dataframe.drop_duplicates(
            subset=["Date", "Challan No.", "Material"], keep="first",
        )
        duplicates_removed = row_count_before - len(dataframe)
        if duplicates_removed > 0:
            print(
                f"  ⚠️  Removed {duplicates_removed} duplicate row(s)",
                file=sys.stderr,
            )

    dataframe = dataframe.reset_index(drop=True)

    if dataframe.empty:
        raise ValueError("All records were filtered out during validation.")

    return dataframe


def validate_extraction(dataframe: pd.DataFrame) -> list[str]:
    """
    Run post-extraction sanity checks and return a list of warning messages.

    These warnings don't block processing — they alert the user to values
    that may need manual verification.

    Checks performed:
    - Missing Quantity or Rate values
    - Zero Quantity or Rate values
    - Suspiciously large line amounts (> ₹10,00,000)
    - Empty Site or Material fields
    """
    warnings = []

    # Missing values
    missing_quantity = int(dataframe["Quantity"].isna().sum())
    missing_rate = int(dataframe["Rate"].isna().sum())
    if missing_quantity > 0:
        warnings.append(f"{missing_quantity} row(s) have missing Quantity")
    if missing_rate > 0:
        warnings.append(f"{missing_rate} row(s) have missing Rate")

    # Zero values
    zero_quantity = int((dataframe["Quantity"] == 0).sum())
    zero_rate = int((dataframe["Rate"] == 0).sum())
    if zero_quantity > 0:
        warnings.append(f"{zero_quantity} row(s) have Quantity = 0")
    if zero_rate > 0:
        warnings.append(f"{zero_rate} row(s) have Rate = 0")

    # Suspiciously large amounts
    line_amounts = dataframe["Quantity"].fillna(0) * dataframe["Rate"].fillna(0)
    large_amounts = line_amounts[line_amounts > 1_000_000]
    if len(large_amounts) > 0:
        warnings.append(
            f"{len(large_amounts)} row(s) have Amount > ₹10,00,000 — verify these are correct"
        )

    # Empty fields
    empty_sites = int((dataframe["Site"].astype(str).str.strip() == "").sum())
    empty_materials = int((dataframe["Material"].astype(str).str.strip() == "").sum())
    if empty_sites > 0:
        warnings.append(f"{empty_sites} row(s) have empty Site")
    if empty_materials > 0:
        warnings.append(f"{empty_materials} row(s) have empty Material")

    return warnings


def _print_warnings(warnings: list[str]) -> None:
    """Print validation warnings to stderr."""
    for warning in warnings:
        print(f"  ⚠️  {warning}", file=sys.stderr)


# ─── Public API ──────────────────────────────────────────────────────────────

def extract_from_image(
    image_path: Union[str, Path],
    api_key: Optional[str] = None,
    force_llm: bool = False,
    force_ocr: bool = False,
) -> pd.DataFrame:
    """
    Extract delivery records from a single challan image.

    Tries three strategies in order, falling back gracefully:
    1. LLM vision — sends image to free OpenRouter model (best accuracy)
    2. OCR + LLM text — EasyOCR extracts text, LLM structures it (cheaper)
    3. Pure OCR heuristics — offline spatial parsing (no API needed)

    Args:
        image_path: Path to the challan image file.
        api_key: OpenRouter API key. Reads from OPENROUTER_API_KEY env var if None.
        force_llm: Only use LLM strategies (fail if API unavailable).
        force_ocr: Only use OCR (skip LLM entirely, useful for offline work).

    Returns:
        DataFrame with columns: Date, Challan No., Vehicle No., Site,
        Material, Quantity, Rate, Per.

    Raises:
        FileNotFoundError: If the image doesn't exist.
        RuntimeError: If all extraction strategies fail.
    """
    file_path = Path(image_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Image not found: {file_path}")

    if api_key is None:
        api_key = os.environ.get("OPENROUTER_API_KEY")

    print(f"\n📸 Processing: {file_path.name}", file=sys.stderr)

    # ── Strategy 1: LLM Vision (highest accuracy) ──
    if not force_ocr and api_key:
        try:
            dataframe = extract_with_llm_vision(file_path, api_key)
            _print_warnings(validate_extraction(dataframe))
            return dataframe
        except Exception as llm_vision_error:
            print(f"  ⚠️  LLM vision failed: {llm_vision_error}", file=sys.stderr)
            if force_llm:
                raise

    # ── Strategy 2: OCR + LLM text structuring (middle ground) ──
    if not force_ocr and api_key:
        try:
            print("  🔍 Falling back to OCR + LLM text...", file=sys.stderr)
            ocr_text, confidence = extract_with_ocr(file_path)
            print(f"  📊 OCR confidence: {confidence:.0%}", file=sys.stderr)

            if ocr_text.strip():
                dataframe = extract_with_llm_text(ocr_text, api_key)
                _print_warnings(validate_extraction(dataframe))
                return dataframe
        except Exception as ocr_llm_error:
            print(f"  ⚠️  OCR + LLM text failed: {ocr_llm_error}", file=sys.stderr)

    # ── Strategy 3: Pure OCR with heuristics (offline) ──
    if api_key and force_llm:
        raise RuntimeError(
            "LLM extraction failed and force_llm=True prevents OCR fallback."
        )

    print("  🔍 Falling back to pure OCR (offline mode)...", file=sys.stderr)
    ocr_text, confidence = extract_with_ocr(file_path)
    print(f"  📊 OCR confidence: {confidence:.0%}", file=sys.stderr)

    if confidence < OCR_CONFIDENCE_THRESHOLD:
        print(
            f"  ⚠️  Low OCR confidence ({confidence:.0%} < "
            f"{OCR_CONFIDENCE_THRESHOLD:.0%}). Results may be unreliable.",
            file=sys.stderr,
        )

    dataframe = ocr_text_to_dataframe(ocr_text)
    _print_warnings(validate_extraction(dataframe))
    return dataframe


def extract_from_images(
    image_paths: list[Union[str, Path]],
    api_key: Optional[str] = None,
    **kwargs: Any,
) -> pd.DataFrame:
    """
    Extract delivery records from multiple images and combine into one DataFrame.

    Useful for multi-page challans or combining several deliveries into a
    single data file. Each image is processed independently, then all results
    are concatenated.

    Args:
        image_paths: List of image file paths.
        api_key: OpenRouter API key.
        **kwargs: Passed through to extract_from_image (force_llm, force_ocr).

    Returns:
        Combined DataFrame with all records from all images.

    Raises:
        RuntimeError: If no images were successfully processed.
    """
    successful_frames = []
    failed_images = []

    for image_path in image_paths:
        try:
            dataframe = extract_from_image(image_path, api_key=api_key, **kwargs)
            successful_frames.append(dataframe)
        except Exception as extraction_error:
            failed_images.append((str(image_path), str(extraction_error)))
            print(f"  ❌ {image_path}: {extraction_error}", file=sys.stderr)

    if failed_images:
        print(
            f"\n⚠️  {len(failed_images)} image(s) failed to process:",
            file=sys.stderr,
        )
        for path, error_message in failed_images:
            print(f"     {path}: {error_message}", file=sys.stderr)

    if not successful_frames:
        raise RuntimeError("No images were successfully processed.")

    return pd.concat(successful_frames, ignore_index=True)


def batch_process_images(
    input_dir: Union[str, Path],
    start_num: int = 1,
    output_dir: Optional[Union[str, Path]] = None,
    pdf: bool = False,
    config: Optional[dict] = None,
    api_key: Optional[str] = None,
    **kwargs: Any,
) -> list[BatchResult]:
    """
    Process a folder of challan images, generating one invoice per image.

    Each image is extracted into a DataFrame, saved as Excel data, and then
    passed through the invoice generator pipeline. Invoice numbers auto-increment
    starting from start_num.

    This mirrors the batch_process() function in generate_invoices.py but
    works with images instead of Excel files.

    Args:
        input_dir: Directory containing challan images.
        start_num: Invoice number for the first image (auto-increments).
        output_dir: Where to save outputs. Defaults to input_dir.
        pdf: Whether to also generate PDF invoices.
        config: Business configuration dict for invoice generation.
        api_key: OpenRouter API key.
        **kwargs: Passed through to extract_from_image (force_llm, force_ocr).

    Returns:
        List of BatchResult objects, one per image processed.
    """
    from .generator import generate_invoice, generate_pdf

    input_directory = Path(input_dir)
    output_directory = Path(output_dir) if output_dir else input_directory
    output_directory.mkdir(parents=True, exist_ok=True)

    if api_key is None:
        api_key = os.environ.get("OPENROUTER_API_KEY")

    # Find all image files, sorted alphabetically
    image_files = sorted(
        file_path
        for file_path in input_directory.iterdir()
        if file_path.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
        and not file_path.name.startswith(".")
    )

    if not image_files:
        print(f"No images found in {input_directory}", file=sys.stderr)
        return []

    results = []

    for file_index, image_path in enumerate(image_files):
        invoice_number = start_num + file_index
        result = BatchResult(
            input_path=str(image_path),
            invoice_number=invoice_number,
        )

        try:
            # Step 1: Extract data from image
            dataframe = extract_from_image(
                image_path, api_key=api_key, **kwargs,
            )
            result.record_count = len(dataframe)

            # Step 2: Save extracted data
            extracted_path = output_directory / f"Data_{invoice_number}.xlsx"
            dataframe.to_excel(extracted_path, index=False, engine="openpyxl")
            result.extracted_data_path = str(extracted_path)

            # Step 3: Generate invoice
            workbook = generate_invoice(extracted_path, invoice_number, config)
            excel_output = output_directory / f"Invoice_{invoice_number}.xlsx"
            workbook.save(str(excel_output))
            result.excel_path = str(excel_output)
            print(
                f"  ✅ [{invoice_number}] {image_path.name} "
                f"→ {excel_output.name} ({result.record_count} records)",
            )

            # Step 4: Generate PDF (if requested)
            if pdf:
                pdf_output = output_directory / f"Invoice_{invoice_number}.pdf"
                generate_pdf(extracted_path, invoice_number, pdf_output, config)
                result.pdf_path = str(pdf_output)
                print(f"        📄 {pdf_output.name}")

        except Exception as processing_error:
            result.error = str(processing_error)
            print(
                f"  ❌ [{invoice_number}] {image_path.name}: {processing_error}",
                file=sys.stderr,
            )

        results.append(result)

    return results


def images_to_invoice(
    images: list[Union[str, Path]],
    invoice_number: int,
    output_excel: Optional[Union[str, Path]] = None,
    output_pdf: Optional[Union[str, Path]] = None,
    api_key: Optional[str] = None,
    config: Optional[dict] = None,
    **kwargs: Any,
) -> tuple[Any, Optional[bytes]]:
    """
    Full end-to-end pipeline: challan images → combined invoice.

    Unlike batch_process_images (which generates one invoice per image),
    this function combines all images into a single invoice — useful for
    multi-page challans or consolidating a month's deliveries.

    Args:
        images: List of image file paths to combine.
        invoice_number: Invoice number to assign.
        output_excel: Path to save the Excel invoice. If None, not saved.
        output_pdf: Path to save the PDF invoice. If None, not generated.
        api_key: OpenRouter API key.
        config: Business configuration dict.
        **kwargs: Passed through to extract_from_image.

    Returns:
        Tuple of (openpyxl.Workbook, pdf_bytes_or_None).
    """
    from .generator import generate_invoice, generate_pdf as gen_pdf

    # Extract data from all images
    combined_dataframe = extract_from_images(images, api_key=api_key, **kwargs)

    # Save as intermediate Excel for the invoice generator pipeline
    temp_excel = Path(f"_extracted_{invoice_number}.xlsx")
    combined_dataframe.to_excel(temp_excel, index=False, engine="openpyxl")

    try:
        workbook = generate_invoice(temp_excel, invoice_number, config)

        if output_excel:
            workbook.save(str(output_excel))

        pdf_bytes = None
        if output_pdf:
            gen_pdf(temp_excel, invoice_number, output_pdf, config)
            pdf_bytes = Path(output_pdf).read_bytes()

        return workbook, pdf_bytes

    finally:
        if temp_excel.exists():
            temp_excel.unlink()


# ─── CLI ─────────────────────────────────────────────────────────────────────

def main() -> None:
    """Command-line interface for image-to-invoice processing."""
    parser = argparse.ArgumentParser(
        description="Extract delivery data from challan images",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  %(prog)s --image challan.jpg --output data.xlsx     # Extract data only\n"
            "  %(prog)s --image challan.jpg -i 178 --pdf           # Single invoice\n"
            "  %(prog)s --batch ./photos/ --start 178 --pdf        # Batch: 1 invoice per image\n"
            "  %(prog)s --combine img1.jpg img2.jpg -i 178         # Combine into 1 invoice\n"
            "  %(prog)s --image challan.jpg --force-ocr            # Offline (no API)\n"
        ),
    )

    # Input modes (mutually exclusive in practice)
    parser.add_argument("--image", type=str,
                        help="Single image file to process")
    parser.add_argument("--batch", type=str, metavar="DIR",
                        help="Batch process: 1 invoice per image in directory")
    parser.add_argument("--combine", nargs="+", type=str, metavar="IMG",
                        help="Combine multiple images into 1 invoice")

    # Output options
    parser.add_argument("--output", type=str, default="extracted_data.xlsx",
                        help="Output Excel file (default: extracted_data.xlsx)")
    parser.add_argument("--output-dir", type=str, metavar="DIR",
                        help="Output directory for batch mode")
    parser.add_argument("-i", "--invoice", type=int,
                        help="Invoice number (triggers invoice generation)")
    parser.add_argument("--start", type=int, default=1,
                        help="Starting invoice number for batch mode (default: 1)")
    parser.add_argument("--pdf", action="store_true",
                        help="Also generate PDF output")

    # Processing options
    parser.add_argument("--force-llm", action="store_true",
                        help="Only use LLM (fail if API unavailable)")
    parser.add_argument("--force-ocr", action="store_true",
                        help="Only use OCR (skip LLM entirely)")
    parser.add_argument("--api-key", type=str,
                        help="OpenRouter API key (default: OPENROUTER_API_KEY env var)")
    parser.add_argument("--config", type=str,
                        help="Path to config YAML file")
    parser.add_argument("--version", action="version",
                        version=f"%(prog)s {__version__}")

    args = parser.parse_args()

    api_key = args.api_key or os.environ.get("OPENROUTER_API_KEY")
    force_ocr = args.force_ocr

    if not api_key and not force_ocr:
        print(
            "⚠️  No API key found. Set OPENROUTER_API_KEY in .env or use --api-key.\n"
            "    Falling back to OCR-only mode (less accurate).\n",
            file=sys.stderr,
        )
        force_ocr = True

    # Load config if specified
    config = None
    if args.config:
        from .generator import load_config
        config = load_config(args.config)

    try:
        # ── Batch mode: 1 invoice per image ──
        if args.batch:
            batch_dir = Path(args.batch)
            if not batch_dir.is_dir():
                print(f"Error: {args.batch} is not a directory.", file=sys.stderr)
                sys.exit(1)

            print(f"📦 Batch processing: {batch_dir}/ (starting at #{args.start})\n")

            results = batch_process_images(
                batch_dir,
                start_num=args.start,
                output_dir=args.output_dir,
                pdf=args.pdf,
                config=config,
                api_key=api_key,
                force_llm=args.force_llm,
                force_ocr=force_ocr,
            )

            succeeded = sum(1 for r in results if r.error is None)
            failed = len(results) - succeeded
            total_records = sum(r.record_count for r in results)
            print(
                f"\n✅ {succeeded} succeeded, ❌ {failed} failed "
                f"({total_records} total records)"
            )
            sys.exit(1 if failed > 0 else 0)

        # ── Combine mode: multiple images → 1 invoice ──
        if args.combine:
            invoice_number = args.invoice
            if invoice_number is None:
                try:
                    invoice_number = int(input("Enter invoice number: "))
                except (ValueError, EOFError):
                    print("Error: Invoice number required for combine mode.",
                          file=sys.stderr)
                    sys.exit(1)

            print(f"📸 Combining {len(args.combine)} images into invoice #{invoice_number}\n")

            excel_out = args.output.replace("extracted_data", f"Invoice_{invoice_number}")
            pdf_out = Path(excel_out).with_suffix(".pdf") if args.pdf else None

            images_to_invoice(
                args.combine, invoice_number,
                output_excel=excel_out,
                output_pdf=pdf_out,
                api_key=api_key,
                config=config,
                force_llm=args.force_llm,
                force_ocr=force_ocr,
            )

            print(f"\n✅ Invoice Excel: {excel_out}")
            if pdf_out:
                print(f"📄 Invoice PDF:   {pdf_out}")
            return

        # ── Single image mode ──
        if args.image:
            dataframe = extract_from_image(
                args.image, api_key=api_key,
                force_llm=args.force_llm, force_ocr=force_ocr,
            )

            print(f"\n📊 Extracted {len(dataframe)} records")

            # Save extracted data
            dataframe.to_excel(args.output, index=False, engine="openpyxl")
            print(f"💾 Saved: {args.output}")

            # Generate invoice if requested
            if args.invoice:
                from .generator import generate_invoice, generate_pdf

                excel_out = f"Invoice_{args.invoice}.xlsx"
                workbook = generate_invoice(args.output, args.invoice, config)
                workbook.save(excel_out)
                print(f"✅ Invoice Excel: {excel_out}")

                if args.pdf:
                    pdf_out = f"Invoice_{args.invoice}.pdf"
                    generate_pdf(args.output, args.invoice, pdf_out, config)
                    print(f"📄 Invoice PDF:   {pdf_out}")

            return

        # No input specified
        parser.print_help()
        sys.exit(1)

    except Exception as error:
        print(f"\n❌ Error: {error}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
