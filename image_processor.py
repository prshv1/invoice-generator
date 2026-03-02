#!/usr/bin/env python3
"""
Image Processor — Extract delivery data from challan/bill images.

Converts photos of handwritten or printed delivery challans into structured
Excel data that feeds into the Invoice Generator pipeline.

Architecture:
    1. LLM Vision (primary) — sends image to OpenRouter free model for extraction
    2. OCR + LLM Text (fallback) — EasyOCR extracts text, LLM structures it
    3. OCR + Heuristics (offline) — pure-offline extraction via spatial parsing

CLI Usage:
    python3 image_processor.py --image challan.jpg -i 178
    python3 image_processor.py --images ./photos/ --start 178 --pdf
    python3 image_processor.py --image challan.jpg --output raw_data.xlsx

Module Usage:
    from image_processor import extract_from_image, images_to_invoice
    df = extract_from_image("challan.jpg")
    wb, pdf_bytes = images_to_invoice(["img1.jpg"], invoice_number=178)
"""

__version__ = "0.2.0"

import argparse
import base64
import io
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Optional, Union

import pandas as pd
from dotenv import load_dotenv

# Load .env file for API keys
load_dotenv()

# ─── Constants ───────────────────────────────────────────────────────────────

OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"

# Model cascade: try each in order until one works.
# Verified available via OpenRouter /api/v1/models endpoint.
LLM_MODELS = [
    "google/gemma-3-27b-it:free",                   # Best quality free vision model
    "mistralai/mistral-small-3.1-24b-instruct:free", # Strong backup
    "google/gemma-3-12b-it:free",                    # Lighter, faster
    "nvidia/nemotron-nano-12b-v2-vl:free",           # Supports video too
    "google/gemma-3-4b-it:free",                     # Smallest, last resort
]

MAX_IMAGE_WIDTH = 2000         # Resize images wider than this (px)
JPEG_QUALITY = 85              # Compression quality for API uploads
API_TIMEOUT_SECONDS = 60       # Max wait for LLM response
MAX_API_RETRIES = 3            # Retries per model before moving to next
OCR_CONFIDENCE_THRESHOLD = 0.80
MAX_DESKEW_ANGLE = 15.0        # Don't deskew more than this (degrees)

SUPPORTED_IMAGE_EXTENSIONS = frozenset({
    ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp",
})

# The expected output schema for extracted delivery records
OUTPUT_COLUMNS = [
    "Date", "Challan No.", "Vehicle No.", "Site",
    "Material", "Quantity", "Rate", "Per",
]

# The prompt sent to the LLM for image extraction
EXTRACTION_PROMPT = """You are an expert at reading Indian delivery challans, bills, and handwritten records.

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

# Simpler prompt for when we send OCR text (not image) to LLM
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


# ─── Image Preprocessing ────────────────────────────────────────────────────

def preprocess_image(image_path: Union[str, Path]) -> bytes:
    """
    Load, preprocess, and return image as JPEG bytes optimized for OCR/LLM.

    Preprocessing steps:
    1. Load image (handling various formats)
    2. Auto-rotate via EXIF orientation
    3. Resize to max width (saves API bandwidth and OCR time)
    4. Convert to grayscale
    5. Apply adaptive thresholding for better contrast
    6. Denoise

    Returns:
        JPEG-encoded bytes of the preprocessed image.
    """
    import cv2
    import numpy as np
    from PIL import Image, ExifTags

    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(f"Image not found: {path}")

    if path.suffix.lower() not in SUPPORTED_IMAGE_EXTENSIONS:
        raise ValueError(
            f"Unsupported image format: {path.suffix}. "
            f"Supported: {', '.join(sorted(SUPPORTED_IMAGE_EXTENSIONS))}"
        )

    # Load with Pillow first (handles EXIF rotation)
    pil_image = Image.open(path)

    # Auto-rotate based on EXIF orientation tag
    try:
        exif = pil_image._getexif()
        if exif:
            orientation_key = next(
                (k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None
            )
            if orientation_key and orientation_key in exif:
                orientation = exif[orientation_key]
                rotation_map = {3: 180, 6: 270, 8: 90}
                if orientation in rotation_map:
                    pil_image = pil_image.rotate(rotation_map[orientation], expand=True)
    except (AttributeError, StopIteration):
        pass  # No EXIF data, skip rotation

    # Convert to RGB if necessary (handles RGBA, palette images)
    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")

    # Resize if too large
    width, height = pil_image.size
    if width > MAX_IMAGE_WIDTH:
        scale = MAX_IMAGE_WIDTH / width
        new_size = (MAX_IMAGE_WIDTH, int(height * scale))
        pil_image = pil_image.resize(new_size, Image.LANCZOS)

    # Convert to OpenCV format for preprocessing
    cv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)

    # Grayscale
    gray = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)

    # Denoise (removes speckles from paper texture)
    denoised = cv2.fastNlMeansDenoising(gray, h=10)

    # Adaptive threshold (handles uneven lighting)
    # We keep the original too — some images work better without thresholding
    thresholded = cv2.adaptiveThreshold(
        denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 11, 2,
    )

    # Encode as JPEG bytes
    _, jpeg_bytes = cv2.imencode(
        ".jpg", thresholded,
        [cv2.IMWRITE_JPEG_QUALITY, JPEG_QUALITY],
    )
    return jpeg_bytes.tobytes()


def load_image_as_base64(image_path: Union[str, Path]) -> str:
    """
    Load an image file and return it as a base64-encoded JPEG string.

    Applies preprocessing (resize, denoise) to optimize for LLM consumption.
    """
    jpeg_bytes = preprocess_image(image_path)
    return base64.b64encode(jpeg_bytes).decode("utf-8")


def load_raw_image_as_base64(image_path: Union[str, Path]) -> str:
    """
    Load an image file as base64 WITHOUT heavy preprocessing.

    Used for LLM vision API where the model handles interpretation itself.
    We only resize (for bandwidth) and fix EXIF rotation.
    """
    from PIL import Image, ExifTags

    path = Path(image_path)
    pil_image = Image.open(path)

    # EXIF rotation
    try:
        exif = pil_image._getexif()
        if exif:
            orientation_key = next(
                (k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None
            )
            if orientation_key and orientation_key in exif:
                orientation = exif[orientation_key]
                rotation_map = {3: 180, 6: 270, 8: 90}
                if orientation in rotation_map:
                    pil_image = pil_image.rotate(rotation_map[orientation], expand=True)
    except (AttributeError, StopIteration):
        pass

    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")

    # Resize for bandwidth
    width, height = pil_image.size
    if width > MAX_IMAGE_WIDTH:
        scale = MAX_IMAGE_WIDTH / width
        pil_image = pil_image.resize((MAX_IMAGE_WIDTH, int(height * scale)), Image.LANCZOS)

    # Encode as JPEG
    buffer = io.BytesIO()
    pil_image.save(buffer, format="JPEG", quality=JPEG_QUALITY)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


# ─── JSON Repair ─────────────────────────────────────────────────────────────

def repair_json(raw_text: str) -> list[dict]:
    """
    Attempt to parse JSON from potentially messy LLM output.

    Handles common issues:
    - Markdown code fences (```json ... ```)
    - Leading/trailing commentary
    - Trailing commas
    - Single quotes instead of double quotes

    Returns:
        Parsed list of dicts, or raises ValueError if unfixable.
    """
    text = raw_text.strip()

    # Strip markdown code fences
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s*```$', '', text, flags=re.MULTILINE)
    text = text.strip()

    # Try direct parse first
    try:
        result = json.loads(text)
        if isinstance(result, list):
            return result
        if isinstance(result, dict):
            return [result]
    except json.JSONDecodeError:
        pass

    # Try to extract JSON array from surrounding text
    array_match = re.search(r'\[[\s\S]*\]', text)
    if array_match:
        try:
            result = json.loads(array_match.group())
            if isinstance(result, list):
                return result
        except json.JSONDecodeError:
            pass

    # Fix trailing commas: ,] → ]  and ,} → }
    cleaned = re.sub(r',\s*([}\]])', r'\1', text)
    try:
        result = json.loads(cleaned)
        if isinstance(result, list):
            return result
    except json.JSONDecodeError:
        pass

    # Last resort: try extracting individual objects
    objects = re.findall(r'\{[^{}]+\}', text)
    if objects:
        parsed = []
        for obj_str in objects:
            try:
                parsed.append(json.loads(obj_str))
            except json.JSONDecodeError:
                continue
        if parsed:
            return parsed

    raise ValueError(f"Could not parse JSON from LLM response:\n{raw_text[:500]}")


# ─── LLM Extraction (Primary) ───────────────────────────────────────────────

def _call_openrouter(
    messages: list[dict],
    model: str,
    api_key: str,
) -> str:
    """
    Make a single API call to OpenRouter.

    Returns:
        The assistant's response text.

    Raises:
        requests.HTTPError: On API errors (4xx, 5xx).
        requests.Timeout: On timeout.
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
    data = response.json()

    if "choices" not in data or len(data["choices"]) == 0:
        raise ValueError(f"Empty response from model {model}: {data}")

    return data["choices"][0]["message"]["content"]


def extract_with_llm_vision(
    image_path: Union[str, Path],
    api_key: str,
    models: Optional[list[str]] = None,
) -> pd.DataFrame:
    """
    Extract delivery records from an image using LLM vision API.

    Sends the image directly to a vision-capable model on OpenRouter.
    Cascades through multiple free models if one fails.

    Args:
        image_path: Path to the challan image.
        api_key: OpenRouter API key.
        models: List of model IDs to try (in priority order).

    Returns:
        DataFrame with columns matching OUTPUT_COLUMNS.

    Raises:
        RuntimeError: If all models fail.
    """
    import requests

    if models is None:
        models = LLM_MODELS.copy()

    # Encode image — use raw (minimal preprocessing) for LLM
    # LLMs understand images better without heavy binarization
    image_base64 = load_raw_image_as_base64(image_path)

    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": EXTRACTION_PROMPT},
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/jpeg;base64,{image_base64}",
                    },
                },
            ],
        }
    ]

    last_error = None

    for model in models:
        for attempt in range(MAX_API_RETRIES):
            try:
                print(f"  🤖 Trying {model} (attempt {attempt + 1})...", file=sys.stderr)
                response_text = _call_openrouter(messages, model, api_key)
                records = repair_json(response_text)
                dataframe = _records_to_dataframe(records)
                print(f"  ✅ Extracted {len(dataframe)} records via {model}", file=sys.stderr)
                return dataframe

            except requests.exceptions.HTTPError as error:
                status = error.response.status_code if error.response else "?"
                if status == 429:
                    # Rate limited — wait and retry
                    wait_time = 2 ** attempt
                    print(f"  ⏳ Rate limited, waiting {wait_time}s...", file=sys.stderr)
                    time.sleep(wait_time)
                    continue
                elif status in (502, 503):
                    # Model temporarily down — try next model
                    print(f"  ⚠️  {model} unavailable ({status}), trying next...", file=sys.stderr)
                    last_error = error
                    break
                else:
                    last_error = error
                    break

            except requests.exceptions.Timeout:
                print(f"  ⏰ Timeout on {model}, retrying...", file=sys.stderr)
                last_error = TimeoutError(f"Timeout calling {model}")
                continue

            except (ValueError, json.JSONDecodeError) as error:
                # JSON parsing failed — retry with explicit instruction
                print(f"  ⚠️  JSON parse error: {error}", file=sys.stderr)
                last_error = error
                if attempt < MAX_API_RETRIES - 1:
                    # Retry with stricter prompt
                    messages[0]["content"][0]["text"] = (
                        EXTRACTION_PROMPT + "\n\nIMPORTANT: Return ONLY valid JSON. "
                        "No markdown, no explanation, just the array."
                    )
                continue

            except Exception as error:
                last_error = error
                break

    raise RuntimeError(
        f"All LLM models failed. Last error: {last_error}\n"
        f"Models tried: {', '.join(models)}"
    )


def extract_with_llm_text(
    ocr_text: str,
    api_key: str,
    models: Optional[list[str]] = None,
) -> pd.DataFrame:
    """
    Send raw OCR text to an LLM for structuring into table format.

    This is the middle-ground fallback: OCR handles character recognition,
    LLM handles understanding the structure. Cheaper than sending images
    and works with non-vision models too.
    """
    import requests

    if models is None:
        models = LLM_MODELS.copy()

    messages = [
        {
            "role": "user",
            "content": TEXT_STRUCTURING_PROMPT + ocr_text,
        }
    ]

    last_error = None

    for model in models:
        try:
            response_text = _call_openrouter(messages, model, api_key)
            records = repair_json(response_text)
            return _records_to_dataframe(records)
        except Exception as error:
            last_error = error
            continue

    raise RuntimeError(f"LLM text structuring failed. Last error: {last_error}")


# ─── OCR Extraction (Fallback) ───────────────────────────────────────────────

def _get_ocr_reader():
    """Lazy-load EasyOCR reader (first load downloads ~500MB model)."""
    import easyocr
    print("  📦 Loading OCR model (first run downloads ~500MB)...", file=sys.stderr)
    reader = easyocr.Reader(["en", "hi"], verbose=False)
    return reader


def extract_with_ocr(image_path: Union[str, Path]) -> tuple[str, float]:
    """
    Extract text from an image using EasyOCR.

    Returns:
        Tuple of (extracted_text, median_confidence).
        The text is arranged line-by-line based on spatial position.
    """
    import numpy as np

    reader = _get_ocr_reader()
    preprocessed_bytes = preprocess_image(image_path)

    # Convert bytes to numpy array for EasyOCR
    nparr = np.frombuffer(preprocessed_bytes, np.uint8)
    import cv2
    image = cv2.imdecode(nparr, cv2.IMREAD_GRAYSCALE)

    # Run OCR
    results = reader.readtext(image, detail=1)
    # results = [(bbox, text, confidence), ...]

    if not results:
        return "", 0.0

    # Compute median confidence
    confidences = [conf for _, _, conf in results]
    median_confidence = float(np.median(confidences))

    # Group detections into lines based on Y-coordinate
    lines = _group_into_lines(results)

    # Build text output
    text_lines = []
    for line in lines:
        line_text = " | ".join(word[1] for word in line)
        text_lines.append(line_text)

    full_text = "\n".join(text_lines)
    return full_text, median_confidence


def _group_into_lines(
    detections: list[tuple],
    y_tolerance: int = 15,
) -> list[list[tuple]]:
    """
    Group OCR detections into lines based on vertical position.

    Detections on similar Y-coordinates (within y_tolerance pixels)
    are grouped into the same line, then sorted left-to-right.
    """
    if not detections:
        return []

    # Get center Y of each detection
    items = []
    for bbox, text, conf in detections:
        # bbox is 4 corner points: [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
        center_y = sum(pt[1] for pt in bbox) / 4
        center_x = sum(pt[0] for pt in bbox) / 4
        items.append((center_x, center_y, bbox, text, conf))

    # Sort by Y coordinate
    items.sort(key=lambda item: item[1])

    # Group into lines
    lines = []
    current_line = [items[0]]

    for item in items[1:]:
        if abs(item[1] - current_line[-1][1]) <= y_tolerance:
            current_line.append(item)
        else:
            lines.append(current_line)
            current_line = [item]
    lines.append(current_line)

    # Sort each line left-to-right by X
    for line in lines:
        line.sort(key=lambda item: item[0])

    # Convert back to (bbox, text, conf) format
    return [
        [(item[2], item[3], item[4]) for item in line]
        for line in lines
    ]


def ocr_text_to_dataframe(ocr_text: str) -> pd.DataFrame:
    """
    Attempt to parse structured OCR text into a DataFrame using heuristics.

    This is the pure-offline fallback when no LLM API is available.
    It uses regex patterns to identify field types:
    - Dates: dd/mm/yyyy or dd-mm-yyyy patterns
    - Numbers: sequences of digits (challan, vehicle, qty, rate)
    - Text: site names, material types
    """
    lines = [line.strip() for line in ocr_text.strip().split("\n") if line.strip()]

    # Skip header-like lines
    skip_words = {"date", "challan", "vehicle", "site", "material", "quantity",
                  "rate", "per", "amount", "total", "sum", "grand"}

    date_pattern = re.compile(r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}')
    number_pattern = re.compile(r'^\d+\.?\d*$')

    records = []

    for line in lines:
        # Split by pipe (our OCR output format) or tabs
        parts = [p.strip() for p in re.split(r'[|\t]', line) if p.strip()]

        # Skip headers and totals
        if any(word in " ".join(parts).lower() for word in skip_words):
            continue

        if len(parts) < 4:
            continue  # Too few fields to be a data row

        # Try to identify a date in this line
        date_val = None
        for part in parts:
            if date_pattern.search(part):
                date_val = date_pattern.search(part).group()
                break

        if date_val is None:
            continue  # Can't identify row without a date

        # Collect numeric values
        numeric_vals = []
        text_vals = []
        for part in parts:
            if part == date_val:
                continue
            cleaned = part.replace(",", "").strip()
            if number_pattern.match(cleaned):
                numeric_vals.append(float(cleaned))
            else:
                text_vals.append(part)

        # Heuristic mapping:
        # numeric_vals: [challan, vehicle, qty, rate] (roughly in order)
        # text_vals: [site, material, per]
        record = {
            "Date": date_val,
            "Challan No.": numeric_vals[0] if len(numeric_vals) > 0 else None,
            "Vehicle No.": numeric_vals[1] if len(numeric_vals) > 1 else None,
            "Site": text_vals[0] if len(text_vals) > 0 else "",
            "Material": text_vals[1] if len(text_vals) > 1 else "",
            "Quantity": numeric_vals[-2] if len(numeric_vals) > 3 else (
                numeric_vals[2] if len(numeric_vals) > 2 else 0),
            "Rate": numeric_vals[-1] if len(numeric_vals) > 2 else 0,
            "Per": text_vals[2] if len(text_vals) > 2 else "Tonne",
        }

        records.append(record)

    if not records:
        raise ValueError("Could not extract any records from OCR text using heuristics.")

    return pd.DataFrame(records, columns=OUTPUT_COLUMNS)


# ─── Data Validation ─────────────────────────────────────────────────────────

def _records_to_dataframe(records: list[dict]) -> pd.DataFrame:
    """
    Convert a list of parsed records into a validated DataFrame.

    Applies:
    - Column name normalization
    - Date format enforcement (dd/mm/yyyy)
    - Numeric coercion for Quantity and Rate
    - Duplicate row detection
    - Totals row filtering
    """
    if not records:
        raise ValueError("No records to convert.")

    dataframe = pd.DataFrame(records)

    # Normalize column names (case-insensitive matching)
    column_map = {}
    for expected_col in OUTPUT_COLUMNS:
        for actual_col in dataframe.columns:
            if actual_col.lower().replace("_", " ").strip() == expected_col.lower():
                column_map[actual_col] = expected_col
                break
    dataframe = dataframe.rename(columns=column_map)

    # Ensure all required columns exist
    for col in OUTPUT_COLUMNS:
        if col not in dataframe.columns:
            dataframe[col] = None if col in ("Challan No.", "Vehicle No.") else ""

    # Keep only expected columns, in order
    dataframe = dataframe[[col for col in OUTPUT_COLUMNS if col in dataframe.columns]]

    # Fill missing Per with "Tonne"
    if "Per" in dataframe.columns:
        dataframe["Per"] = dataframe["Per"].fillna("Tonne").replace("", "Tonne")

    # Coerce numeric columns
    for col in ["Quantity", "Rate"]:
        if col in dataframe.columns:
            dataframe[col] = pd.to_numeric(dataframe[col], errors="coerce")

    # Coerce Challan No. and Vehicle No. to numeric (allow NaN)
    for col in ["Challan No.", "Vehicle No."]:
        if col in dataframe.columns:
            dataframe[col] = pd.to_numeric(dataframe[col], errors="coerce")

    # Filter out totals/summary rows
    if "Site" in dataframe.columns:
        total_keywords = ["total", "sum", "grand", "subtotal", "jodh", "कुल"]
        mask = dataframe.apply(
            lambda row: not any(
                keyword in str(val).lower()
                for val in row.values
                for keyword in total_keywords
            ),
            axis=1,
        )
        dataframe = dataframe[mask]

    # Drop rows where both Quantity and Rate are NaN/0
    dataframe = dataframe.dropna(subset=["Quantity", "Rate"], how="all")

    # Deduplicate by Challan No. + Date + Material (if Challan No. exists)
    if dataframe["Challan No."].notna().any():
        before_count = len(dataframe)
        dataframe = dataframe.drop_duplicates(
            subset=["Date", "Challan No.", "Material"], keep="first"
        )
        dupes_removed = before_count - len(dataframe)
        if dupes_removed > 0:
            print(f"  ⚠️  Removed {dupes_removed} duplicate row(s)", file=sys.stderr)

    dataframe = dataframe.reset_index(drop=True)

    if dataframe.empty:
        raise ValueError("All records were filtered out during validation.")

    return dataframe


def validate_extraction(dataframe: pd.DataFrame) -> list[str]:
    """
    Run post-extraction validation checks and return a list of warnings.

    Checks:
    - Missing critical fields
    - Suspicious values (zero qty/rate, very large amounts)
    - Date plausibility
    """
    warnings = []

    # Check for missing quantities or rates
    null_qty = dataframe["Quantity"].isna().sum()
    null_rate = dataframe["Rate"].isna().sum()
    if null_qty > 0:
        warnings.append(f"{null_qty} row(s) have missing Quantity")
    if null_rate > 0:
        warnings.append(f"{null_rate} row(s) have missing Rate")

    # Check for zero values
    zero_qty = (dataframe["Quantity"] == 0).sum()
    zero_rate = (dataframe["Rate"] == 0).sum()
    if zero_qty > 0:
        warnings.append(f"{zero_qty} row(s) have Quantity = 0")
    if zero_rate > 0:
        warnings.append(f"{zero_rate} row(s) have Rate = 0")

    # Check for suspiciously large amounts
    amounts = dataframe["Quantity"].fillna(0) * dataframe["Rate"].fillna(0)
    huge = amounts[amounts > 1_000_000]
    if len(huge) > 0:
        warnings.append(
            f"{len(huge)} row(s) have Amount > ₹10,00,000 — verify these are correct"
        )

    # Check for empty Site or Material
    empty_site = (dataframe["Site"].astype(str).str.strip() == "").sum()
    empty_mat = (dataframe["Material"].astype(str).str.strip() == "").sum()
    if empty_site > 0:
        warnings.append(f"{empty_site} row(s) have empty Site")
    if empty_mat > 0:
        warnings.append(f"{empty_mat} row(s) have empty Material")

    return warnings


# ─── Public API ──────────────────────────────────────────────────────────────

def extract_from_image(
    image_path: Union[str, Path],
    api_key: Optional[str] = None,
    force_llm: bool = False,
    force_ocr: bool = False,
) -> pd.DataFrame:
    """
    Extract delivery records from a single challan image.

    Pipeline:
    1. Try LLM vision extraction (best accuracy)
    2. If LLM fails, try OCR + LLM text structuring (fallback)
    3. If all LLM calls fail, try pure OCR with heuristics (offline)

    Args:
        image_path: Path to the challan image file.
        api_key: OpenRouter API key. If None, reads from OPENROUTER_API_KEY env var.
        force_llm: If True, only use LLM (skip OCR fallback).
        force_ocr: If True, only use OCR (skip LLM entirely).

    Returns:
        DataFrame with columns: Date, Challan No., Vehicle No., Site,
        Material, Quantity, Rate, Per.
    """
    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(f"Image not found: {path}")

    if api_key is None:
        api_key = os.environ.get("OPENROUTER_API_KEY")

    print(f"\n📸 Processing: {path.name}", file=sys.stderr)

    # ── Strategy 1: LLM Vision (primary) ──
    if not force_ocr and api_key:
        try:
            dataframe = extract_with_llm_vision(path, api_key)
            warnings = validate_extraction(dataframe)
            for warning in warnings:
                print(f"  ⚠️  {warning}", file=sys.stderr)
            return dataframe
        except Exception as error:
            print(f"  ⚠️  LLM vision failed: {error}", file=sys.stderr)
            if force_llm:
                raise

    # ── Strategy 2: OCR + LLM text structuring (fallback) ──
    if not force_ocr and api_key:
        try:
            print("  🔍 Falling back to OCR + LLM text...", file=sys.stderr)
            ocr_text, confidence = extract_with_ocr(path)
            print(f"  📊 OCR confidence: {confidence:.0%}", file=sys.stderr)
            if ocr_text.strip():
                dataframe = extract_with_llm_text(ocr_text, api_key)
                warnings = validate_extraction(dataframe)
                for warning in warnings:
                    print(f"  ⚠️  {warning}", file=sys.stderr)
                return dataframe
        except Exception as error:
            print(f"  ⚠️  OCR + LLM text failed: {error}", file=sys.stderr)

    # ── Strategy 3: Pure OCR with heuristics (offline) ──
    if api_key and force_llm:
        raise RuntimeError("LLM extraction failed and force_llm=True prevents OCR fallback.")

    print("  🔍 Falling back to pure OCR (offline mode)...", file=sys.stderr)
    ocr_text, confidence = extract_with_ocr(path)
    print(f"  📊 OCR confidence: {confidence:.0%}", file=sys.stderr)

    if confidence < OCR_CONFIDENCE_THRESHOLD:
        print(
            f"  ⚠️  Low OCR confidence ({confidence:.0%} < {OCR_CONFIDENCE_THRESHOLD:.0%}). "
            f"Results may be unreliable.",
            file=sys.stderr,
        )

    dataframe = ocr_text_to_dataframe(ocr_text)
    warnings = validate_extraction(dataframe)
    for warning in warnings:
        print(f"  ⚠️  {warning}", file=sys.stderr)

    return dataframe


def extract_from_images(
    image_paths: list[Union[str, Path]],
    api_key: Optional[str] = None,
    **kwargs,
) -> pd.DataFrame:
    """
    Extract delivery records from multiple challan images and combine them.

    Each image is processed independently, then all results are concatenated
    into a single DataFrame.

    Returns:
        Combined DataFrame with all records from all images.
    """
    all_frames = []
    errors = []

    for image_path in image_paths:
        try:
            df = extract_from_image(image_path, api_key=api_key, **kwargs)
            all_frames.append(df)
        except Exception as error:
            errors.append((str(image_path), str(error)))
            print(f"  ❌ {image_path}: {error}", file=sys.stderr)

    if errors:
        print(f"\n⚠️  {len(errors)} image(s) failed to process:", file=sys.stderr)
        for path, err in errors:
            print(f"     {path}: {err}", file=sys.stderr)

    if not all_frames:
        raise RuntimeError("No images were successfully processed.")

    combined = pd.concat(all_frames, ignore_index=True)
    return combined


def images_to_invoice(
    images: list[Union[str, Path]],
    invoice_number: int,
    output_excel: Optional[Union[str, Path]] = None,
    output_pdf: Optional[Union[str, Path]] = None,
    api_key: Optional[str] = None,
    config: Optional[dict] = None,
    **kwargs,
) -> tuple:
    """
    Full pipeline: images → extract data → generate invoice.

    Args:
        images: List of image file paths.
        invoice_number: Invoice number to assign.
        output_excel: Path to save Excel output.
        output_pdf: Path to save PDF output.
        api_key: OpenRouter API key.
        config: Business config dict for invoice generation.

    Returns:
        Tuple of (Workbook, pdf_bytes_or_None).
    """
    from generate_invoices import generate_invoice, generate_pdf as gen_pdf

    # Step 1: Extract data from images
    dataframe = extract_from_images(images, api_key=api_key, **kwargs)

    # Step 2: Save intermediate data as Excel
    temp_excel = Path("_extracted_data.xlsx")
    dataframe.to_excel(temp_excel, index=False, engine="openpyxl")

    try:
        # Step 3: Generate invoice from extracted data
        workbook = generate_invoice(temp_excel, invoice_number, config)

        if output_excel:
            workbook.save(str(output_excel))

        pdf_bytes = None
        if output_pdf:
            gen_pdf(temp_excel, invoice_number, output_pdf, config)
            pdf_bytes = Path(output_pdf).read_bytes()

        return workbook, pdf_bytes

    finally:
        # Clean up temp file
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
            "  %(prog)s --image challan.jpg --output raw_data.xlsx\n"
            "  %(prog)s --image challan.jpg -i 178\n"
            "  %(prog)s --images ./photos/ --start 178 --pdf\n"
            "  %(prog)s --image challan.jpg --force-ocr\n"
        ),
    )
    parser.add_argument("--image", type=str, help="Single image file to process")
    parser.add_argument("--images", type=str, metavar="DIR",
                        help="Directory of images to process")
    parser.add_argument("--output", type=str, default="extracted_data.xlsx",
                        help="Output Excel file for extracted data (default: extracted_data.xlsx)")
    parser.add_argument("-i", "--invoice", type=int,
                        help="Generate invoice with this number (triggers full pipeline)")
    parser.add_argument("--start", type=int, default=1,
                        help="Starting invoice number for batch mode")
    parser.add_argument("--pdf", action="store_true", help="Also generate PDF")
    parser.add_argument("--force-llm", action="store_true",
                        help="Only use LLM (fail if API unavailable)")
    parser.add_argument("--force-ocr", action="store_true",
                        help="Only use OCR (skip LLM entirely)")
    parser.add_argument("--api-key", type=str,
                        help="OpenRouter API key (default: from OPENROUTER_API_KEY env var)")
    parser.add_argument("--version", action="version",
                        version=f"%(prog)s {__version__}")

    args = parser.parse_args()

    api_key = args.api_key or os.environ.get("OPENROUTER_API_KEY")

    if not api_key and not args.force_ocr:
        print(
            "⚠️  No API key found. Set OPENROUTER_API_KEY in .env or use --api-key.\n"
            "    Falling back to OCR-only mode (less accurate).\n",
            file=sys.stderr,
        )
        args.force_ocr = True

    # ── Collect image paths ──
    if args.image:
        image_paths = [Path(args.image)]
    elif args.images:
        images_dir = Path(args.images)
        if not images_dir.is_dir():
            print(f"Error: {args.images} is not a directory.", file=sys.stderr)
            sys.exit(1)
        image_paths = sorted(
            f for f in images_dir.iterdir()
            if f.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
        )
        if not image_paths:
            print(f"No images found in {images_dir}", file=sys.stderr)
            sys.exit(1)
    else:
        parser.print_help()
        sys.exit(1)

    print(f"📸 Found {len(image_paths)} image(s) to process\n")

    try:
        # ── Extract data ──
        dataframe = extract_from_images(
            image_paths, api_key=api_key,
            force_llm=args.force_llm, force_ocr=args.force_ocr,
        )

        print(f"\n📊 Extracted {len(dataframe)} total records")

        # ── Save extracted data ──
        dataframe.to_excel(args.output, index=False, engine="openpyxl")
        print(f"💾 Saved: {args.output}")

        # ── Generate invoice (if requested) ──
        if args.invoice:
            from generate_invoices import generate_invoice, generate_pdf

            excel_out = f"Invoice_{args.invoice}.xlsx"
            workbook = generate_invoice(args.output, args.invoice)
            workbook.save(excel_out)
            print(f"✅ Invoice Excel: {excel_out}")

            if args.pdf:
                pdf_out = f"Invoice_{args.invoice}.pdf"
                generate_pdf(args.output, args.invoice, pdf_out)
                print(f"📄 Invoice PDF:   {pdf_out}")

    except Exception as error:
        print(f"\n❌ Error: {error}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
