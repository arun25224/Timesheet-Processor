import streamlit as st
import pandas as pd
import numpy as np
import cv2
import pytesseract
import re
import io
import os
import gc

from collections import Counter
from datetime import datetime
from difflib import SequenceMatcher

from pdf2image import convert_from_bytes, pdfinfo_from_bytes

from openpyxl import load_workbook
from openpyxl.styles import (
    Font,
    Alignment,
    Border,
    Side,
    PatternFill
)
from openpyxl.utils import get_column_letter


# ============================================================
# STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="Timesheet & Invoice Automation",
    layout="wide"
)


# ============================================================
# SETTINGS
# ============================================================

DPI = 300

NORMAL_START = "08:00"
NORMAL_END = "16:00"

FINAL_COLUMNS = [
    "Engineer Name",
    "Date",
    "Travel Time",
    "Travel OT Time",
    "Working Time",
    "Working OT Time",
    "Waiting Time",
    "Waiting OT Time",
    "Preparation Time",
    "Preparation OT Time",
    "Maintenance Time",
    "Maintenance OT Time",
    "Meeting Time",
    "Meeting OT Time",
    "Training Time",
    "Training OT Time",
    "Other Time",
    "Other OT Time"
]


# Exact timesheet column structure from the PDF
#
# 0 = #
# 1 = Date
# 2 = Start time
# 3 = End time
# 4 = Work code
# 5 = Short description
# 6 = Engineer
# 7 = Total hour

EXPECTED_COLUMN_RATIOS = [
    0.000,
    0.043,
    0.205,
    0.325,
    0.445,
    0.610,
    0.772,
    0.934,
    1.000
]


CATEGORY_KEYWORDS = {

    "Travel": [
        "travel",
        "travelling",
        "traveling",
        "journey",
        "transit",
        "transport",
        "transfer"
    ],

    "Waiting": [
        "waiting",
        "wait",
        "standby",
        "stand by"
    ],

    "Preparation": [
        "preparation",
        "prepare",
        "preparing",
        "prep",
        "setup",
        "set up",
        "mobilisation",
        "mobilization"
    ],

    "Maintenance": [
        "maintenance",
        "maint",
        "servicing",
        "service"
    ],

    "Meeting": [
        "meeting",
        "discussion",
        "briefing"
    ],

    "Training": [
        "training",
        "course",
        "induction"
    ],

    "Working": [
        "working",
        "working hours",
        "work",
        "repair",
        "installation",
        "install",
        "overhaul",
        "operation",
        "operating",
        "job",
        "site work"
    ]
}


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value)

    text = (
        text
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def ocr_normalize_text(value):

    text = clean_text(value)

    # Common OCR substitutions
    text = (
        text
        .replace("—", "-")
        .replace("–", "-")
        .replace("|", "I")
    )

    return text


# ============================================================
# DATE PARSING
# ============================================================

def parse_date(value):

    if not value:
        return None

    text = ocr_normalize_text(
        value
    )

    # Only perform these substitutions when looking for dates.
    text = (
        text
        .replace("O", "0")
        .replace("o", "0")
        .replace("I", "1")
        .replace("l", "1")
    )

    patterns = [

        r"\b(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{4})\b",

        r"\b(\d{1,2})\s+(\d{1,2})\s+(\d{4})\b"

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text
        )

        if not match:
            continue

        day = int(
            match.group(1)
        )

        month = int(
            match.group(2)
        )

        year = int(
            match.group(3)
        )

        try:

            return datetime(
                year,
                month,
                day
            )

        except ValueError:

            pass

    return None


def format_date(value):

    parsed = parse_date(
        value
    )

    if parsed is None:
        return ""

    return parsed.strftime(
        "%d.%m.%Y"
    )


# ============================================================
# TIME PARSING
# ============================================================

def normalize_time(value):

    if not value:
        return ""

    text = ocr_normalize_text(
        value
    )

    text = (
        text
        .replace(".", ":")
        .replace(";", ":")
    )

    # 08:00
    match = re.search(
        r"\b([0-2]?\d)\s*:\s*([0-5]\d)\b",
        text
    )

    if match:

        hour = int(
            match.group(1)
        )

        minute = int(
            match.group(2)
        )

        if (
            0 <= hour <= 23
            and
            0 <= minute <= 59
        ):

            return (
                f"{hour:02d}:"
                f"{minute:02d}"
            )

    # 0800
    match = re.search(
        r"\b([0-2]\d)([0-5]\d)\b",
        text
    )

    if match:

        hour = int(
            match.group(1)
        )

        minute = int(
            match.group(2)
        )

        if (
            0 <= hour <= 23
            and
            0 <= minute <= 59
        ):

            return (
                f"{hour:02d}:"
                f"{minute:02d}"
            )

    return ""


def time_to_minutes(value):

    value = normalize_time(
        value
    )

    if not value:
        return None

    try:

        hour, minute = map(
            int,
            value.split(":")
        )

        return (
            hour * 60
            + minute
        )

    except Exception:

        return None


# ============================================================
# OVERTIME CALCULATION
# ============================================================

def calculate_regular_ot(
    start_time,
    end_time
):

    start = time_to_minutes(
        start_time
    )

    end = time_to_minutes(
        end_time
    )

    if (
        start is None
        or
        end is None
    ):

        return 0.0, 0.0

    # Overnight work
    if end < start:

        end += 24 * 60

    normal_start = time_to_minutes(
        NORMAL_START
    )

    normal_end = time_to_minutes(
        NORMAL_END
    )

    regular_minutes = 0
    overtime_minutes = 0

    for minute in range(
        start,
        end
    ):

        current = (
            minute
            % (24 * 60)
        )

        if (
            normal_start
            <= current
            < normal_end
        ):

            regular_minutes += 1

        else:

            overtime_minutes += 1

    return (
        round(
            regular_minutes / 60,
            2
        ),
        round(
            overtime_minutes / 60,
            2
        )
    )


# ============================================================
# CATEGORY CLASSIFICATION
# ============================================================

def classify_activity(
    work_code,
    description
):

    work_code = clean_text(
        work_code
    ).lower()

    description = clean_text(
        description
    ).lower()

    # Work code has priority.
    for category in [
        "Travel",
        "Waiting",
        "Preparation",
        "Maintenance",
        "Meeting",
        "Training",
        "Working"
    ]:

        for keyword in CATEGORY_KEYWORDS[
            category
        ]:

            if keyword in work_code:

                return category

    combined = (
        work_code
        + " "
        + description
    )

    for category in [
        "Travel",
        "Waiting",
        "Preparation",
        "Maintenance",
        "Meeting",
        "Training",
        "Working"
    ]:

        for keyword in CATEGORY_KEYWORDS[
            category
        ]:

            if keyword in combined:

                return category

    return "Other"


# ============================================================
# ENGINEER NAME CLEANING
# ============================================================

def clean_engineer_name(
    value
):

    text = clean_text(
        value
    )

    if not text:
        return ""

    text = re.sub(
        r"[^A-Za-z0-9 .'-]",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip().upper()


def choose_engineer_name(
    names
):

    names = [
        clean_engineer_name(
            name
        )
        for name in names
        if clean_engineer_name(name)
    ]

    if not names:
        return "UNKNOWN"

    groups = []

    for name in names:

        placed = False

        for group in groups:

            similarity = (
                SequenceMatcher(
                    None,
                    name,
                    group[0]
                ).ratio()
            )

            if similarity >= 0.80:

                group.append(
                    name
                )

                placed = True

                break

        if not placed:

            groups.append(
                [name]
            )

    groups.sort(
        key=len,
        reverse=True
    )

    return Counter(
        groups[0]
    ).most_common(1)[0][0]


# ============================================================
# IMAGE PREPROCESSING
# ============================================================

def prepare_page(
    pil_image
):

    image = np.array(
        pil_image
    )

    if image.ndim == 3:

        gray = cv2.cvtColor(
            image,
            cv2.COLOR_RGB2GRAY
        )

    else:

        gray = image

    # Increase resolution slightly.
    gray = cv2.resize(
        gray,
        None,
        fx=1.20,
        fy=1.20,
        interpolation=cv2.INTER_CUBIC
    )

    # Contrast enhancement.
    clahe = cv2.createCLAHE(
        clipLimit=2.0,
        tileGridSize=(8, 8)
    )

    gray = clahe.apply(
        gray
    )

    return gray


# ============================================================
# GRID DETECTION
# ============================================================

def cluster_positions(
    values,
    tolerance=8
):

    if not values:
        return []

    values = sorted(
        int(v)
        for v in values
    )

    clusters = [
        [values[0]]
    ]

    for value in values[1:]:

        current = clusters[-1]

        if (
            abs(
                value
                - np.mean(current)
            )
            <= tolerance
        ):

            current.append(
                value
            )

        else:

            clusters.append(
                [value]
            )

    return [
        int(
            round(
                np.mean(cluster)
            )
        )
        for cluster in clusters
    ]


def detect_vertical_lines(
    image
):

    height, width = image.shape[:2]

    # Binary image.
    binary = cv2.adaptiveThreshold(
        image,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        31,
        15
    )

    vertical_kernel_length = max(
        30,
        width // 40
    )

    vertical_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (
            1,
            vertical_kernel_length
        )
    )

    vertical = cv2.morphologyEx(
        binary,
        cv2.MORPH_OPEN,
        vertical_kernel
    )

    contours, _ = cv2.findContours(
        vertical,
        cv2.RETR_EXTERNAL,
        cv2.CHAIN_APPROX_SIMPLE
    )

    candidates = []

    for contour in contours:

        x, y, w, h = cv2.boundingRect(
            contour
        )

        if (
            h >= height * 0.20
            and
            h > w * 5
        ):

            candidates.append(
                x + w // 2
            )

    candidates = cluster_positions(
        candidates,
        tolerance=12
    )

    # If the table borders are detected,
    # choose the strongest 9 candidates.
    if len(candidates) >= 9:

        # Keep the positions closest to
        # expected table structure.
        left = min(
            candidates
        )

        right = max(
            candidates
        )

        expected = [
            int(
                left
                + ratio
                * (
                    right - left
                )
            )
            for ratio
            in EXPECTED_COLUMN_RATIOS
        ]

        selected = []

        for expected_x in expected:

            nearest = min(
                candidates,
                key=lambda x:
                abs(x - expected_x)
            )

            selected.append(
                nearest
            )

        selected = sorted(
            set(selected)
        )

        if len(selected) == 9:

            return selected

    # Fallback: use known table layout.
    left = int(
        width * 0.015
    )

    right = int(
        width * 0.985
    )

    return [
        int(
            left
            + ratio
            * (
                right - left
            )
        )
        for ratio
        in EXPECTED_COLUMN_RATIOS
    ]


def detect_horizontal_lines(
    image
):

    height, width = image.shape[:2]

    binary = cv2.adaptiveThreshold(
        image,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        31,
        15
    )

    kernel_length = max(
        50,
        width // 30
    )

    horizontal_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (
            kernel_length,
            1
        )
    )

    horizontal = cv2.morphologyEx(
        binary,
        cv2.MORPH_OPEN,
        horizontal_kernel
    )

    contours, _ = cv2.findContours(
        horizontal,
        cv2.RETR_EXTERNAL,
        cv2.CHAIN_APPROX_SIMPLE
    )

    positions = []

    for contour in contours:

        x, y, w, h = cv2.boundingRect(
            contour
        )

        if (
            w >= width * 0.30
            and
            w > h * 8
        ):

            positions.append(
                y + h // 2
            )

    positions = cluster_positions(
        positions,
        tolerance=10
    )

    positions = sorted(
        positions
    )

    return positions


# ============================================================
# CELL OCR
# ============================================================

def crop_cell(
    image,
    x1,
    y1,
    x2,
    y2
):

    height, width = image.shape[:2]

    x1 = max(
        0,
        int(x1)
    )

    y1 = max(
        0,
        int(y1)
    )

    x2 = min(
        width,
        int(x2)
    )

    y2 = min(
        height,
        int(y2)
    )

    if (
        x2 <= x1
        or
        y2 <= y1
    ):

        return None

    crop = image[
        y1:y2,
        x1:x2
    ]

    if crop.size == 0:
        return None

    return crop


def preprocess_cell(
    crop
):

    if crop is None:
        return None

    height, width = crop.shape[:2]

    # Remove a small border area.
    px = max(
        2,
        int(width * 0.025)
    )

    py = max(
        2,
        int(height * 0.06)
    )

    if (
        width > px * 2
        and
        height > py * 2
    ):

        crop = crop[
            py:height-py,
            px:width-px
        ]

    crop = cv2.resize(
        crop,
        None,
        fx=4,
        fy=4,
        interpolation=cv2.INTER_CUBIC
    )

    # Light denoising.
    crop = cv2.GaussianBlur(
        crop,
        (3, 3),
        0
    )

    return crop


def tesseract_cell(
    crop,
    cell_type="text"
):

    crop = preprocess_cell(
        crop
    )

    if crop is None:
        return ""

    if cell_type == "date":

        config = (
            "--oem 3 "
            "--psm 7 "
            "-c tessedit_char_whitelist=0123456789./-"
        )

    elif cell_type == "time":

        config = (
            "--oem 3 "
            "--psm 7 "
            "-c tessedit_char_whitelist=0123456789:."
        )

    elif cell_type == "number":

        config = (
            "--oem 3 "
            "--psm 7 "
            "-c tessedit_char_whitelist=0123456789."
        )

    else:

        config = (
            "--oem 3 "
            "--psm 6"
        )

    result = pytesseract.image_to_string(
        crop,
        config=config
    )

    return clean_text(
        result
    )


# ============================================================
# WHOLE PAGE OCR
# ============================================================

def page_word_data(
    image
):

    # Whole page OCR gives us a second OCR pass.
    # It is used to locate row centres and to improve
    # detection where table lines are weak.

    config = (
        "--oem 3 "
        "--psm 6"
    )

    try:

        data = pytesseract.image_to_data(
            image,
            config=config,
            output_type=pytesseract.Output.DATAFRAME
        )

    except Exception:

        return pd.DataFrame()

    if data is None or data.empty:
        return pd.DataFrame()

    data = data.dropna(
        subset=["text"]
    )

    data["text"] = data["text"].astype(
        str
    ).map(
        clean_text
    )

    data = data[
        data["text"] != ""
    ]

    return data


# ============================================================
# FIND TABLE ROWS
# ============================================================

def get_row_boundaries(
    image,
    horizontal_lines,
    word_data
):

    height, width = image.shape[:2]

    # Use detected horizontal table lines first.
    if len(horizontal_lines) >= 4:

        lines = [
            y
            for y in horizontal_lines
            if 0 <= y <= height
        ]

        lines = sorted(
            set(lines)
        )

        # Remove lines too close together.
        filtered = []

        for line in lines:

            if (
                not filtered
                or
                line - filtered[-1] > 10
            ):

                filtered.append(
                    line
                )

        if len(filtered) >= 4:

            return filtered

    # Fallback using OCR line centres.
    if (
        word_data is not None
        and
        not word_data.empty
    ):

        centers = []

        for _, row in word_data.iterrows():

            try:

                text = clean_text(
                    row["text"]
                )

                if not text:
                    continue

                y = (
                    float(row["top"])
                    + float(row["height"]) / 2
                )

                centers.append(
                    y
                )

            except Exception:

                pass

        centers.sort()

        groups = []

        for center in centers:

            if not groups:

                groups.append(
                    [center]
                )

                continue

            if (
                abs(
                    center
                    - np.mean(
                        groups[-1]
                    )
                )
                <= 22
            ):

                groups[-1].append(
                    center
                )

            else:

                groups.append(
                    [center]
                )

        row_centers = [
            int(
                round(
                    np.mean(group)
                )
            )
            for group in groups
        ]

        if len(row_centers) >= 2:

            boundaries = [
                max(
                    0,
                    row_centers[0] - 25
                )
            ]

            for i in range(
                len(row_centers) - 1
            ):

                boundaries.append(
                    int(
                        (
                            row_centers[i]
                            +
                            row_centers[i + 1]
                        ) / 2
                    )
                )

            boundaries.append(
                min(
                    height,
                    row_centers[-1] + 25
                )
            )

            return boundaries

    # Last-resort fallback.
    return [
        0,
        height
    ]


# ============================================================
# EXTRACT A TABLE ROW
# ============================================================

def extract_table_row(
    image,
    x_bounds,
    y1,
    y2
):

    if (
        y2 - y1
        < 15
    ):

        return None

    cells = []

    for i in range(8):

        cell = crop_cell(
            image,
            x_bounds[i],
            y1,
            x_bounds[i + 1],
            y2
        )

        cells.append(
            cell
        )

    # OCR each important cell separately.
    row_number = tesseract_cell(
        cells[0],
        "number"
    )

    date_text = tesseract_cell(
        cells[1],
        "date"
    )

    start_text = tesseract_cell(
        cells[2],
        "time"
    )

    end_text = tesseract_cell(
        cells[3],
        "time"
    )

    work_code = tesseract_cell(
        cells[4],
        "text"
    )

    description = tesseract_cell(
        cells[5],
        "text"
    )

    engineer = tesseract_cell(
        cells[6],
        "text"
    )

    total_hour = tesseract_cell(
        cells[7],
        "number"
    )

    date_value = format_date(
        date_text
    )

    start_value = normalize_time(
        start_text
    )

    end_value = normalize_time(
        end_text
    )

    work_code = clean_text(
        work_code
    )

    description = clean_text(
        description
    )

    engineer = clean_engineer_name(
        engineer
    )

    total_hour = clean_text(
        total_hour
    )

    # Ignore header rows.
    header_text = (
        work_code
        + " "
        + description
        + " "
        + engineer
    ).lower()

    if (
        "work code" in header_text
        or
        "short description" in header_text
        or
        "engineer" in header_text
    ):

        return None

    # Ignore empty rows.
    if (
        not date_value
        and
        not start_value
        and
        not end_value
        and
        not work_code
        and
        not description
        and
        not engineer
    ):

        return None

    return {
        "Date": date_value,
        "Start Time": start_value,
        "End Time": end_value,
        "Work Code": work_code,
        "Short Description": description,
        "Engineer Name": engineer,
        "Total Hour OCR": total_hour
    }


# ============================================================
# PROCESS ONE PAGE
# ============================================================

def process_page(
    pil_image,
    page_number
):

    image = prepare_page(
        pil_image
    )

    height, width = image.shape[:2]

    x_bounds = detect_vertical_lines(
        image
    )

    horizontal_lines = detect_horizontal_lines(
        image
    )

    word_data = page_word_data(
        image
    )

    # We need at least 9 vertical boundaries.
    if len(x_bounds) != 9:

        left = int(
            width * 0.015
        )

        right = int(
            width * 0.985
        )

        x_bounds = [
            int(
                left
                + ratio
                * (
                    right - left
                )
            )
            for ratio
            in EXPECTED_COLUMN_RATIOS
        ]

    y_bounds = get_row_boundaries(
        image,
        horizontal_lines,
        word_data
    )

    rows = []

    # When horizontal grid lines exist,
    # every interval is a candidate row.
    for i in range(
        len(y_bounds) - 1
    ):

        y1 = y_bounds[i]
        y2 = y_bounds[i + 1]

        row = extract_table_row(
            image,
            x_bounds,
            y1,
            y2
        )

        if row is None:
            continue

        row["Page"] = page_number

        rows.append(
            row
        )

    del image
    del word_data

    gc.collect()

    return rows


# ============================================================
# CLEAN AND VALIDATE RAW ROWS
# ============================================================

def clean_raw_rows(
    raw_rows
):

    if not raw_rows:
        return []

    cleaned = []

    last_date = None
    last_engineer = None

    for row in raw_rows:

        current_date = (
            row.get("Date", "")
        )

        current_engineer = (
            clean_engineer_name(
                row.get(
                    "Engineer Name",
                    ""
                )
            )
        )

        # Carry date down when the PDF has
        # blank repeated date cells.
        if current_date:

            last_date = current_date

        else:

            current_date = last_date

        # Carry engineer name down.
        if current_engineer:

            last_engineer = (
                current_engineer
            )

        else:

            current_engineer = (
                last_engineer
            )

        start_time = normalize_time(
            row.get(
                "Start Time",
                ""
            )
        )

        end_time = normalize_time(
            row.get(
                "End Time",
                ""
            )
        )

        work_code = clean_text(
            row.get(
                "Work Code",
                ""
            )
        )

        description = clean_text(
            row.get(
                "Short Description",
                ""
            )
        )

        # Reject rows that are clearly not data.
        if not current_date:
            continue

        if (
            not start_time
            and
            not end_time
            and
            not work_code
            and
            not description
        ):
            continue

        # If OCR produced a date but the row is obviously a header.
        header = (
            work_code
            + " "
            + description
        ).lower()

        if (
            "work code" in header
            or
            "short description" in header
        ):
            continue

        cleaned.append(
            {
                "Date": current_date,
                "Start Time": start_time,
                "End Time": end_time,
                "Work Code": work_code,
                "Short Description": description,
                "Engineer Name": current_engineer
            }
        )

    return cleaned


# ============================================================
# CREATE FINAL DETAILED TIMESHEET
# ============================================================

def create_final_dataframe(
    raw_rows
):

    if not raw_rows:

        return pd.DataFrame(
            columns=FINAL_COLUMNS
        )

    engineer_names = [
        row["Engineer Name"]
        for row in raw_rows
        if row["Engineer Name"]
    ]

    canonical_engineer = (
        choose_engineer_name(
            engineer_names
        )
    )

    aggregated = {}

    for row in raw_rows:

        engineer = (
            row["Engineer Name"]
            or canonical_engineer
        )

        # Fix OCR variants of engineer name.
        if engineer:
            engineer = clean_engineer_name(
                engineer
            )

        date_value = row["Date"]

        if not date_value:
            continue

        category = classify_activity(
            row["Work Code"],
            row["Short Description"]
        )

        regular_hours = 0.0
        overtime_hours = 0.0

        if (
            row["Start Time"]
            and
            row["End Time"]
        ):

            regular_hours, overtime_hours = (
                calculate_regular_ot(
                    row["Start Time"],
                    row["End Time"]
                )
            )

        # All work on weekends is OT.
        parsed = parse_date(
            date_value
        )

        if (
            parsed is not None
            and
            parsed.weekday() >= 5
        ):

            overtime_hours = (
                regular_hours
                + overtime_hours
            )

            regular_hours = 0.0

        key = (
            engineer,
            date_value
        )

        if key not in aggregated:

            aggregated[key] = {
                "Engineer Name": engineer,
                "Date": date_value
            }

            for column in FINAL_COLUMNS[2:]:

                aggregated[key][
                    column
                ] = 0.0

        mapping = {

            "Travel": (
                "Travel Time",
                "Travel OT Time"
            ),

            "Working": (
                "Working Time",
                "Working OT Time"
            ),

            "Waiting": (
                "Waiting Time",
                "Waiting OT Time"
            ),

            "Preparation": (
                "Preparation Time",
                "Preparation OT Time"
            ),

            "Maintenance": (
                "Maintenance Time",
                "Maintenance OT Time"
            ),

            "Meeting": (
                "Meeting Time",
                "Meeting OT Time"
            ),

            "Training": (
                "Training Time",
                "Training OT Time"
            ),

            "Other": (
                "Other Time",
                "Other OT Time"
            )
        }

        regular_col, ot_col = mapping.get(
            category,
            mapping["Other"]
        )

        aggregated[key][
            regular_col
        ] += regular_hours

        aggregated[key][
            ot_col
        ] += overtime_hours

    final_df = pd.DataFrame(
        list(
            aggregated.values()
        )
    )

    if final_df.empty:

        return pd.DataFrame(
            columns=FINAL_COLUMNS
        )

    for column in FINAL_COLUMNS:

        if column not in final_df.columns:

            if column in [
                "Engineer Name",
                "Date"
            ]:

                final_df[column] = ""

            else:

                final_df[column] = 0.0

    final_df = final_df[
        FINAL_COLUMNS
    ]

    final_df["_actual_date"] = (
        final_df["Date"]
        .apply(parse_date)
    )

    final_df = (
        final_df
        .sort_values(
            [
                "_actual_date",
                "Engineer Name"
            ],
            kind="stable"
        )
        .drop(
            columns="_actual_date"
        )
        .reset_index(
            drop=True
        )
    )

    for column in FINAL_COLUMNS[2:]:

        final_df[column] = (
            pd.to_numeric(
                final_df[column],
                errors="coerce"
            )
            .fillna(0.0)
            .round(2)
        )

    return final_df


# ============================================================
# CREATE CLIENT TOTALS FOR INVOICE
# ============================================================

def create_invoice_summary(
    dataframe
):

    if dataframe is None or dataframe.empty:

        return {
            "travel": 0.0,
            "nt": 0.0,
            "ot": 0.0,
            "waiting": 0.0,
            "preparation": 0.0,
            "local_transport": 0.0
        }

    travel = (
        dataframe["Travel Time"]
        + dataframe["Travel OT Time"]
    ).sum()

    nt = (
        dataframe["Working Time"]
    ).sum()

    ot = (
        dataframe["Working OT Time"]
    ).sum()

    waiting = (
        dataframe["Waiting Time"]
        + dataframe["Waiting OT Time"]
    ).sum()

    preparation = (
        dataframe["Preparation Time"]
        + dataframe["Preparation OT Time"]
    ).sum()

    # Existing invoice logic:
    # Local transport = 1 when there is activity but no waiting.
    local_transport = 0

    for _, row in dataframe.iterrows():

        activity = (
            row["Travel Time"]
            + row["Travel OT Time"]
            + row["Working Time"]
            + row["Working OT Time"]
            + row["Preparation Time"]
            + row["Preparation OT Time"]
            + row["Waiting Time"]
            + row["Waiting OT Time"]
        )

        waiting_total = (
            row["Waiting Time"]
            + row["Waiting OT Time"]
        )

        if (
            activity > 0
            and
            waiting_total == 0
        ):

            local_transport += 1

    return {
        "travel": round(
            float(travel),
            2
        ),
        "nt": round(
            float(nt),
            2
        ),
        "ot": round(
            float(ot),
            2
        ),
        "waiting": round(
            float(waiting),
            2
        ),
        "preparation": round(
            float(preparation),
            2
        ),
        "local_transport": round(
            float(local_transport),
            2
        )
    }


# ============================================================
# EXCEL GENERATION
# ============================================================

def create_detailed_timesheet_excel(
    dataframe
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # ONLY ONE TAB.
        dataframe.to_excel(
            writer,
            sheet_name="Detailed Timesheet",
            index=False
        )

    output.seek(0)

    workbook = load_workbook(
        output
    )

    worksheet = workbook[
        "Detailed Timesheet"
    ]

    thin = Side(
        style="thin"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # Header
    for cell in worksheet[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.border = border

    worksheet.row_dimensions[
        1
    ].height = 42

    worksheet.freeze_panes = "C2"

    # Body
    for row in worksheet.iter_rows():

        for cell in row:

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            if (
                cell.row > 1
                and
                cell.column >= 3
            ):

                cell.number_format = (
                    "0.00"
                )

    # Column widths.
    widths = {
        "A": 24,
        "B": 14,
        "C": 16,
        "D": 18,
        "E": 16,
        "F": 18,
        "G": 16,
        "H": 18,
        "I": 20,
        "J": 22,
        "K": 20,
        "L": 22,
        "M": 17,
        "N": 19,
        "O": 17,
        "P": 19,
        "Q": 15,
        "R": 17
    }

    for column, width in widths.items():

        worksheet.column_dimensions[
            column
        ].width = width

    final_output = io.BytesIO()

    workbook.save(
        final_output
    )

    final_output.seek(0)

    return final_output.getvalue()


# ============================================================
# INVOICE TEMPLATE PROCESSING
# ============================================================

def generate_invoice_from_template(
    template_bytes,
    dataframe,
    customer_name,
    invoicing_address,
    delivery_address,
    reference,
    customer_po,
    project_no,
    service_type,
    vessel_name,
    vessel_no,
    engineer_name,
    position
):

    summary = create_invoice_summary(
        dataframe
    )

    template_stream = io.BytesIO(
        template_bytes
    )

    workbook = load_workbook(
        template_stream
    )

    if "SG" not in workbook.sheetnames:

        raise ValueError(
            "The invoice template does not contain "
            "an 'SG' worksheet."
        )

    worksheet = workbook[
        "SG"
    ]

    # ========================================================
    # CUSTOMER / PROJECT DETAILS
    # ========================================================

    worksheet["C7"] = (
        customer_name
    )

    worksheet["C8"] = (
        invoicing_address
    )

    worksheet["C9"] = (
        delivery_address
    )

    worksheet["C10"] = (
        reference
    )

    worksheet["C11"] = (
        customer_po
    )

    worksheet["C12"] = (
        project_no
    )

    worksheet["C13"] = (
        service_type
    )

    worksheet["C14"] = (
        vessel_name
    )

    worksheet["C15"] = (
        vessel_no
    )

    # ========================================================
    # ENGINEER POSITION
    # ========================================================

    position_offsets = {

        "Service Technician": 20,

        "Service Engineer": 30,

        "Senior Service Engineer": 40,

        "Specialist Service Engineer": 50
    }

    r_offset = position_offsets.get(
        position,
        20
    )

    # ========================================================
    # HOURS
    # ========================================================

    worksheet[
        f"D{r_offset + 1}"
    ] = (
        summary["travel"]
        if summary["travel"] > 0
        else ""
    )

    worksheet[
        f"D{r_offset + 2}"
    ] = (
        summary["nt"]
        if summary["nt"] > 0
        else ""
    )

    worksheet[
        f"D{r_offset + 3}"
    ] = (
        summary["ot"]
        if summary["ot"] > 0
        else ""
    )

    worksheet[
        f"D{r_offset + 4}"
    ] = (
        summary["waiting"]
        if summary["waiting"] > 0
        else ""
    )

    worksheet[
        f"D{r_offset + 5}"
    ] = (
        summary["preparation"]
        if summary["preparation"] > 0
        else ""
    )

    # ========================================================
    # EXPENSES
    # ========================================================

    expense_row = None
    local_transport_row = None

    for row_index in range(
        1,
        min(
            worksheet.max_row,
            300
        ) + 1
    ):

        col_b = worksheet.cell(
            row=row_index,
            column=2
        ).value

        col_c = worksheet.cell(
            row=row_index,
            column=3
        ).value

        if (
            col_b
            and
            str(col_b).strip().lower()
            == "expenses"
        ):

            expense_row = (
                row_index
            )

        if (
            col_c
            and
            "local transport"
            in str(col_c).lower()
        ):

            local_transport_row = (
                row_index
            )

    if expense_row:

        worksheet.cell(
            row=expense_row + 2,
            column=3
        ).value = (
            engineer_name
        )

    if (
        local_transport_row
        and
        summary["local_transport"] > 0
    ):

        worksheet.cell(
            row=local_transport_row,
            column=5
        ).value = (
            summary["local_transport"]
        )

    # ========================================================
    # SAVE
    # ========================================================

    output = io.BytesIO()

    workbook.save(
        output
    )

    output.seek(0)

    return output.getvalue()


# ============================================================
# SESSION STATE
# ============================================================

if "timesheet_df" not in st.session_state:

    st.session_state[
        "timesheet_df"
    ] = None

if "timesheet_excel" not in st.session_state:

    st.session_state[
        "timesheet_excel"
    ] = None


# ============================================================
# PAGE HEADER
# ============================================================

st.title(
    "Timesheet & Invoice Automation"
)

st.caption(
    "Tesseract OCR • Multi-page PDF processing • "
    "Overtime calculation • Invoice generation"
)


# ============================================================
# TABS
# ============================================================

timesheet_tab, invoice_tab = st.tabs(
    [
        "📋 Timesheet Processing",
        "🧾 Invoice Generation"
    ]
)


# ============================================================
# TAB 1 — TIMESHEET
# ============================================================

with timesheet_tab:

    st.header(
        "Timesheet Processing"
    )

    st.write(
        "Upload your scanned timesheet PDF. "
        "Every page in the document will be processed."
    )

    uploaded_pdf = st.file_uploader(
        "Upload Timesheet PDF",
        type=["pdf"],
        key="timesheet_pdf"
    )

    if uploaded_pdf is not None:

        st.success(
            f"Selected: {uploaded_pdf.name}"
        )

        if st.button(
            "🔍 Scan & Process ALL Pages",
            type="primary",
            use_container_width=True
        ):

            try:

                pdf_bytes = (
                    uploaded_pdf.getvalue()
                )

                pdf_info = (
                    pdfinfo_from_bytes(
                        pdf_bytes
                    )
                )

                total_pages = int(
                    pdf_info["Pages"]
                )

                st.info(
                    f"PDF contains "
                    f"{total_pages} page(s). "
                    f"All pages will be scanned."
                )

                progress = st.progress(
                    0
                )

                status = st.empty()

                all_raw_rows = []

                for page_number in range(
                    1,
                    total_pages + 1
                ):

                    status.write(
                        f"Scanning page "
                        f"{page_number} "
                        f"of "
                        f"{total_pages}..."
                    )

                    progress.progress(
                        int(
                            (
                                page_number - 1
                            )
                            / total_pages
                            * 100
                        )
                    )

                    # Render ONE page at a time.
                    # This prevents memory issues on Streamlit.
                    page_images = (
                        convert_from_bytes(
                            pdf_bytes,
                            dpi=DPI,
                            fmt="png",
                            first_page=page_number,
                            last_page=page_number,
                            thread_count=1
                        )
                    )

                    if not page_images:

                        continue

                    page_rows = process_page(
                        page_images[0],
                        page_number
                    )

                    all_raw_rows.extend(
                        page_rows
                    )

                    del page_images

                    gc.collect()

                progress.progress(
                    100
                )

                status.write(
                    "Cleaning OCR results..."
                )

                cleaned_rows = clean_raw_rows(
                    all_raw_rows
                )

                if not cleaned_rows:

                    st.error(
                        "No usable timesheet rows "
                        "were detected."
                    )

                    st.info(
                        "Make sure the PDF is a clear "
                        "scan of the timesheet."
                    )

                else:

                    final_df = (
                        create_final_dataframe(
                            cleaned_rows
                        )
                    )

                    if final_df.empty:

                        st.error(
                            "The PDF was scanned, "
                            "but no valid timesheet "
                            "entries could be created."
                        )

                    else:

                        excel_bytes = (
                            create_detailed_timesheet_excel(
                                final_df
                            )
                        )

                        st.session_state[
                            "timesheet_df"
                        ] = final_df.copy()

                        st.session_state[
                            "timesheet_excel"
                        ] = excel_bytes

                        st.success(
                            f"Completed. "
                            f"{len(cleaned_rows)} "
                            f"timesheet entries found "
                            f"across "
                            f"{total_pages} page(s)."
                        )

            except Exception as error:

                st.error(
                    "An error occurred while "
                    "processing the PDF."
                )

                st.exception(
                    error
                )

    # --------------------------------------------------------
    # DISPLAY RESULTS
    # --------------------------------------------------------

    if (
        st.session_state[
            "timesheet_df"
        ]
        is not None
    ):

        st.divider()

        st.subheader(
            "Detailed Timesheet"
        )

        st.info(
            "Review the OCR result below. "
            "You can directly correct any OCR mistakes "
            "before generating the invoice."
        )

        edited_df = st.data_editor(
            st.session_state[
                "timesheet_df"
            ],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="detailed_timesheet_editor"
        )

        # ----------------------------------------------------
        # SAVE EDITED DATA
        # ----------------------------------------------------

        if st.button(
            "💾 Save Timesheet Changes",
            use_container_width=True
        ):

            edited_df["_sort_date"] = (
                edited_df["Date"]
                .apply(parse_date)
            )

            edited_df = (
                edited_df
                .sort_values(
                    "_sort_date",
                    kind="stable"
                )
                .drop(
                    columns="_sort_date"
                )
                .reset_index(
                    drop=True
                )
            )

            for column in FINAL_COLUMNS[2:]:

                edited_df[column] = (
                    pd.to_numeric(
                        edited_df[column],
                        errors="coerce"
                    )
                    .fillna(0.0)
                    .round(2)
                )

            st.session_state[
                "timesheet_df"
            ] = edited_df

            st.session_state[
                "timesheet_excel"
            ] = create_detailed_timesheet_excel(
                edited_df
            )

            st.success(
                "Timesheet saved and re-sorted chronologically."
            )

        # ----------------------------------------------------
        # DOWNLOAD EXCEL
        # ----------------------------------------------------

        st.download_button(
            "⬇️ Download Detailed Timesheet Excel",
            data=st.session_state[
                "timesheet_excel"
            ],
            file_name="Detailed_Timesheet.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )


# ============================================================
# TAB 2 — INVOICE GENERATION
# ============================================================

with invoice_tab:

    st.header(
        "Invoice Generation"
    )

    if (
        st.session_state[
            "timesheet_df"
        ]
        is None
    ):

        st.warning(
            "Please process the timesheet in "
            "the first tab before generating an invoice."
        )

    else:

        dataframe = (
            st.session_state[
                "timesheet_df"
            ]
        )

        st.success(
            "Processed timesheet is ready "
            "for invoice generation."
        )

        # ====================================================
        # INVOICE TEMPLATE
        # ====================================================

        st.subheader(
            "1. Invoice Template"
        )

        template_excel = st.file_uploader(
            "Upload Blank Invoice Template (.xlsx)",
            type=["xlsx"],
            key="invoice_template"
        )

        st.caption(
            "The template must contain an 'SG' worksheet."
        )

        # ====================================================
        # CUSTOMER INFORMATION
        # ====================================================

        st.subheader(
            "2. Customer Details"
        )

        col1, col2 = st.columns(2)

        with col1:

            cust_name = st.text_input(
                "Customer Name",
                key="cust_name"
            )

            inv_address = st.text_area(
                "Invoicing Address",
                key="inv_address"
            )

            del_address = st.text_area(
                "Delivery Address",
                key="del_address"
            )

            reference = st.text_input(
                "Reference",
                key="reference"
            )

            cust_po = st.text_input(
                "Customer PO",
                key="cust_po"
            )

        with col2:

            proj_no = st.text_input(
                "Project No",
                key="proj_no"
            )

            svc_type = st.text_input(
                "Service Type",
                key="svc_type"
            )

            vessel_name = st.text_input(
                "Vessel Name",
                key="vessel_name"
            )

            vessel_no = st.text_input(
                "Vessel No.",
                key="vessel_no"
            )

            engineer_name_invoice = st.text_input(
                "Engineer Name (For Expenses)",
                key="engineer_invoice"
            )

        # ====================================================
        # ENGINEER POSITION
        # ====================================================

        st.subheader(
            "3. Engineer Position"
        )

        position = st.selectbox(
            "Assign Hours to Position",
            [
                "Service Technician",
                "Service Engineer",
                "Senior Service Engineer",
                "Specialist Service Engineer"
            ],
            key="engineer_position"
        )

        # ====================================================
        # TIMESHEET SUMMARY
        # ====================================================

        st.subheader(
            "4. Timesheet Summary"
        )

        summary = create_invoice_summary(
            dataframe
        )

        s1, s2, s3 = st.columns(3)

        with s1:

            st.metric(
                "Travel",
                f"{summary['travel']:.2f} h"
            )

            st.metric(
                "Normal Working",
                f"{summary['nt']:.2f} h"
            )

        with s2:

            st.metric(
                "Overtime",
                f"{summary['ot']:.2f} h"
            )

            st.metric(
                "Waiting",
                f"{summary['waiting']:.2f} h"
            )

        with s3:

            st.metric(
                "Preparation",
                f"{summary['preparation']:.2f} h"
            )

            st.metric(
                "Local Transport",
                f"{summary['local_transport']:.0f}"
            )

        # ====================================================
        # GENERATE INVOICE
        # ====================================================

        st.subheader(
            "5. Generate Invoice"
        )

        if st.button(
            "🧾 Generate Final Invoice",
            type="primary",
            use_container_width=True
        ):

            if template_excel is None:

                st.error(
                    "Please upload the blank invoice "
                    "template first."
                )

            else:

                try:

                    invoice_bytes = (
                        generate_invoice_from_template(
                            template_excel.getvalue(),
                            dataframe,
                            cust_name,
                            inv_address,
                            del_address,
                            reference,
                            cust_po,
                            proj_no,
                            svc_type,
                            vessel_name,
                            vessel_no,
                            engineer_name_invoice,
                            position
                        )
                    )

                    st.success(
                        "Invoice Generated Successfully."
                    )

                    filename_customer = (
                        re.sub(
                            r"[^A-Za-z0-9_-]",
                            "_",
                            cust_name
                            or
                            "Completed"
                        )
                    )

                    st.download_button(
                        "⬇️ Download Final Invoice",
                        data=invoice_bytes,
                        file_name=(
                            f"Invoice_"
                            f"{filename_customer}.xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        use_container_width=True
                    )

                except Exception as error:

                    st.error(
                        "Invoice generation failed."
                    )

                    st.exception(
                        error
                    )
