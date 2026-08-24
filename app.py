import re
from collections import Counter
from datetime import datetime

import cv2
import numpy as np
import pandas as pd
import pytesseract
import streamlit as st

from PIL import Image
from pdf2image import convert_from_bytes
from pytesseract import Output
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Timesheet Processor",
    page_icon="📋",
    layout="wide"
)


# ============================================================
# CONSTANTS
# ============================================================

DPI = 300

# Normal working hours.
# Change these if your company's normal working hours are different.
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


# Approximate column positions from the supplied timesheet.
#
# PDF table:
#
# # | Date | Start time | End time | Work code |
#   | Short description | Engineer | Total hour
#
COLUMN_RATIOS = [
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


CATEGORIES = {
    "Travel": [
        "travel",
        "travelling",
        "traveling",
        "journey",
        "transit",
        "transport"
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
        "set up"
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


CATEGORY_COLUMNS = {
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


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(text):
    if text is None:
        return ""

    text = str(text)

    text = (
        text
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
    )

    return re.sub(r"\s+", " ", text).strip()


def normalize_ocr(text):
    text = clean_text(text)

    return (
        text
        .replace("O", "0")
        .replace("o", "0")
        .replace("I", "1")
        .replace("l", "1")
    )


# ============================================================
# DATE RECOGNITION
# ============================================================

def parse_date(text):
    text = normalize_ocr(text)

    patterns = [
        r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b",
        r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{2})\b"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text
        )

        if not match:
            continue

        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))

        if year < 100:
            year += 2000

        try:
            return datetime(
                year,
                month,
                day
            )

        except ValueError:
            continue

    return None


def format_date(text):
    date = parse_date(text)

    if date is None:
        return ""

    return date.strftime(
        "%d.%m.%Y"
    )


# ============================================================
# TIME RECOGNITION
# ============================================================

def parse_time(text):

    text = normalize_ocr(text)

    text = text.replace(
        ".",
        ":"
    )

    # Normal HH:MM
    match = re.search(
        r"\b([0-2]?\d):([0-5]\d)\b",
        text
    )

    if match:

        hour = int(match.group(1))
        minute = int(match.group(2))

        if hour <= 23:

            return (
                f"{hour:02d}:"
                f"{minute:02d}"
            )

    # OCR sometimes reads 08:00 as 0800
    match = re.search(
        r"\b([0-2]\d)([0-5]\d)\b",
        text
    )

    if match:

        hour = int(match.group(1))
        minute = int(match.group(2))

        if hour <= 23:

            return (
                f"{hour:02d}:"
                f"{minute:02d}"
            )

    return ""


def time_to_minutes(text):

    value = parse_time(text)

    if not value:
        return None

    hour, minute = map(
        int,
        value.split(":")
    )

    return hour * 60 + minute


# ============================================================
# HOURS / OT CALCULATION
# ============================================================

def calculate_hours(
    start_time,
    end_time
):

    start = time_to_minutes(
        start_time
    )

    end = time_to_minutes(
        end_time
    )

    normal_start = time_to_minutes(
        NORMAL_START
    )

    normal_end = time_to_minutes(
        NORMAL_END
    )

    if None in (
        start,
        end,
        normal_start,
        normal_end
    ):
        return 0.0, 0.0

    # Overnight shift
    if end < start:
        end += 1440

    regular_minutes = 0

    for minute in range(
        start,
        end
    ):

        current = minute % 1440

        if (
            normal_start
            <= current
            < normal_end
        ):
            regular_minutes += 1

    total_minutes = end - start

    overtime_minutes = (
        total_minutes
        - regular_minutes
    )

    regular_hours = round(
        regular_minutes / 60,
        2
    )

    overtime_hours = round(
        overtime_minutes / 60,
        2
    )

    return (
        regular_hours,
        overtime_hours
    )


# ============================================================
# ACTIVITY CLASSIFICATION
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

    # Work code is given priority because
    # it is normally more reliable than OCR
    # of the description.

    for category in [
        "Travel",
        "Waiting",
        "Preparation",
        "Maintenance",
        "Meeting",
        "Training",
        "Working"
    ]:

        for keyword in CATEGORIES[
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

        for keyword in CATEGORIES[
            category
        ]:

            if keyword in combined:
                return category

    return "Other"


# ============================================================
# IMAGE PREPROCESSING
# ============================================================

def prepare_page(pil_image):

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

    # Slight enlargement improves
    # OCR accuracy.
    gray = cv2.resize(
        gray,
        None,
        fx=1.15,
        fy=1.15,
        interpolation=cv2.INTER_CUBIC
    )

    # Improve faded/uneven scans.
    clahe = cv2.createCLAHE(
        clipLimit=2.0,
        tileGridSize=(8, 8)
    )

    gray = clahe.apply(
        gray
    )

    return gray


# ============================================================
# TABLE COLUMN POSITIONS
# ============================================================

def get_column_boundaries(
    width
):

    return [
        int(
            round(
                width * ratio
            )
        )
        for ratio in COLUMN_RATIOS
    ]


# ============================================================
# OCR DATE DETECTION
# ============================================================

def detect_date_rows(
    image,
    date_x1,
    date_x2
):

    crop = image[
        :,
        date_x1:date_x2
    ]

    crop = cv2.resize(
        crop,
        None,
        fx=4,
        fy=4,
        interpolation=cv2.INTER_CUBIC
    )

    # First OCR pass
    config = (
        "--oem 3 "
        "--psm 6 "
        "-c tessedit_char_whitelist=0123456789./-"
    )

    data = pytesseract.image_to_data(
        crop,
        config=config,
        output_type=Output.DATAFRAME
    )

    if data is None or data.empty:
        return []

    data = data.dropna(
        subset=["text"]
    )

    detected = []

    for _, item in data.iterrows():

        text = clean_text(
            item["text"]
        )

        date = parse_date(
            text
        )

        if date is None:
            continue

        try:
            confidence = float(
                item["conf"]
            )
        except Exception:
            confidence = 0

        y = int(
            item["top"] / 4
            + item["height"] / 8
        )

        detected.append(
            (
                y,
                format_date(text),
                confidence
            )
        )

    detected.sort(
        key=lambda x: x[0]
    )

    # Remove duplicate OCR detections
    # of the same physical row.
    unique = []

    for item in detected:

        if not unique:

            unique.append(
                item
            )

            continue

        previous = unique[-1]

        if abs(
            item[0]
            - previous[0]
        ) > 18:

            unique.append(
                item
            )

        elif (
            item[2]
            > previous[2]
        ):

            unique[-1] = item

    return unique


# ============================================================
# CELL OCR
# ============================================================

def ocr_cell(
    crop,
    cell_type="text"
):

    if (
        crop is None
        or crop.size == 0
    ):
        return ""

    height, width = crop.shape[:2]

    # Remove a tiny border around
    # each cell.
    py = max(
        2,
        int(height * 0.05)
    )

    px = max(
        2,
        int(width * 0.025)
    )

    if (
        height > py * 2
        and width > px * 2
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

    else:

        config = (
            "--oem 3 "
            "--psm 6"
        )

    return clean_text(
        pytesseract.image_to_string(
            crop,
            config=config
        )
    )


# ============================================================
# ROW EXTRACTION
# ============================================================

def extract_row(
    image,
    boundaries,
    y1,
    y2,
    page_number,
    row_number
):

    cells = []

    for i in range(8):

        cells.append(
            image[
                y1:y2,
                boundaries[i]:
                boundaries[i + 1]
            ]
        )

    date = format_date(
        ocr_cell(
            cells[1],
            "date"
        )
    )

    if not date:
        return None

    start_time = parse_time(
        ocr_cell(
            cells[2],
            "time"
        )
    )

    end_time = parse_time(
        ocr_cell(
            cells[3],
            "time"
        )
    )

    work_code = ocr_cell(
        cells[4]
    )

    description = ocr_cell(
        cells[5]
    )

    engineer = clean_text(
        ocr_cell(
            cells[6]
        )
    ).upper()

    category = classify_activity(
        work_code,
        description
    )

    if (
        start_time
        and end_time
    ):

        regular_hours, overtime_hours = (
            calculate_hours(
                start_time,
                end_time
            )
        )

    else:

        regular_hours = 0.0
        overtime_hours = 0.0

    return {
        "page": page_number,
        "row": row_number,
        "date": date,
        "engineer": engineer,
        "category": category,
        "regular": regular_hours,
        "ot": overtime_hours
    }


# ============================================================
# ENGINEER NAME
# ============================================================

def determine_engineer(
    names
):

    cleaned = [
        clean_text(name).upper()
        for name in names
        if clean_text(name)
    ]

    if not cleaned:
        return "Unknown"

    return Counter(
        cleaned
    ).most_common(1)[0][0]


# ============================================================
# CREATE FINAL DATAFRAME
# ============================================================

def create_final_dataframe(
    rows
):

    engineer = determine_engineer(
        [
            row["engineer"]
            for row in rows
        ]
    )

    grouped = {}

    for row in rows:

        key = (
            engineer,
            row["date"]
        )

        if key not in grouped:

            grouped[key] = {
                "Engineer Name": engineer,
                "Date": row["date"]
            }

            for column in FINAL_COLUMNS[2:]:

                grouped[key][column] = 0.0

        regular_column, ot_column = (
            CATEGORY_COLUMNS.get(
                row["category"],
                CATEGORY_COLUMNS["Other"]
            )
        )

        grouped[key][
            regular_column
        ] += row["regular"]

        grouped[key][
            ot_column
        ] += row["ot"]

    dataframe = pd.DataFrame(
        list(grouped.values()),
        columns=FINAL_COLUMNS
    )

    # Chronological date order.
    dataframe["_sort_date"] = (
        pd.to_datetime(
            dataframe["Date"],
            format="%d.%m.%Y"
        )
    )

    dataframe = (
        dataframe
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

        dataframe[column] = (
            pd.to_numeric(
                dataframe[column],
                errors="coerce"
            )
            .fillna(0)
            .round(2)
        )

    return dataframe


# ============================================================
# EXCEL FORMATTING
# ============================================================

def create_excel(
    dataframe
):

    from io import BytesIO

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

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
    ].height = 45

    # Freeze engineer/date columns.
    worksheet.freeze_panes = "C2"

    widths = [
        24,
        14,
        16,
        16,
        18,
        18,
        18,
        18,
        20,
        20,
        21,
        21,
        18,
        18,
        18,
        18,
        16,
        17
    ]

    for index, width in enumerate(
        widths,
        start=1
    ):

        worksheet.column_dimensions[
            chr(64 + index)
        ].width = width

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
                and cell.column >= 3
            ):

                cell.number_format = "0.00"

    final_output = BytesIO()

    workbook.save(
        final_output
    )

    final_output.seek(0)

    return final_output.getvalue()


# ============================================================
# PROCESS PDF
# ============================================================

def process_pdf(
    pdf_bytes,
    progress_callback=None
):

    pages = convert_from_bytes(
        pdf_bytes,
        dpi=DPI,
        fmt="png",
        thread_count=2
    )

    all_rows = []

    for page_index, page in enumerate(
        pages
    ):

        page_number = page_index + 1

        if progress_callback:

            progress_callback(
                page_number,
                len(pages)
            )

        image = prepare_page(
            page
        )

        boundaries = get_column_boundaries(
            image.shape[1]
        )

        date_rows = detect_date_rows(
            image,
            boundaries[1],
            boundaries[2]
        )

        if not date_rows:
            continue

        centers = [
            item[0]
            for item in date_rows
        ]

        row_boundaries = [
            max(
                0,
                centers[0] - 35
            )
        ]

        for i in range(
            len(centers) - 1
        ):

            row_boundaries.append(
                (
                    centers[i]
                    + centers[i + 1]
                ) // 2
            )

        row_boundaries.append(
            min(
                image.shape[0],
                centers[-1] + 35
            )
        )

        for row_index in range(
            len(date_rows)
        ):

            row = extract_row(
                image,
                boundaries,
                row_boundaries[row_index],
                row_boundaries[row_index + 1],
                page_number,
                row_index + 1
            )

            if row:
                all_rows.append(
                    row
                )

    if not all_rows:

        raise RuntimeError(
            "No timesheet rows could be detected. "
            "Please make sure the uploaded PDF contains "
            "a clear scanned timesheet."
        )

    dataframe = create_final_dataframe(
        all_rows
    )

    excel_bytes = create_excel(
        dataframe
    )

    return (
        dataframe,
        excel_bytes
    )


# ============================================================
# STREAMLIT USER INTERFACE
# ============================================================

st.title(
    "📋 Timesheet Processor"
)

st.write(
    "Upload your scanned timesheet PDF "
    "and the system will extract the timesheet "
    "information into an Excel file."
)

st.info(
    "The output contains only the "
    "'Detailed Timesheet' tab."
)

uploaded_file = st.file_uploader(
    "Upload scanned timesheet PDF",
    type=["pdf"],
    accept_multiple_files=False
)

if uploaded_file:

    st.success(
        f"File selected: {uploaded_file.name}"
    )

    if st.button(
        "🚀 Process Timesheet",
        type="primary",
        use_container_width=True
    ):

        progress = st.progress(
            0
        )

        status = st.empty()

        try:

            def update_progress(
                page,
                total
            ):

                progress.progress(
                    int(
                        page / total * 100
                    )
                )

                status.write(
                    f"Processing page "
                    f"{page} of {total}..."
                )

            with st.spinner(
                "Scanning PDF and extracting timesheet data..."
            ):

                dataframe, excel_bytes = process_pdf(
                    uploaded_file.getvalue(),
                    update_progress
                )

            progress.progress(
                100
            )

            status.success(
                "Processing complete."
            )

            st.success(
                f"Successfully extracted "
                f"{len(dataframe)} date(s)."
            )

            st.subheader(
                "Detailed Timesheet"
            )

            st.dataframe(
                dataframe,
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                label="⬇️ Download Excel",
                data=excel_bytes,
                file_name="Detailed_Timesheet.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        except Exception as error:

            st.error(
                "The timesheet could not be processed."
            )

            st.exception(
                error
            )

else:

    st.caption(
        "Upload a PDF to begin."
    )
