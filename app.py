import argparse
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

import cv2
import numpy as np
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from pytesseract import Output
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

DPI = 300
NORMAL_START = "08:00"
NORMAL_END = "16:00"

COLUMNS = [
    "Engineer Name", "Date",
    "Travel Time", "Travel OT Time",
    "Working Time", "Working OT Time",
    "Waiting Time", "Waiting OT Time",
    "Preparation Time", "Preparation OT Time",
    "Maintenance Time", "Maintenance OT Time",
    "Meeting Time", "Meeting OT Time",
    "Training Time", "Training OT Time",
    "Other Time", "Other OT Time"
]

COLUMN_RATIOS = [0.000, 0.043, 0.205, 0.325, 0.445, 0.610, 0.772, 0.934, 1.000]

CATEGORIES = {
    "Travel": ["travel", "travelling", "traveling", "journey", "transit", "transport"],
    "Waiting": ["waiting", "wait", "standby", "stand by"],
    "Preparation": ["preparation", "prepare", "preparing", "prep", "setup", "set up"],
    "Maintenance": ["maintenance", "maint", "servicing", "service"],
    "Meeting": ["meeting", "discussion", "briefing"],
    "Training": ["training", "course", "induction"],
    "Working": [
        "working", "work", "repair", "installation", "install",
        "overhaul", "operation", "operating", "job", "site work"
    ]
}

CATEGORY_COLUMNS = {
    "Travel": ("Travel Time", "Travel OT Time"),
    "Working": ("Working Time", "Working OT Time"),
    "Waiting": ("Waiting Time", "Waiting OT Time"),
    "Preparation": ("Preparation Time", "Preparation OT Time"),
    "Maintenance": ("Maintenance Time", "Maintenance OT Time"),
    "Meeting": ("Meeting Time", "Meeting OT Time"),
    "Training": ("Training Time", "Training OT Time"),
    "Other": ("Other Time", "Other OT Time")
}


def clean(text):
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).replace("\n", " ").replace("\r", " ")).strip()


def normalize_ocr(text):
    return clean(text).replace("O", "0").replace("o", "0").replace("I", "1").replace("l", "1")


def parse_date(text):
    text = normalize_ocr(text)
    match = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b", text)
    if not match:
        return None
    try:
        return datetime(
            int(match.group(3)),
            int(match.group(2)),
            int(match.group(1))
        )
    except ValueError:
        return None


def format_date(text):
    date = parse_date(text)
    return date.strftime("%d.%m.%Y") if date else ""


def parse_time(text):
    text = normalize_ocr(text).replace(".", ":")
    match = re.search(r"\b([0-2]?\d):([0-5]\d)\b", text)
    if not match:
        match = re.search(r"\b([0-2]\d)([0-5]\d)\b", text)
    if not match:
        return ""
    hour, minute = int(match.group(1)), int(match.group(2))
    return f"{hour:02d}:{minute:02d}" if hour <= 23 else ""


def time_minutes(text):
    value = parse_time(text)
    if not value:
        return None
    hour, minute = map(int, value.split(":"))
    return hour * 60 + minute


def calculate_hours(start, end):
    start = time_minutes(start)
    end = time_minutes(end)
    normal_start = time_minutes(NORMAL_START)
    normal_end = time_minutes(NORMAL_END)

    if None in (start, end, normal_start, normal_end):
        return 0.0, 0.0

    if end < start:
        end += 1440

    regular_minutes = 0
    for minute in range(start, end):
        current = minute % 1440
        if normal_start <= current < normal_end:
            regular_minutes += 1

    total_minutes = end - start
    ot_minutes = total_minutes - regular_minutes

    return round(regular_minutes / 60, 2), round(ot_minutes / 60, 2)


def classify_activity(work_code, description):
    work_code = clean(work_code).lower()
    combined = f"{work_code} {clean(description).lower()}"

    for category in ["Travel", "Waiting", "Preparation", "Maintenance",
                     "Meeting", "Training", "Working"]:
        if any(word in work_code for word in CATEGORIES[category]):
            return category

    for category in ["Travel", "Waiting", "Preparation", "Maintenance",
                     "Meeting", "Training", "Working"]:
        if any(word in combined for word in CATEGORIES[category]):
            return category

    return "Other"


def prepare_page(pil_image):
    image = np.array(pil_image)
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY) if image.ndim == 3 else image
    gray = cv2.resize(gray, None, fx=1.15, fy=1.15, interpolation=cv2.INTER_CUBIC)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    return clahe.apply(gray)


def column_boundaries(width):
    return [int(round(width * ratio)) for ratio in COLUMN_RATIOS]


def find_dates(image, x1, x2):
    crop = image[:, x1:x2]
    crop = cv2.resize(crop, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)

    config = "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789./-"
    data = pytesseract.image_to_data(
        crop,
        config=config,
        output_type=Output.DATAFRAME
    )

    if data is None or data.empty:
        return []

    data = data.dropna(subset=["text"])
    results = []

    for _, row in data.iterrows():
        text = clean(row["text"])
        date = parse_date(text)

        if not date:
            continue

        try:
            confidence = float(row["conf"])
        except Exception:
            confidence = 0

        y = int(row["top"] / 4 + row["height"] / 8)
        results.append((y, format_date(text), confidence))

    results.sort(key=lambda x: x[0])

    unique = []
    for result in results:
        if not unique or abs(result[0] - unique[-1][0]) > 15:
            unique.append(result)
        elif result[2] > unique[-1][2]:
            unique[-1] = result

    return unique


def ocr_cell(crop, cell_type="text"):
    if crop is None or crop.size == 0:
        return ""

    height, width = crop.shape[:2]
    py = max(2, int(height * 0.06))
    px = max(2, int(width * 0.03))

    if height > py * 2 and width > px * 2:
        crop = crop[py:height-py, px:width-px]

    crop = cv2.resize(
        crop,
        None,
        fx=4,
        fy=4,
        interpolation=cv2.INTER_CUBIC
    )

    if cell_type == "date":
        config = "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789./-"
    elif cell_type == "time":
        config = "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789:."
    elif cell_type == "number":
        config = "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789."
    else:
        config = "--oem 3 --psm 6"

    return clean(
        pytesseract.image_to_string(
            crop,
            config=config
        )
    )


def extract_row(image, xs, y1, y2, page_number, row_number):
    cells = [
        image[y1:y2, xs[i]:xs[i + 1]]
        for i in range(8)
    ]

    date = format_date(
        ocr_cell(cells[1], "date")
    )

    if not date:
        return None

    start = parse_time(
        ocr_cell(cells[2], "time")
    )

    end = parse_time(
        ocr_cell(cells[3], "time")
    )

    work_code = ocr_cell(cells[4])
    description = ocr_cell(cells[5])
    engineer = clean(
        ocr_cell(cells[6])
    ).upper()

    regular, overtime = (
        calculate_hours(start, end)
        if start and end
        else (0.0, 0.0)
    )

    return {
        "page": page_number,
        "row": row_number,
        "date": date,
        "engineer": engineer,
        "category": classify_activity(
            work_code,
            description
        ),
        "regular": regular,
        "ot": overtime
    }


def get_engineer_name(names):
    names = [
        clean(name).upper()
        for name in names
        if clean(name)
    ]

    if not names:
        return "Unknown"

    return Counter(names).most_common(1)[0][0]


def create_dataframe(rows):
    engineer = get_engineer_name(
        [row["engineer"] for row in rows]
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

            for column in COLUMNS[2:]:
                grouped[key][column] = 0.0

        regular_column, ot_column = CATEGORY_COLUMNS.get(
            row["category"],
            CATEGORY_COLUMNS["Other"]
        )

        grouped[key][regular_column] += row["regular"]
        grouped[key][ot_column] += row["ot"]

    df = pd.DataFrame(
        list(grouped.values()),
        columns=COLUMNS
    )

    df["_date_sort"] = pd.to_datetime(
        df["Date"],
        format="%d.%m.%Y"
    )

    df = df.sort_values(
        "_date_sort",
        kind="stable"
    ).drop(
        columns="_date_sort"
    ).reset_index(drop=True)

    for column in COLUMNS[2:]:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        ).fillna(0).round(2)

    return df


def format_excel(output_file):
    workbook = load_workbook(output_file)

    for sheet in list(workbook.sheetnames):
        if sheet != "Detailed Timesheet":
            del workbook[sheet]

    worksheet = workbook["Detailed Timesheet"]

    thin = Side(style="thin")
    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = border

    worksheet.row_dimensions[1].height = 42
    worksheet.freeze_panes = "C2"

    widths = [
        24, 14, 16, 18, 16, 18, 16, 18, 20,
        22, 20, 22, 17, 19, 17, 19, 15, 17
    ]

    for index, width in enumerate(widths, 1):
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

            if cell.row > 1 and cell.column >= 3:
                cell.number_format = "0.00"

    workbook.save(output_file)


def process_pdf(pdf_file, output_file):
    print("=" * 60)
    print("TIMESHEET PROCESSOR")
    print("=" * 60)
    print(f"Input: {pdf_file}")
    print(f"Rendering at {DPI} DPI...")

    pages = convert_from_path(
        pdf_file,
        dpi=DPI,
        fmt="png",
        thread_count=2
    )

    print(f"Total pages: {len(pages)}")

    all_rows = []

    for page_number, page in enumerate(
        pages,
        start=1
    ):
        print(
            f"\nProcessing page "
            f"{page_number}/{len(pages)}..."
        )

        image = prepare_page(page)

        xs = column_boundaries(
            image.shape[1]
        )

        dates = find_dates(
            image,
            xs[1],
            xs[2]
        )

        print(
            f"Date rows detected: "
            f"{len(dates)}"
        )

        if not dates:
            print("No dated rows found.")
            continue

        centers = [
            item[0]
            for item in dates
        ]

        boundaries = [
            max(0, centers[0] - 35)
        ]

        boundaries += [
            (centers[i] + centers[i + 1]) // 2
            for i in range(len(centers) - 1)
        ]

        boundaries.append(
            min(
                image.shape[0],
                centers[-1] + 35
            )
        )

        for row_number in range(
            len(dates)
        ):
            row = extract_row(
                image,
                xs,
                boundaries[row_number],
                boundaries[row_number + 1],
                page_number,
                row_number + 1
            )

            if row:
                all_rows.append(row)

                print(
                    f"  {row['date']} | "
                    f"{row['category']} | "
                    f"{row['regular']:.2f}h + "
                    f"{row['ot']:.2f}h OT"
                )

    if not all_rows:
        raise RuntimeError(
            "No usable timesheet rows were extracted."
        )

    df = create_dataframe(
        all_rows
    )

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:
        df.to_excel(
            writer,
            sheet_name="Detailed Timesheet",
            index=False
        )

    format_excel(
        output_file
    )

    print("\n" + "=" * 60)
    print("PROCESSING COMPLETE")
    print("=" * 60)
    print(
        f"Rows extracted: "
        f"{len(all_rows)}"
    )
    print(
        f"Dates in Excel: "
        f"{len(df)}"
    )
    print(
        f"Output: "
        f"{output_file}"
    )


def main():
    parser = argparse.ArgumentParser(
        description="Convert scanned timesheet PDF to Excel."
    )

    parser.add_argument(
        "pdf",
        help="Path to the scanned timesheet PDF"
    )

    parser.add_argument(
        "-o",
        "--output",
        default="Detailed_Timesheet.xlsx",
        help="Output Excel filename"
    )

    args = parser.parse_args()

    pdf = Path(
        args.pdf
    )

    if not pdf.exists():
        raise FileNotFoundError(
            f"PDF not found: {pdf}"
        )

    process_pdf(
        str(pdf),
        args.output
    )


if __name__ == "__main__":
    main()
