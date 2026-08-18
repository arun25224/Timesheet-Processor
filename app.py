import streamlit as st
import os
import re
import cv2
import numpy as np
import pandas as pd
from collections import Counter
from datetime import datetime
from PIL import Image
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
import pytesseract
from pytesseract import Output
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import io
import gc

# ============================================================
# SETTINGS & UTILS
# ============================================================
st.set_page_config(page_title="Timesheet & Invoice Processor", layout="wide")
DPI = 300
NORMAL_START = "08:00"
NORMAL_END = "16:00"

FINAL_COLUMNS = [
    "Engineer Name", "Date", "Travel Time", "Travel OT Time", 
    "Working Time", "Working OT Time", "Waiting Time", "Waiting OT Time", 
    "Preparation Time", "Preparation OT Time", "Maintenance Time", 
    "Maintenance OT Time", "Meeting Time", "Meeting OT Time", 
    "Training Time", "Training OT Time", "Other Time", "Other OT Time"
]

def clean_text(value):
    if value is None or pd.isna(value): return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ").replace("\r", " ").replace("\t", " ")).strip()

def search_text(value):
    return clean_text(value).lower().replace("—", "-").replace("–", "-")

def parse_date(value):
    if not value: return None
    text = clean_text(value).replace("O", "0").replace("o", "0").replace("I", "1").replace("l", "1")
    match = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b", text)
    if not match: return None
    try: return datetime(int(match.group(3)), int(match.group(2)), int(match.group(1)))
    except: return None

def format_date(value):
    parsed = parse_date(value)
    return parsed.strftime("%d.%m.%Y") if parsed else ""

def normalise_time(value):
    if not value: return ""
    text = clean_text(value).replace("O", "0").replace("o", "0").replace("I", "1").replace("l", "1").replace(".", ":")
    match = re.search(r"\b([0-2]?\d):([0-5]\d)\b", text)
    if match:
        h, m = int(match.group(1)), int(match.group(2))
        if 0 <= h <= 23 and 0 <= m <= 59: return f"{h:02d}:{m:02d}"
    match = re.search(r"\b([0-2]\d)([0-5]\d)\b", text)
    if match:
        h, m = int(match.group(1)), int(match.group(2))
        if 0 <= h <= 23 and 0 <= m <= 59: return f"{h:02d}:{m:02d}"
    return ""

def time_to_minutes(value):
    val = normalise_time(value)
    if not val: return None
    try:
        h, m = map(int, val.split(":"))
        return h * 60 + m
    except: return None

def calculate_regular_ot(start_time, end_time):
    start, end = time_to_minutes(start_time), time_to_minutes(end_time)
    if start is None or end is None: return (0.0, 0.0)
    if end < start: end += 24 * 60
    n_start, n_end = time_to_minutes(NORMAL_START), time_to_minutes(NORMAL_END)
    if n_start is None or n_end is None: return (0.0, round((end - start) / 60, 2))
    
    reg, ot = 0, 0
    for minute in range(start, end):
        current = minute % (24 * 60)
        if n_start <= current < n_end: reg += 1
        else: ot += 1
    return (round(reg / 60, 2), round(ot / 60, 2))

CATEGORY_KEYWORDS = {
    "Travel": ["travel", "journey", "transit", "transport", "transfer"],
    "Waiting": ["waiting", "wait", "standby", "stand by"],
    "Preparation": ["preparation", "prepare", "prep", "setup", "set up", "mobilisation"],
    "Maintenance": ["maintenance", "maint", "servicing", "service"],
    "Meeting": ["meeting", "discussion", "briefing"],
    "Training": ["training", "course", "induction"],
    "Working": ["working", "work", "repair", "installation", "install", "overhaul", "operation", "job", "site work"]
}

def classify_activity(work_code, description):
    wc_clean, desc_clean = search_text(work_code), search_text(description)
    combined = wc_clean + " " + desc_clean
    for cat in ["travel", "waiting", "preparation", "maintenance", "meeting", "training"]:
        if cat in wc_clean: return cat.capitalize()
    if "working" in wc_clean or "work hours" in wc_clean: return "Working"
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(kw in combined for kw in keywords): return category
    return "Other"

# ============================================================
# OCR ENGINE
# ============================================================
def prepare_page(pil_image):
    img = np.array(pil_image)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY) if len(img.shape) == 3 else img
    gray = cv2.resize(gray, None, fx=1.10, fy=1.10, interpolation=cv2.INTER_CUBIC)
    return cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(gray)

def cluster_values(values, tolerance=5):
    if not values: return []
    values = sorted([int(v) for v in values])
    clusters, current = [], [values[0]]
    for val in values[1:]:
        if abs(val - np.mean(current)) <= tolerance: current.append(val)
        else:
            clusters.append(current)
            current = [val]
    clusters.append(current)
    return [int(round(np.mean(c))) for c in clusters]

def detect_vertical_lines(image):
    height, width = image.shape[:2]
    edges = cv2.Canny(image, 40, 150, apertureSize=3)
    lines = cv2.HoughLinesP(edges, 1, np.pi / 180, max(50, int(height * 0.03)), minLineLength=max(100, int(height * 0.30)), maxLineGap=30)
    candidates = []
    if lines is not None:
        lines = np.asarray(lines)
        if lines.ndim == 3: lines = lines[:, 0, :] 
        for line in lines:
            if len(line) == 4:
                x1, y1, x2, y2 = map(int, line)
                if abs(x2 - x1) <= 5 and abs(y2 - y1) >= height * 0.30:
                    candidates.append(int(round((x1 + x2) / 2)))
    candidates = cluster_values(candidates, tolerance=8)
    filtered = []
    min_space = max(20, int(width * 0.01))
    for x in candidates:
        if not filtered or x - filtered[-1] >= min_space: filtered.append(x)
    expected_ratios = [0.000, 0.043, 0.205, 0.325, 0.445, 0.610, 0.772, 0.934, 1.000]
    if len(filtered) == 9: return filtered
    left = filtered[0] if len(filtered) >= 2 else int(width * 0.015)
    right = filtered[-1] if len(filtered) >= 2 else int(width * 0.985)
    return [int(round(left + r * (right - left))) for r in expected_ratios]

def detect_horizontal_lines(image):
    height, width = image.shape[:2]
    edges = cv2.Canny(image, 40, 150, apertureSize=3)
    lines = cv2.HoughLinesP(edges, 1, np.pi / 180, max(40, int(width * 0.05)), minLineLength=max(100, int(width * 0.15)), maxLineGap=30)
    candidates = []
    if lines is not None:
        lines = np.asarray(lines)
        if lines.ndim == 3: lines = lines[:, 0, :]
        for line in lines:
            if len(line) == 4:
                x1, y1, x2, y2 = map(int, line)
                if abs(y2 - y1) <= 5 and abs(x2 - x1) >= width * 0.15:
                    candidates.append(int(round((y1 + y2) / 2)))
    return cluster_values(candidates, tolerance=6)

def detect_date_rows(image, x_left, x_right):
    top, bottom = int(image.shape[0] * 0.03), int(image.shape[0] * 0.98)
    crop = image[top:bottom, x_left:x_right]
    if crop is None or crop.shape[0] < 5 or crop.shape[1] < 5: return []
    processed = cv2.resize(crop, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)
    processed = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(processed)
    
    results = []
    configs = ["--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789./-", "--oem 3 --psm 11 -c tessedit_char_whitelist=0123456789./-"]
    for config in configs:
        try:
            data = pytesseract.image_to_data(processed, config=config, output_type=Output.DATAFRAME)
            if data is not None and not data.empty:
                for _, item in data.dropna(subset=["text"]).iterrows():
                    text = clean_text(item["text"])
                    parsed = parse_date(text)
                    if parsed:
                        center_y = top + int(item["top"] / 4) + max(1, int(item["height"] / 4)) // 2
                        results.append({"date": parsed, "date_text": format_date(text), "center_y": center_y, "confidence": float(item.get("conf", 0))})
        except: continue
    unique = []
    for item in sorted(results, key=lambda x: x["center_y"]):
        dup_idx = next((i for i, ex in enumerate(unique) if abs(item["center_y"] - ex["center_y"]) <= 12), None)
        if dup_idx is None: unique.append(item)
        elif item["confidence"] > unique[dup_idx]["confidence"]: unique[dup_idx] = item
    return sorted(unique, key=lambda x: x["center_y"])

def create_row_boundaries(date_rows, horizontal_lines, image_height):
    if not date_rows: return []
    centers = [int(r["center_y"]) for r in date_rows]
    boundaries = []
    above = [y for y in horizontal_lines if y < centers[0]]
    boundaries.append(max(above) if above else max(0, centers[0] - 30))
    for i in range(len(centers) - 1):
        midpoint = int(round((centers[i] + centers[i + 1]) / 2))
        nearby = [y for y in horizontal_lines if abs(y - midpoint) <= 35]
        boundaries.append(min(nearby, key=lambda y: abs(y - midpoint)) if nearby else midpoint)
    below = [y for y in horizontal_lines if y > centers[-1]]
    boundaries.append(min(below) if below else min(image_height - 1, centers[-1] + 30))
    cleaned = []
    for b in map(int, boundaries):
        if not cleaned or b > cleaned[-1] + 2: cleaned.append(b)
        else: cleaned[-1] = max(cleaned[-1], b)
    if len(cleaned) != len(date_rows) + 1:
        cleaned = [max(0, centers[0] - 30)] + [int((centers[i] + centers[i+1])/2) for i in range(len(centers)-1)] + [min(image_height - 1, centers[-1] + 30)]
    return cleaned

def ocr_cell(crop, cell_type):
    if crop is None or crop.shape[0] < 5 or crop.shape[1] < 5: return ""
    pad_y, pad_x = max(2, int(crop.shape[0] * 0.08)), max(2, int(crop.shape[1] * 0.04))
    if crop.shape[0] - 2*pad_y < 5 or crop.shape[1] - 2*pad_x < 5: return ""
    prep = cv2.resize(crop[pad_y:-pad_y, pad_x:-pad_x], None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)
    prep = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(prep)
    configs = {
        "date": ["--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789./-", "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789./-"],
        "time": ["--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789:.", "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789:."],
        "number": ["--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789."],
    }.get(cell_type, ["--oem 3 --psm 7", "--oem 3 --psm 6"])
    results = []
    for config in configs:
        try:
            text = clean_text(pytesseract.image_to_string(prep, config=config))
            if text: results.append(text)
        except: pass
    if not results: return ""
    return results[0] if cell_type in ["date", "time", "number"] else max(results, key=len)

def extract_row(image, x_bounds, y_top, y_bottom):
    cells = []
    for i in range(8):
        x1, x2 = max(0, x_bounds[i]), min(image.shape[1], x_bounds[i+1])
        y1, y2 = max(0, y_top), min(image.shape[0], y_bottom)
        cells.append(image[y1:y2, x1:x2] if x2 > x1 and y2 > y1 else None)
        
    parsed_date = parse_date(ocr_cell(cells[1], "date"))
    if not parsed_date: return None
    date_clean = parsed_date.strftime("%d.%m.%Y")
    
    start_clean = normalise_time(ocr_cell(cells[2], "time"))
    end_clean = normalise_time(ocr_cell(cells[3], "time"))
    eng_clean = clean_text(ocr_cell(cells[6], "name")).strip().upper()
    wc_clean = clean_text(ocr_cell(cells[4], "text"))
    desc_clean = clean_text(ocr_cell(cells[5], "text"))
    
    return {
        "Engineer Name": eng_clean, 
        "Date": date_clean, 
        "Start Time": start_clean, 
        "End Time": end_clean, 
        "Work Code": wc_clean, 
        "Description": desc_clean
    }

def apply_total_row(ws, sum_min_col, sum_max_col, start_row, end_row, table_max_col):
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    total_row = end_row + 1
    ws.cell(row=total_row, column=1, value="Total").font = Font(bold=True)
    for c in range(1, table_max_col + 1):
        cell = ws.cell(row=total_row, column=c)
        cell.border = border
        if sum_min_col <= c <= sum_max_col:
            col_letter = get_column_letter(c)
            cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"
            cell.font = Font(bold=True)
            cell.number_format = "0.00"

# ============================================================
# STREAMLIT UI
# ============================================================
st.title("Timesheet & Invoice Automation")

tab1, tab2 = st.tabs(["Timesheet Extraction", "Invoice Generation"])

with tab1:
    st.header("Upload your timesheet")
    st.write("The system will extract the data, and you can edit any mistakes directly in the grid before generating the final Excel file.")

    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
    
    if "extracted_data" not in st.session_state:
        st.session_state.extracted_data = None

    if uploaded_file is not None and st.session_state.extracted_data is None:
        if st.button("Run Initial Extraction"):
            with st.spinner('Running OCR extraction...'):
                try:
                    pdf_bytes = uploaded_file.read()
                    pdf_info = pdfinfo_from_bytes(pdf_bytes)
                    total_pages = pdf_info["Pages"]
                    
                    raw_rows = []
                    progress_bar = st.progress(0)
                    
                    for page_num in range(1, total_pages + 1):
                        page_images = convert_from_bytes(pdf_bytes, dpi=DPI, fmt="png", thread_count=1, first_page=page_num, last_page=page_num)
                        image = prepare_page(page_images[0])
                        h, w = image.shape[:2]
                        
                        x_bounds = detect_vertical_lines(image)
                        date_rows = detect_date_rows(image, x_bounds[1], x_bounds[2])
                        row_bounds = create_row_boundaries(date_rows, detect_horizontal_lines(image), h)
                        
                        for row_idx in range(len(date_rows)):
                            extracted = extract_row(image, x_bounds, row_bounds[row_idx], row_bounds[row_idx+1])
                            if extracted:
                                raw_rows.append(extracted)
                                
                        progress_bar.progress(page_num / total_pages)
                        del page_images, image
                        gc.collect()
                        
                    if raw_rows:
                        st.session_state.extracted_data = pd.DataFrame(raw_rows)
                        st.rerun()
                    else:
                        st.error("No data extracted. Ensure scan is clear.")
                except Exception as e:
                    st.error(f"Error: {e}")

    # Editable Grid Logic
    if st.session_state.extracted_data is not None:
        st.success("Initial extraction complete! Please review and correct any errors in the grid below:")
        
        # Display editable dataframe
        edited_df = st.data_editor(st.session_state.extracted_data, num_rows="dynamic", use_container_width=True)
        
        if st.button("Generate Final Excel with Verified Data"):
            try:
                aggregated = {}
                canonical_engineer = edited_df["Engineer Name"].mode()[0] if not edited_df.empty else "Unknown"
                
                for idx, row in edited_df.iterrows():
                    eng = str(row["Engineer Name"])
                    dt = str(row["Date"])
                    st_time = str(row.get("Start Time", ""))
                    ed_time = str(row.get("End Time", ""))
                    wc = str(row.get("Work Code", ""))
                    desc = str(row.get("Description", ""))
                    
                    cat = classify_activity(wc, desc)
                    
                    # Manual calc based on verified inputs
                    is_weekend = False
                    parsed_d = parse_date(dt)
                    if parsed_d and parsed_d.weekday() >= 5: is_weekend = True
                    
                    reg, ot = 0.0, 0.0
                    if st_time and ed_time and st_time != "nan" and ed_time != "nan":
                        r, o = calculate_regular_ot(st_time, ed_time)
                        if is_weekend:
                            ot = r + o
                            reg = 0.0
                        else:
                            reg, ot = r, o
                    
                    key = (eng, dt)
                    if key not in aggregated:
                        aggregated[key] = {"Engineer Name": eng, "Date": dt}
                        for col in FINAL_COLUMNS[2:]: aggregated[key][col] = 0.0
                        
                    mapping = {"Travel": ("Travel Time", "Travel OT Time"), "Working": ("Working Time", "Working OT Time"),
                               "Waiting": ("Waiting Time", "Waiting OT Time"), "Preparation": ("Preparation Time", "Preparation OT Time"),
                               "Maintenance": ("Maintenance Time", "Maintenance OT Time"), "Meeting": ("Meeting Time", "Meeting OT Time"),
                               "Training": ("Training Time", "Training OT Time"), "Other": ("Other Time", "Other OT Time")}
                    r_col, ot_col = mapping.get(cat, mapping["Other"])
                    
                    aggregated[key][r_col] += float(reg)
                    aggregated[key][ot_col] += float(ot)

                final_df = pd.DataFrame(list(aggregated.values()))
                final_df["_actual_date"] = final_df["Date"].apply(parse_date)
                final_df = final_df.sort_values(by=["_actual_date"]).reset_index(drop=True)
                
                # Continuous Date filling
                min_date = final_df["_actual_date"].min()
                max_date = final_df["_actual_date"].max()
                dates_df = pd.DataFrame({"_actual_date": pd.date_range(start=min_date, end=max_date, freq='D')})
                dates_df["Date"] = dates_df["_actual_date"].dt.strftime("%d.%m.%Y")
                dates_df["Date_formatted"] = dates_df["_actual_date"].dt.strftime("%d-%b")
                dates_df["Day"] = dates_df["_actual_date"].dt.strftime("%a")
                
                temp_df = pd.merge(dates_df, final_df.drop(columns=["_actual_date", "Engineer Name"], errors="ignore"), on="Date", how="left").fillna(0.0)
                
                wait_times = temp_df["Waiting Time"] + temp_df["Waiting OT Time"]
                l_trpt = []
                for i in range(len(temp_df)):
                    w = wait_times[i]
                    act = sum([temp_df.loc[i, c] for c in ["Travel Time", "Travel OT Time", "Working Time", "Working OT Time", "Preparation Time", "Preparation OT Time"]])
                    if act > 0 and w == 0: l_trpt.append(1)
                    else: l_trpt.append("")
                
                client_df = pd.DataFrame({
                    "Date": temp_df["Date_formatted"], "Day": temp_df["Day"], "PH": "",
                    "Travel": temp_df["Travel Time"] + temp_df["Travel OT Time"], "NT": temp_df["Working Time"], "OT": temp_df["Working OT Time"],
                    "Waiting time": wait_times, "Preparation": temp_df["Preparation Time"] + temp_df["Preparation OT Time"],
                    "L.Trpt": l_trpt, "Remark": ""
                })
                engineer_df = pd.DataFrame({
                    "Date": temp_df["Date_formatted"], "Day": temp_df["Day"], "PH": "",
                    "Travel": temp_df["Travel Time"], "Travel OT": temp_df["Travel OT Time"], "Normal Time": temp_df["Working Time"],
                    "OT": temp_df["Working OT Time"], "Preparation": temp_df["Preparation Time"] + temp_df["Preparation OT Time"], "Remark": ""
                })
                
                for df in [client_df, engineer_df]:
                    for col in df.columns:
                        if df[col].dtype == 'float64': df[col] = df[col].replace(0.0, "")
                
                final_df = final_df.drop(columns=["_actual_date"])
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    final_df.to_excel(writer, sheet_name="Detailed Timesheet", index=False)
                    client_df.to_excel(writer, sheet_name="Client", index=False, startrow=2)
                    engineer_df.to_excel(writer, sheet_name="Engineer", index=False, startrow=2)
                
                output.seek(0)
                workbook = load_workbook(output)
                border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                for sheet, header_cols, is_client in [("Detailed Timesheet", 18, False), ("Client", 10, True), ("Engineer", 9, False)]:
                    ws = workbook[sheet]
                    if sheet != "Detailed Timesheet":
                        for cell in ws[3]:
                            cell.font = Font(bold=True)
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                            cell.border = border
                            if cell.value in ["Travel OT", "OT"]: cell.fill = yellow_fill
                            else: cell.fill = gray_fill
                        min_r = 4
                        if is_client: apply_total_row(ws, 4, 9, 4, ws.max_row, 10)
                        else: apply_total_row(ws, 4, 8, 4, ws.max_row, 9)
                    else:
                        for cell in ws[1]:
                            cell.font = Font(bold=True)
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        min_r = 2
                        apply_total_row(ws, 3, 18, 2, ws.max_row, 18)
                        
                    for row in ws.iter_rows(min_row=min_r, max_row=ws.max_row):
                        for cell in row:
                            cell.border = border
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                
                final_output = io.BytesIO()
                workbook.save(final_output)
                final_output.seek(0)
                
                st.download_button("Download Verified Excel", data=final_output, file_name=f"Verified_Timesheet_{canonical_engineer}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
            except Exception as e:
                st.error(f"Error compiling Excel: {e}")
                
        if st.button("Reset / Upload New File"):
            st.session_state.extracted_data = None
            st.rerun()

# ------------------------------------------------------------
# TAB 2: INVOICE GENERATION
# ------------------------------------------------------------
with tab2:
    st.header("Final Invoice Generation")
    st.markdown("### 1. Upload Required Files")
    col1, col2 = st.columns(2)
    with col1:
        timesheet_excel = st.file_uploader("Upload Processed Timesheet (Excel)", type=["xlsx"], key="ts_upload")
    with col2:
        template_excel = st.file_uploader("Upload Blank Invoice Template (Excel)", type=["xlsx"], key="inv_upload")
        
    st.markdown("### 2. Enter Information")
    c1, c2 = st.columns(2)
    with c1:
        cust_name = st.text_input("Customer name")
        inv_address = st.text_input("Invoicing address")
        del_address = st.text_input("Delivery address")
        reference = st.text_input("Reference")
        cust_po = st.text_input("Customer PO")
    with c2:
        proj_no = st.text_input("Project No")
        svc_type = st.text_input("Service Type")
        vessel_name = st.text_input("Vessel Name")
        vessel_no = st.text_input("Vessel No (if applicable)")
        engineer_name_invoice = st.text_input("Engineer Name (For Expenses)")
        
    st.markdown("### 3. Select Engineer Role")
    position = st.selectbox("Assign Hours to Position:", ["Service Technician", "Service Engineer", "Senior Service Engineer", "Specialist Service Engineer"])
    
    if st.button("Generate Final Invoice", type="primary"):
        if not timesheet_excel or not template_excel:
            st.error("Please upload both the Processed Timesheet AND the Invoice Template first.")
        else:
            try:
                client_df = pd.read_excel(timesheet_excel, sheet_name="Client", skiprows=2)
                travel_sum = pd.to_numeric(client_df['Travel'], errors='coerce').sum()
                nt_sum = pd.to_numeric(client_df['NT'], errors='coerce').sum()
                ot_sum = pd.to_numeric(client_df['OT'], errors='coerce').sum()
                waiting_sum = pd.to_numeric(client_df['Waiting time'], errors='coerce').sum()
                prep_sum = pd.to_numeric(client_df['Preparation'], errors='coerce').sum()
                
                wb = load_workbook(template_excel)
                ws = wb["SG"]
                
                ws["C7"], ws["C8"], ws["C9"], ws["C10"], ws["C11"] = cust_name, inv_address, del_address, reference, cust_po
                ws["C12"], ws["C13"], ws["C14"], ws["C15"] = proj_no, svc_type, vessel_name, vessel_no
                
                r_offset = {"Service Technician": 20, "Service Engineer": 30, "Senior Service Engineer": 40, "Specialist Service Engineer": 50}[position]
                    
                ws[f"D{r_offset + 1}"] = travel_sum if travel_sum > 0 else ""
                ws[f"D{r_offset + 2}"] = nt_sum if nt_sum > 0 else ""
                ws[f"D{r_offset + 3}"] = ot_sum if ot_sum > 0 else ""
                ws[f"D{r_offset + 4}"] = waiting_sum if waiting_sum > 0 else ""
                ws[f"D{r_offset + 5}"] = prep_sum if prep_sum > 0 else ""
                
                for row_idx in range(1, 150):
                    if ws.cell(row=row_idx, column=2).value and str(ws.cell(row=row_idx, column=2).value).strip() == "Expenses":
                        ws.cell(row=row_idx + 2, column=3).value = engineer_name_invoice
                        break
                
                invoice_output = io.BytesIO()
                wb.save(invoice_output)
                invoice_output.seek(0)
                
                st.success("Invoice Generated Successfully!")
                st.download_button("Download Final Invoice", data=invoice_output, file_name=f"Invoice_{cust_name or 'Completed'}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
