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
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import io
import gc
import pytesseract

# ============================================================
# SETTINGS & CONSTANTS
# ============================================================
st.set_page_config(page_title="Timesheet and Invoice Processor", layout="wide")
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

CATEGORY_KEYWORDS = {
    "Travel": ["travel", "journey", "transit", "transport", "transfer"],
    "Waiting": ["waiting", "wait", "standby", "stand by"],
    "Preparation": ["preparation", "prepare", "prep", "setup", "set up", "mobilisation"],
    "Maintenance": ["maintenance", "maint", "servicing", "service"],
    "Meeting": ["meeting", "discussion", "briefing"],
    "Training": ["training", "course", "induction"],
    "Working": ["working", "work", "repair", "installation", "install", "overhaul", "operation", "job", "site work"]
}

# ============================================================
# UTILITY FUNCTIONS
# ============================================================
def clean_text(value):
    if value is None or pd.isna(value): return ""
    val = str(value).replace("\n", " ").replace("\r", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", val).strip()

def search_text(value):
    return clean_text(value).lower().replace("—", "-").replace("–", "-")

def parse_date(value):
    if not value: return None
    text = clean_text(value).replace("O", "0").replace("o", "0").replace("I", "1").replace("l", "1")
    match = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b", text)
    if not match: return None
    try: return datetime(int(match.group(3)), int(match.group(2)), int(match.group(1)))
    except: return None

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
# OPENCV GRID DETECTION
# ============================================================
def prepare_page(pil_image):
    img = np.array(pil_image)
    if len(img.shape) == 3: gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    else: gray = img
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
        lines = np.asarray(lines).reshape(-1, 4)
        for x1, y1, x2, y2 in lines:
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

def clean_engineer_name(value):
    val = clean_text(value)
    if not val: return ""
    return re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9 .'-]", " ", val)).strip().upper()

def choose_engineer_name(names):
    names = [clean_engineer_name(n) for n in names if clean_engineer_name(n)]
    if not names: return "Unknown"
    from difflib import SequenceMatcher
    groups = []
    for name in names:
        placed = False
        for group in groups:
            if SequenceMatcher(None, name, group[0]).ratio() >= 0.80:
                group.append(name)
                placed = True
                break
        if not placed: groups.append([name])
    groups.sort(key=len, reverse=True)
    return Counter(groups[0]).most_common(1)[0][0]

def empty_final_row(engineer, date):
    row = {"Engineer Name": engineer, "Date": date}
    for col in FINAL_COLUMNS:
        if col not in row: row[col] = 0.0
    return row

def add_hours(row, category, regular, overtime):
    mapping = {
        "Travel": ("Travel Time", "Travel OT Time"), "Working": ("Working Time", "Working OT Time"),
        "Waiting": ("Waiting Time", "Waiting OT Time"), "Preparation": ("Preparation Time", "Preparation OT Time"),
        "Maintenance": ("Maintenance Time", "Maintenance OT Time"), "Meeting": ("Meeting Time", "Meeting OT Time"),
        "Training": ("Training Time", "Training OT Time"), "Other": ("Other Time", "Other OT Time")
    }
    r_col, ot_col = mapping.get(category, mapping["Other"])
    row[r_col] += float(regular)
    row[ot_col] += float(overtime)

def apply_total_row(ws, sum_min_col, sum_max_col, start_row, end_row, table_max_col):
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    total_row = end_row + 1
    first_cell = ws.cell(row=total_row, column=1, value="Total")
    first_cell.font = Font(bold=True)
    first_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    for c in range(1, table_max_col + 1):
        cell = ws.cell(row=total_row, column=c)
        cell.border = border
        if sum_min_col <= c <= sum_max_col:
            col_letter = get_column_letter(c)
            cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "0.00"

# ============================================================
# STREAMLIT UI & TABS
# ============================================================
st.title("Timesheet and Invoice Automation")
tab1, tab2 = st.tabs(["Step 1: Timesheet Extraction", "Step 2: Invoice Generation"])

# ------------------------------------------------------------
# TAB 1: TIMESHEET EXTRACTION (TESSERACT)
# ------------------------------------------------------------
with tab1:
    st.header("Deep Learning Timesheet Processor")
    st.write("Upload a scanned PDF timesheet to extract data across all pages into the Excel format.")

    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

    if uploaded_file is not None:
        if st.button("Extract Data from PDF"):
            with st.spinner('Analyzing image and extracting data across all pages...'):
                try:
                    pdf_bytes = uploaded_file.read()
                    pdf_info = pdfinfo_from_bytes(pdf_bytes)
                    total_pages = pdf_info["Pages"]
                    
                    raw_rows = []
                    engineer_names = []
                    progress_bar = st.progress(0)
                    
                    for page_num in range(1, total_pages + 1):
                        page_images = convert_from_bytes(
                            pdf_bytes, dpi=DPI, fmt="png", thread_count=1, first_page=page_num, last_page=page_num
                        )
                        pil_page = page_images[0]
                        img_rgb = np.array(pil_page)
                        gray = prepare_page(pil_page)
                        h, w = gray.shape[:2]
                        
                        # Detect grid lines using OpenCV
                        x_bounds = detect_vertical_lines(gray)
                        if len(x_bounds) < 8:
                            expected_ratios = [0.000, 0.043, 0.205, 0.325, 0.445, 0.610, 0.772, 0.934, 1.000]
                            x_bounds = [int(w * r) for r in expected_ratios]
                        
                        # OCR with Tesseract
                        data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT)
                        
                        text_boxes = []
                        n_boxes = len(data['text'])
                        for i in range(n_boxes):
                            conf = int(data['conf'][i])
                            text = str(data['text'][i]).strip()
                            # Filter out low confidence and empty strings
                            if conf > 20 and text:
                                left = data['left'][i]
                                top = data['top'][i]
                                width = data['width'][i]
                                height = data['height'][i]
                                cx = left + width / 2
                                cy = top + height / 2
                                text_boxes.append({"text": text, "cx": cx, "cy": cy})
                        
                        # Group text into visual rows based on Y-coordinate
                        text_boxes.sort(key=lambda x: x["cy"])
                        visual_rows = []
                        current_row = []
                        for tb in text_boxes:
                            if not current_row: current_row.append(tb)
                            elif abs(tb["cy"] - current_row[0]["cy"]) <= 20: current_row.append(tb)
                            else:
                                visual_rows.append(current_row)
                                current_row = [tb]
                        if current_row: visual_rows.append(current_row)
                        
                        # Map visual rows to the 8-column grid
                        for row_tbs in visual_rows:
                            cells_text = [""] * 8
                            for tb in row_tbs:
                                for i in range(8):
                                    left_bound = x_bounds[i]
                                    right_bound = x_bounds[i+1] if i+1 < len(x_bounds) else w
                                    if left_bound - 15 <= tb["cx"] <= right_bound + 15:
                                        cells_text[i] += (" " + tb["text"] if cells_text[i] else tb["text"])
                            
                            start_clean = normalise_time(cells_text[2])
                            end_clean = normalise_time(cells_text[3])
                            wc_clean = clean_text(cells_text[4])
                            
                            if not start_clean and not end_clean and not wc_clean: continue
                                
                            parsed_date = parse_date(cells_text[1])
                            if parsed_date: date_clean = parsed_date.strftime("%d.%m.%Y")
                            else: continue
                            
                            is_weekend = datetime.strptime(date_clean, "%d.%m.%Y").weekday() >= 5
                            eng_clean = clean_engineer_name(cells_text[6])
                            if eng_clean: engineer_names.append(eng_clean)
                            desc_clean = clean_text(cells_text[5])
                            
                            reg_hrs, ot_hrs = 0.0, 0.0
                            if start_clean and end_clean:
                                r, o = calculate_regular_ot(start_clean, end_clean)
                                if is_weekend: ot_hrs, reg_hrs = r + o, 0.0
                                else: reg_hrs, ot_hrs = r, o
                                    
                            raw_rows.append({
                                "Engineer Name": eng_clean, "Date": date_clean, "Start Time": start_clean, 
                                "End Time": end_clean, "Work Code": wc_clean, "Short Description": desc_clean, 
                                "Category": classify_activity(wc_clean, desc_clean), "Regular Hours": reg_hrs, "OT Hours": ot_hrs
                            })
                            
                        progress_bar.progress(page_num / total_pages)
                        del page_images, pil_page, gray, img_rgb
                        gc.collect()
                    
                    if not raw_rows:
                        st.error("No usable rows extracted. Please ensure the PDF scan is clear.")
                        st.stop()
                    
                    canonical_engineer = choose_engineer_name(engineer_names)
                    for r in raw_rows: 
                        if not r["Engineer Name"]: r["Engineer Name"] = canonical_engineer
                    
                    aggregated = {}
                    for r in raw_rows:
                        key = (r["Engineer Name"], r["Date"])
                        if key not in aggregated: aggregated[key] = empty_final_row(key[0], key[1])
                        add_hours(aggregated[key], r["Category"], r["Regular Hours"], r["OT Hours"])
                    
                    final_df = pd.DataFrame(list(aggregated.values()))
                    for col in FINAL_COLUMNS:
                        if col not in final_df.columns: final_df[col] = "" if col in ["Engineer Name", "Date"] else 0.0
                    final_df = final_df[FINAL_COLUMNS]
                    
                    final_df["_actual_date"] = final_df["Date"].apply(parse_date)
                    final_df = final_df.sort_values(by=["_actual_date", "Engineer Name"]).reset_index(drop=True)
                    for col in FINAL_COLUMNS[2:]:
                        final_df[col] = pd.to_numeric(final_df[col], errors="coerce").fillna(0).round(2)
                    
                    min_date = final_df["_actual_date"].min()
                    max_date = final_df["_actual_date"].max()
                    full_date_range = pd.date_range(start=min_date, end=max_date, freq='D')
                    dates_df = pd.DataFrame({"_actual_date": full_date_range})
                    dates_df["Date"] = dates_df["_actual_date"].dt.strftime("%d.%m.%Y")
                    dates_df["Date_formatted"] = dates_df["_actual_date"].dt.strftime("%d-%b")
                    dates_df["Day"] = dates_df["_actual_date"].dt.strftime("%a")
                    
                    temp_df = pd.merge(dates_df, final_df.drop(columns=["_actual_date", "Engineer Name"], errors="ignore"), on="Date", how="left")
                    for col in FINAL_COLUMNS[2:]: temp_df[col] = temp_df[col].fillna(0.0)
                        
                    wait_times = temp_df["Waiting Time"] + temp_df["Waiting OT Time"]
                    l_trpt_vals = []
                    for i in range(len(temp_df)):
                        w = wait_times[i]
                        activity_sum = (temp_df.loc[i, "Travel Time"] + temp_df.loc[i, "Travel OT Time"] + 
                                        temp_df.loc[i, "Working Time"] + temp_df.loc[i, "Working OT Time"] + 
                                        temp_df.loc[i, "Preparation Time"] + temp_df.loc[i, "Preparation OT Time"] + w)
                        if activity_sum > 0 and (w == 0 or pd.isna(w)): l_trpt_vals.append(1)
                        else: l_trpt_vals.append("")
                    
                    client_df = pd.DataFrame({
                        "Date": temp_df["Date_formatted"], "Day": temp_df["Day"], "PH": "",
                        "Travel": temp_df["Travel Time"] + temp_df["Travel OT Time"], "NT": temp_df["Working Time"], "OT": temp_df["Working OT Time"],
                        "Waiting time": wait_times, "Preparation": temp_df["Preparation Time"] + temp_df["Preparation OT Time"],
                        "L.Trpt": l_trpt_vals, "Remark": ""
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
                    
                    # Formatting Detailed Timesheet
                    ws_raw = workbook["Detailed Timesheet"]
                    for cell in ws_raw[1]:
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    ws_raw.row_dimensions[1].height = 45
                    ws_raw.freeze_panes = "C2"
                    for row in ws_raw.iter_rows(max_row=ws_raw.max_row):
                        for cell in row:
                            cell.border = border
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                            if cell.column >= 3 and cell.row > 1 and isinstance(cell.value, (int, float)): cell.number_format = "0.00"
                    if ws_raw.max_row >= 2: apply_total_row(ws_raw, 3, 18, 2, ws_raw.max_row, 18)
                    for col, width in {"A":24, "B":14, "C":16, "D":18, "E":16, "F":18, "G":16, "H":18, "I":20, "J":22, "K":20, "L":22, "M":17, "N":19, "O":17, "P":19, "Q":15, "R":17}.items(): 
                        ws_raw.column_dimensions[col].width = width
                    
                    # Formatting Client Sheet
                    ws_client = workbook["Client"]
                    ws_client["A1"] = "Timesheet Calculation (Singapore)"
                    ws_client["A1"].font = Font(size=14, bold=True)
                    ws_client["A1"].alignment = Alignment(horizontal="center", vertical="center")
                    ws_client.merge_cells("A1:J1")
                    ws_client.row_dimensions[1].height = 25
                    for cell in ws_client[3]:
                        cell.font = Font(bold=True)
                        cell.fill = gray_fill
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = border
                    for row in ws_client.iter_rows(min_row=4, max_row=ws_client.max_row):
                        for cell in row:
                            cell.border = border
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                    if ws_client.max_row >= 4: apply_total_row(ws_client, 4, 9, 4, ws_client.max_row, 10)
                    for col, w in {"A":12, "B":10, "C":8, "D":12, "E":12, "F":12, "G":14, "H":14, "I":10, "J":25}.items(): 
                        ws_client.column_dimensions[col].width = w
                    
                    # Formatting Engineer Sheet
                    ws_eng = workbook["Engineer"]
                    ws_eng["A1"] = "Timesheet Calculation (ENGINEER OT)"
                    ws_eng["A1"].font = Font(size=14, bold=True)
                    ws_eng["A1"].alignment = Alignment(horizontal="center", vertical="center")
                    ws_eng.merge_cells("A1:I1")
                    ws_eng.row_dimensions[1].height = 25
                    for cell in ws_eng[3]:
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = border
                        if cell.value in ["Travel OT", "OT"]: cell.fill = yellow_fill
                        else: cell.fill = gray_fill
                    for row in ws_eng.iter_rows(min_row=4, max_row=ws_eng.max_row):
                        for cell in row:
                            cell.border = border
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                    if ws_eng.max_row >= 4: apply_total_row(ws_eng, 4, 8, 4, ws_eng.max_row, 9)
                    for col, w in {"A":12, "B":10, "C":8, "D":12, "E":12, "F":15, "G":12, "H":14, "I":25}.items(): 
                        ws_eng.column_dimensions[col].width = w
                    
                    final_output = io.BytesIO()
                    workbook.save(final_output)
                    final_output.seek(0)
                    
                    st.success("Extraction Complete")
                    st.download_button(
                        label="Download Processed Timesheet",
                        data=final_output,
                        file_name=f"Timesheet_{canonical_engineer}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")

# ------------------------------------------------------------
# TAB 2: INVOICE GENERATION
# ------------------------------------------------------------
with tab2:
    st.header("Final Invoice Generation")
    st.write("Fill in the customer details and generate the final invoice template.")
    
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
    position = st.selectbox("Assign Hours to Position:", [
        "Service Technician", 
        "Service Engineer", 
        "Senior Service Engineer", 
        "Specialist Service Engineer"
    ])
    
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
                l_trpt_sum = pd.to_numeric(client_df['L.Trpt'], errors='coerce').sum()
                
                wb = load_workbook(template_excel)
                if "SG" not in wb.sheetnames:
                    st.error("The uploaded template does not contain an 'SG' tab.")
                    st.stop()
                    
                ws = wb["SG"]
                
                ws["C7"] = cust_name
                ws["C8"] = inv_address
                ws["C9"] = del_address
                ws["C10"] = reference
                ws["C11"] = cust_po
                ws["C12"] = proj_no
                ws["C13"] = svc_type
                ws["C14"] = vessel_name
                ws["C15"] = vessel_no
                
                r_offset = 20
                if position == "Service Engineer": r_offset = 30
                elif position == "Senior Service Engineer": r_offset = 40
                elif position == "Specialist Service Engineer": r_offset = 50
                    
                ws[f"D{r_offset + 1}"] = travel_sum if travel_sum > 0 else ""
                ws[f"D{r_offset + 2}"] = nt_sum if nt_sum > 0 else ""
                ws[f"D{r_offset + 3}"] = ot_sum if ot_sum > 0 else ""
                ws[f"D{r_offset + 4}"] = waiting_sum if waiting_sum > 0 else ""
                ws[f"D{r_offset + 5}"] = prep_sum if prep_sum > 0 else ""
                
                expense_row = None
                local_transport_row = None
                
                for row_idx in range(1, 150):
                    col_b_val = ws.cell(row=row_idx, column=2).value
                    col_c_val = ws.cell(row=row_idx, column=3).value
                    
                    if col_b_val and str(col_b_val).strip() == "Expenses": expense_row = row_idx
                    if col_c_val and "local transport" in str(col_c_val).lower(): local_transport_row = row_idx
                        
                if expense_row:
                    ws.cell(row=expense_row + 2, column=3).value = engineer_name_invoice
                    
                if local_transport_row and l_trpt_sum > 0:
                    ws.cell(row=local_transport_row, column=5).value = l_trpt_sum
                
                invoice_output = io.BytesIO()
                wb.save(invoice_output)
                invoice_output.seek(0)
                
                st.success("Invoice Generated Successfully")
                st.download_button(
                    label="Download Final Invoice",
                    data=invoice_output,
                    file_name=f"Invoice_{cust_name or 'Completed'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred while processing the invoice: {str(e)}")
