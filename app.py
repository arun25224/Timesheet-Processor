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
import logging

from paddleocr import PaddleOCR

# Suppress noisy PaddleOCR logs
logging.getLogger("ppocr").setLevel(logging.ERROR)

@st.cache_resource
def load_ocr_model():
    """Bulletproof OCR initialization that adapts to different PaddleOCR versions."""
    try:
        return PaddleOCR(use_angle_cls=True, lang='en', show_log=False)
    except TypeError:
        try:
            return PaddleOCR(use_angle_cls=True, lang='en')
        except TypeError:
            return PaddleOCR(lang='en')

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
# OPENCV GRID DETECTION (Tailored for your specific timesheet)
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
    return re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9 .'-]",
