import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# ============================================================
# UTILITY FUNCTIONS
# ============================================================
def safe_write(ws, row_idx, col_idx, value):
    try:
        ws.cell(row=row_idx, column=col_idx).value = value
    except AttributeError:
        coord = ws.cell(row=row_idx, column=col_idx).coordinate
        for merged_range in ws.merged_cells.ranges:
            if coord in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                break

def render_expense_ui(tab_key):
    if f"expenses_{tab_key}" not in st.session_state:
        st.session_state[f"expenses_{tab_key}"] = []
        
    st.markdown("### 4. Additional Expenses (Optional)")
    
    expense_options = [
        "Shipyard Pass", "Taxi Overseas", "Ferry Fare", "Visa", "Hotel", 
        "Laundry", "Agent Fee", "Excess Baggage Fee", "Airport Tax", 
        "Flight Ticket", "Misc", "Local transport"
    ]
    
    for i, exp in enumerate(st.session_state[f"expenses_{tab_key}"]):
        col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
        with col1:
            exp['desc'] = st.selectbox(
                "Description", 
                options=expense_options, 
                key=f"desc_{tab_key}_{i}",
                index=expense_options.index(exp['desc']) if exp.get('desc') in expense_options else 0
            )
        with col2:
            exp['qty'] = st.number_input("Quantity", min_value=0.0, value=float(exp.get('qty', 0.0)), step=1.0, key=f"qty_{tab_key}_{i}")
        with col3:
            exp['price'] = st.number_input("Price (SGD)", min_value=0.0, value=float(exp.get('price', 0.0)), step=0.01, key=f"price_{tab_key}_{i}")
        with col4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("X", key=f"del_{tab_key}_{i}"):
                st.session_state[f"expenses_{tab_key}"].pop(i)
                st.rerun()
                
    if st.button("+ Add Expense", key=f"add_{tab_key}"):
        st.session_state[f"expenses_{tab_key}"].append({'desc': 'Shipyard Pass', 'qty': 0.0, 'price': 0.0})
        st.rerun()
        
    return st.session_state[f"expenses_{tab_key}"]

# ============================================================
# TIMESHEET EXTRACTION (TOTAL ROW)
# ============================================================
def extract_totals_from_timesheet(df):
    df.columns = [str(c).strip() for c in df.columns]
    
    # Locate the Total row
    total_row = None
    for _, row in df.iterrows():
        for val in row.values:
            if isinstance(val, str) and val.strip().lower() == 'total':
                total_row = row
                break
        if total_row is not None:
            break
    
    col_mapping = {
        'Travel':          ['Travel', 'Travel Time'],
        'NT':              ['NT', 'Normal Time'],
        'OT':              ['OT', 'Overtime'],
        'Waiting time':    ['Waiting time', 'Waiting Time'],
        'Waiting time OT': ['Waiting time OT', 'Waiting Time OT'],
        'Preparation':     ['Preparation', 'Preparation Time'],
        'L.Trpt':          ['L.Trpt', 'Local transport'],
    }
    
    totals = {}
    for std_name, possible_names in col_mapping.items():
        matching_col = next((c for c in df.columns if c.lower() in [p.lower() for p in possible_names]), None)
        if matching_col is None:
            totals[std_name] = 0.0
            continue
        if total_row is not None:
            try:
                val = total_row[matching_col]
                totals[std_name] = float(val) if pd.notna(val) else 0.0
            except (ValueError, TypeError):
                totals[std_name] = 0.0
        else:
            totals[std_name] = pd.to_numeric(df[matching_col], errors='coerce').fillna(0).sum()
    
    return totals

# ============================================================
# HOURS INJECTION (DYNAMIC ROW DETECTION, COLUMN D)
# ============================================================
def inject_hours(ws, position, values):
    """values = {'travel time': x, 'normal time': x, 'overtime': x, 'waiting time': x, 'preparation time': x}"""
    target = position.strip().lower()
    written = []
    
    # 1) Find the role block label row (exact match, avoids 'service engineer' inside 'senior service engineer')
    role_row = None
    for r in range(1, ws.max_row + 1):
        for c in (1, 2, 3):
            v = str(ws.cell(row=r, column=c).value or "").strip().lower()
            if v == target:
                role_row = r
                break
        if role_row:
            break
    
    # 2) Within the block, match each line by its Type text and write hours to Column D (4)
    if role_row:
        for r in range(role_row + 1, min(role_row + 13, ws.max_row + 1)):
            type_val = str(ws.cell(row=r, column=3).value or "").strip().lower()
            if not type_val:
                type_val = str(ws.cell(row=r, column=2).value or "").strip().lower()
            for key, val in values.items():
                if type_val == key and val:
                    ws.cell(row=r, column=4).value = val
                    written.append((r, key, val))
    else:
        # Fallback: absolute rows (header at 20/30/40/50, data starts at header+1)
        base = {"Service Technician": 20, "Service Engineer": 30,
                "Senior Service Engineer": 40, "Specialist Service Engineer": 50}[position]
        order = ["travel time", "normal time", "overtime", "waiting time", "preparation time"]
        for i, key in enumerate(order):
            if values.get(key):
                ws.cell(row=base + 1 + i, column=4).value = values[key]
                written.append((base + 1 + i, key, values[key]))
    
    return written

# ============================================================
# MAIN INVOICE LOGIC
# ============================================================
def process_invoice_logic(
    df, template_file, 
    cust_name, inv_address, del_address, reference, cust_po, 
    proj_no, svc_type, vessel_name, vessel_no, engineer_name, 
    include_admin_fee, position, currency, user_expenses
):
    totals = extract_totals_from_timesheet(df)
    st.write("✅ Extracted totals from timesheet:", totals)
    
    travel_sum  = totals.get('Travel', 0)
    nt_sum      = totals.get('NT', 0)
    ot_sum      = totals.get('OT', 0)
    waiting_sum = totals.get('Waiting time', 0) + totals.get('Waiting time OT', 0)
    prep_sum    = totals.get('Preparation', 0)
    l_trpt_sum  = totals.get('L.Trpt', 0)
    
    if template_file.name.lower().endswith('.csv'):
        raise ValueError("The Invoice Template must be an Excel file (.xlsx).")
        
    wb = load_workbook(template_file)
    
    target_sheet = currency
    if target_sheet not in wb.sheetnames:
        matched = next((s for s in wb.sheetnames if target_sheet.lower() in s.lower()), None)
        target_sheet = matched if matched else wb.sheetnames[0]
    ws = wb[target_sheet]
    
    # --- Customer Information (values in Column C) ---
    safe_write(ws, 7, 3, cust_name)
    safe_write(ws, 8, 3, inv_address)
    safe_write(ws, 9, 3, del_address)
    safe_write(ws, 10, 3, reference)
    safe_write(ws, 11, 3, cust_po)
    safe_write(ws, 12, 3, proj_no)
    safe_write(ws, 13, 3, svc_type)
    safe_write(ws, 14, 3, vessel_name)
    safe_write(ws, 15, 3, vessel_no)
    
    # --- Hours -> Column D, rows matched by Type text ---
    values = {
        "travel time": travel_sum,
        "normal time": nt_sum,
        "overtime": ot_sum,
        "waiting time": waiting_sum,
        "preparation time": prep_sum,
    }
    written = inject_hours(ws, position, values)
    for r, k, v in written:
        st.write(f"✏️ {position} | {k} = {v} → row {r}, Column D")
    
    # --- Expenses section ---
    expense_header_row = None
    for r in range(40, 90):
        for c in (1, 2):
            v = str(ws.cell(row=r, column=c).value or "").strip().lower()
            if v == "expenses":
                expense_header_row = r
                break
        if expense_header_row:
            break
    
    if expense_header_row:
        # Engineer name on the Allowance [Engineer 1] description cell
        safe_write(ws, expense_header_row + 2, 3, engineer_name)
        
        expense_queue = [e.copy() for e in user_expenses if e.get('desc')]
        
        # Pass 1: match UI expenses against pre-listed descriptions (Column C)
        for r in range(expense_header_row + 1, expense_header_row + 25):
            desc_val = str(ws.cell(row=r, column=3).value or "").strip().lower()
            if not desc_val:
                continue
            matched = next((e for e in expense_queue if e['desc'].lower() == desc_val), None)
            if matched:
                if matched['qty'] > 0:
                    safe_write(ws, r, 4, matched['qty'])     # Quantity = Column D
                if matched['price'] > 0:
                    safe_write(ws, r, 6, matched['price'])   # Price = Column F
                expense_queue.remove(matched)
        
        # Pass 2: remaining expenses go into empty placeholder rows
        for r in range(expense_header_row + 1, expense_header_row + 25):
            if not expense_queue:
                break
            cat_val  = str(ws.cell(row=r, column=2).value or "").strip()
            desc_val = str(ws.cell(row=r, column=3).value or "").strip()
            if "ADD RELEVANT EXPENSES" in cat_val.upper() and not desc_val:
                exp = expense_queue.pop(0)
                safe_write(ws, r, 3, exp['desc'])            # Description = Column C
                if exp['qty'] > 0:
                    safe_write(ws, r, 4, exp['qty'])
                if exp['price'] > 0:
                    safe_write(ws, r, 6, exp['price'])
        
        # Pass 3: Local transport quantity from timesheet L.Trpt total
        if l_trpt_sum > 0:
            for r in range(expense_header_row + 1, expense_header_row + 25):
                desc_val = str(ws.cell(row=r, column=3).value or "").strip().lower()
                if "local transport" in desc_val:
                    safe_write(ws, r, 4, l_trpt_sum)         # Quantity = Column D
                    st.write(f"✏️ Local transport qty {l_trpt_sum} → row {r}, Column D")
                    break
    
    invoice_output = io.BytesIO()
    wb.save(invoice_output)
    invoice_output.seek(0)
    return invoice_output

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="Invoice Generator", layout="wide")

st.title("Final Invoice Generation")
st.write("Select your timesheet format and generate the final invoice template.")

tab1, tab2 = st.tabs(["Single Timesheet Upload (Combined Excel)", "Standalone Timesheet Upload (.csv or .xlsx)"])

# ------------------------------------------------------------
# TAB 1
# ------------------------------------------------------------
with tab1:
    st.markdown("### 1. Upload Required Files")
    col1, col2 = st.columns(2)
    with col1:
        timesheet_excel_t1 = st.file_uploader("Upload Processed Timesheet (Excel)", type=["xlsx"], key="ts_upload_t1")
    with col2:
        template_excel_t1 = st.file_uploader("Upload Blank Invoice Template", type=["xlsx"], key="inv_upload_t1")
        
    st.markdown("### 2. Enter Information")
    c1_t1, c2_t1 = st.columns(2)
    with c1_t1:
        work_order_t1 = st.text_input("Work Order Number", value="NeedsConfirmation", key="wo_t1")
        cust_name_t1 = st.text_input("Customer name", key="cust_name_t1")
        inv_address_t1 = st.text_input("Invoicing address", key="inv_addr_t1")
        del_address_t1 = st.text_input("Delivery address", key="del_addr_t1")
        reference_t1 = st.text_input("Reference", key="ref_t1")
    with c2_t1:
        cust_po_t1 = st.text_input("Customer PO", key="po_t1")
        proj_no_t1 = st.text_input("Project No", key="proj_t1")
        svc_type_t1 = st.text_input("Service Type", key="svc_t1")
        vessel_name_t1 = st.text_input("Vessel Name", key="vessel_t1")
        vessel_no_t1 = st.text_input("Vessel No (if applicable)", key="vessel_no_t1")
        engineer_name_invoice_t1 = st.text_input("Engineer Name (For Expenses)", key="eng_name_t1")

    st.markdown("### 3. Service & Role Details")
    c3_t1, c4_t1, c5_t1 = st.columns(3)
    with c3_t1:
        currency_t1 = st.selectbox("Select Currency:", ["SG", "CN", "KR", "EUR", "USD"], key="curr_t1")
    with c4_t1:
        include_admin_fee_t1 = st.radio("Include 10% Admin Fee?", ["Yes", "No"], key="admin_fee_t1")
    with c5_t1:
        position_t1 = st.selectbox("Assign Hours to Position:", [
            "Service Technician", "Service Engineer", "Senior Service Engineer", "Specialist Service Engineer"
        ], key="pos_t1")

    user_expenses_t1 = render_expense_ui("t1")

    if st.button("Generate Final Invoice", type="primary", key="btn_t1"):
        if not timesheet_excel_t1 or not template_excel_t1:
            st.error("Please upload the Processed Timesheet AND the Invoice Template.")
        else:
            try:
                xls = pd.ExcelFile(timesheet_excel_t1)
                sheet_names = xls.sheet_names
                eng_sheet = next((s for s in sheet_names if 'engineer' in s.lower()), sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])
                df_t1 = pd.read_excel(timesheet_excel_t1, sheet_name=eng_sheet)
                
                output = process_invoice_logic(
                    df_t1, template_excel_t1, 
                    cust_name_t1, inv_address_t1, del_address_t1, reference_t1, cust_po_t1, 
                    proj_no_t1, svc_type_t1, vessel_name_t1, vessel_no_t1, engineer_name_invoice_t1, 
                    include_admin_fee_t1, position_t1, currency_t1, user_expenses_t1
                )
                
                st.success("Invoice Generated Successfully")
                st.download_button(
                    label="Download Final Invoice",
                    data=output,
                    file_name=f"Invoice_{cust_name_t1 or 'Completed'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_t1"
                )
            except Exception as e:
                st.error(f"Error occurred while processing: {str(e)}")

# ------------------------------------------------------------
# TAB 2
# ------------------------------------------------------------
with tab2:
    st.markdown("### 1. Upload Required Files")
    col1_t2, col2_t2 = st.columns(2)
    with col1_t2:
        client_timesheet_t2 = st.file_uploader("Upload Client Timesheet", type=["xlsx", "csv"], key="cli_upload_t2")
    with col2_t2:
        template_excel_t2 = st.file_uploader("Upload Blank Invoice Template", type=["xlsx", "csv"], key="inv_upload_t2")
        
    st.markdown("### 2. Enter Information")
    c1_t2, c2_t2 = st.columns(2)
    with c1_t2:
        work_order_t2 = st.text_input("Work Order Number", value="NeedsConfirmation", key="wo_t2")
        cust_name_t2 = st.text_input("Customer name", key="cust_name_t2")
        inv_address_t2 = st.text_input("Invoicing address", key="inv_addr_t2")
        del_address_t2 = st.text_input("Delivery address", key="del_addr_t2")
        reference_t2 = st.text_input("Reference", key="ref_t2")
    with c2_t2:
        cust_po_t2 = st.text_input("Customer PO", key="po_t2")
        proj_no_t2 = st.text_input("Project No", key="proj_t2")
        svc_type_t2 = st.text_input("Service Type", key="svc_t2")
        vessel_name_t2 = st.text_input("Vessel Name", key="vessel_t2")
        vessel_no_t2 = st.text_input("Vessel No (if applicable)", key="vessel_no_t2")
        engineer_name_invoice_t2 = st.text_input("Engineer Name (For Expenses)", key="eng_name_t2")

    st.markdown("### 3. Service & Role Details")
    c3_t2, c4_t2, c5_t2 = st.columns(3)
    with c3_t2:
        currency_t2 = st.selectbox("Select Currency:", ["SG", "CN", "KR", "EUR", "USD"], key="curr_t2")
    with c4_t2:
        include_admin_fee_t2 = st.radio("Include 10% Admin Fee?", ["Yes", "No"], key="admin_fee_t2")
    with c5_t2:
        position_t2 = st.selectbox("Assign Hours to Position:", [
            "Service Technician", "Service Engineer", "Senior Service Engineer", "Specialist Service Engineer"
        ], key="pos_t2")

    user_expenses_t2 = render_expense_ui("t2")

    if st.button("Generate Final Invoice", type="primary", key="btn_t2"):
        if not client_timesheet_t2 or not template_excel_t2:
            st.error("Please upload the Client Timesheet AND the Invoice Template.")
        else:
            try:
                if client_timesheet_t2.name.lower().endswith('.csv'):
                    df_t2 = pd.read_csv(client_timesheet_t2)
                else:
                    df_t2 = pd.read_excel(client_timesheet_t2, sheet_name=0)
                
                output = process_invoice_logic(
                    df_t2, template_excel_t2, 
                    cust_name_t2, inv_address_t2, del_address_t2, reference_t2, cust_po_t2, 
                    proj_no_t2, svc_type_t2, vessel_name_t2, vessel_no_t2, engineer_name_invoice_t2, 
                    include_admin_fee_t2, position_t2, currency_t2, user_expenses_t2
                )
                
                st.success("Invoice Generated Successfully")
                st.download_button(
                    label="Download Final Invoice",
                    data=output,
                    file_name=f"Invoice_{cust_name_t2 or 'Completed'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_t2"
                )
            except Exception as e:
                st.error(f"Error occurred while processing: {str(e)}")
