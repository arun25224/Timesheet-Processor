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

def extract_totals_from_timesheet(df):
    """Extract total values from the Total row in the timesheet"""
    df.columns = df.columns.str.strip()
    
    # Find the Total row
    total_row = None
    for idx, row in df.iterrows():
        for val in row.values:
            if isinstance(val, str) and val.lower().strip() == 'total':
                total_row = row
                break
        if total_row is not None:
            break
    
    if total_row is None:
        st.warning("No 'Total' row found. Please ensure your timesheet has a Total row.")
        return {}
    
    # Extract values from Total row with column mapping
    totals = {}
    
    # Column mapping for different possible column names
    col_mapping = {
        'Travel': ['Travel', 'Travel Time'],
        'NT': ['NT', 'Normal Time'],
        'OT': ['OT', 'Overtime'],
        'Waiting time': ['Waiting time', 'Waiting Time'],
        'Waiting time OT': ['Waiting time OT', 'Waiting Time OT'],
        'Preparation': ['Preparation', 'Preparation Time'],
        'L.Trpt': ['L.Trpt', 'Local transport']
    }
    
    for std_name, possible_names in col_mapping.items():
        value = 0
        for name in possible_names:
            matching_col = None
            for col in df.columns:
                if col.lower() == name.lower():
                    matching_col = col
                    break
            
            if matching_col is not None:
                try:
                    val = total_row[matching_col]
                    if pd.notna(val):
                        value = float(val)
                        break
                except (ValueError, TypeError):
                    pass
        
        totals[std_name] = value
    
    return totals

def process_invoice_logic(
    df, template_file, 
    cust_name, inv_address, del_address, reference, cust_po, 
    proj_no, svc_type, vessel_name, vessel_no, engineer_name, 
    include_admin_fee, position, currency, user_expenses
):
    # Extract totals from timesheet
    totals = extract_totals_from_timesheet(df)
    
    st.write(" Extracted totals from timesheet:", totals)
    
    # Calculate sums
    travel_sum = totals.get('Travel', 0)
    nt_sum = totals.get('NT', 0)
    ot_sum = totals.get('OT', 0)
    waiting_sum = totals.get('Waiting time', 0) + totals.get('Waiting time OT', 0)
    prep_sum = totals.get('Preparation', 0)
    l_trpt_sum = totals.get('L.Trpt', 0)
    
    # Load template
    if template_file.name.lower().endswith('.csv'):
        raise ValueError("The Invoice Template must be an Excel file (.xlsx).")
        
    wb = load_workbook(template_file)
    
    # Determine Target Sheet
    sheet_map = {"SG": "SG", "CN": "CN", "KR": "KR", "EUR": "EUR", "USD": "USD"}
    target_sheet = sheet_map.get(currency, currency)
    
    if target_sheet not in wb.sheetnames:
        matched = False
        for s in wb.sheetnames:
            if target_sheet.lower() in s.lower() or s.lower() in target_sheet.lower():
                target_sheet = s
                matched = True
                break
        if not matched:
            target_sheet = wb.sheetnames[0]
            
    ws = wb[target_sheet]
    
    # Inject Customer Information
    safe_write(ws, 7, 3, cust_name)
    safe_write(ws, 8, 3, inv_address)
    safe_write(ws, 9, 3, del_address)
    safe_write(ws, 10, 3, reference)
    safe_write(ws, 11, 3, cust_po)
    safe_write(ws, 12, 3, proj_no)
    safe_write(ws, 13, 3, svc_type)
    safe_write(ws, 14, 3, vessel_name)
    safe_write(ws, 15, 3, vessel_no)
    
    # Map positions to base rows
    role_base_row_map = {
        "Service Technician": 20,
        "Service Engineer": 30,
        "Senior Service Engineer": 40,
        "Specialist Service Engineer": 50
    }
    
    base_row = role_base_row_map.get(position)
    
    if base_row:
        # Write hours to Column C (index 3) - Hours/#days column
        if travel_sum > 0:
            ws.cell(row=base_row + 1, column=3).value = travel_sum
        if nt_sum > 0:
            ws.cell(row=base_row + 2, column=3).value = nt_sum
        if ot_sum > 0:
            ws.cell(row=base_row + 3, column=3).value = ot_sum
        if waiting_sum > 0:
            ws.cell(row=base_row + 4, column=3).value = waiting_sum
        if prep_sum > 0:
            ws.cell(row=base_row + 5, column=3).value = prep_sum
    
    # Handle Local Transport in Expenses section
    if l_trpt_sum > 0:
        for r in range(1, ws.max_row + 1):
            cell_val = str(ws.cell(row=r, column=2).value or "").lower()
            if "local transport" in cell_val:
                ws.cell(row=r, column=3).value = l_trpt_sum
                break
    
    # Inject User Custom Expenses dynamically
    if user_expenses:
        expense_header_row = None
        for r in range(50, 80):
            val = str(ws.cell(row=r, column=2).value or "").strip().lower()
            if "expenses" in val:
                expense_header_row = r
                break
        
        if expense_header_row:
            safe_write(ws, expense_header_row + 2, 3, engineer_name)
            
            expense_queue = user_expenses.copy()
            
            for r in range(expense_header_row + 1, expense_header_row + 25):
                if not expense_queue:
                    break
                
                c2_val = str(ws.cell(row=r, column=2).value or "").strip()
                c3_val = str(ws.cell(row=r, column=3).value or "").strip()
                
                if "ADD RELEVANT EXPENSES" in c2_val or "ADD RELEVANT EXPENSES" in c3_val:
                    exp_to_inject = expense_queue.pop(0)
                    
                    if "ADD RELEVANT EXPENSES" in c2_val:
                        safe_write(ws, r, 2, exp_to_inject['desc'])
                    else:
                        safe_write(ws, r, 3, exp_to_inject['desc'])
                        
                    if exp_to_inject['qty'] > 0:
                        safe_write(ws, r, 3, exp_to_inject['qty'])
                    if exp_to_inject['price'] > 0:
                        safe_write(ws, r, 5, exp_to_inject['price'])
    
    # Export Final Invoice
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
# TAB 1: SINGLE TIMESHEET UPLOAD
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
# TAB 2: STANDALONE TIMESHEET UPLOAD
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
