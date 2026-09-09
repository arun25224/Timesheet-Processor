import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

def safe_write(ws, row_idx, col_idx, value):
    """Safely write to a cell, handling merged cells"""
    try:
        ws.cell(row=row_idx, column=col_idx).value = value
    except Exception as e:
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
    """Extract total values from the timesheet - looking for the Total row"""
    
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Find the Total row (case-insensitive)
    total_row = None
    for idx, row in df.iterrows():
        # Check if any column contains 'Total'
        for val in row.values:
            if isinstance(val, str) and val.lower().strip() == 'total':
                total_row = row
                break
        if total_row is not None:
            break
    
    if total_row is None:
        # If no Total row found, sum all numeric columns
        st.warning("No 'Total' row found. Summing all rows.")
        totals = {}
        
        # Map column names to standard names
        col_mapping = {
            'Travel': 'Travel',
            'NT': 'NT',
            'Normal Time': 'NT',
            'OT': 'OT',
            'Overtime': 'OT',
            'Waiting time': 'Waiting time',
            'Waiting time OT': 'Waiting time OT',
            'Waiting Time OT': 'Waiting time OT',
            'Preparation': 'Preparation',
            'Preparation Time': 'Preparation',
            'L.Trpt': 'L.Trpt',
            'Local transport': 'L.Trpt'
        }
        
        for col in df.columns:
            for key, std_name in col_mapping.items():
                if key.lower() == col.lower():
                    totals[std_name] = pd.to_numeric(df[col], errors='coerce').fillna(0).sum()
                    break
        
        return totals
    
    # Extract values from Total row
    totals = {}
    
    # Map columns with various possible names
    col_mapping = {
        'Travel': ['Travel', 'Travel Time'],
        'NT': ['NT', 'Normal Time', 'Normal time'],
        'OT': ['OT', 'Overtime', 'ot'],
        'Waiting time': ['Waiting time', 'Waiting Time'],
        'Waiting time OT': ['Waiting time OT', 'Waiting Time OT'],
        'Preparation': ['Preparation', 'Preparation Time'],
        'L.Trpt': ['L.Trpt', 'Local transport', 'Local trpt']
    }
    
    for std_name, possible_names in col_mapping.items():
        value = 0
        for name in possible_names:
            # Case-insensitive column search
            matching_col = None
            for col in df.columns:
                if col.lower() == name.lower():
                    matching_col = col
                    break
            
            if matching_col is not None:
                try:
                    val = total_row[matching_col]
                    if pd.notna(val):
                        value += float(val)
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
    
    st.write("Extracted totals:", totals)  # Debug output
    
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
    
    # Map positions to base rows (Row where "Category" header is)
    # Looking at your template:
    # Service technician: Row 20
    # Service engineer: Row 30
    # Senior Service Engineer: Row 40
    # Specialist Service Engineer: Row 50
    
    role_base_row_map = {
        "Service Technician": 20,
        "Service Engineer": 30,
        "Senior Service Engineer": 40,
        "Specialist Service Engineer": 50
    }
    
    base_row = role_base_row_map.get(position)
    
    if base_row:
        # Write hours to Column D (Hours/#days column)
        # Travel Time is at base_row + 1
        # Normal Time is at base_row + 2
        # Overtime is at base_row + 3
        # Waiting Time is at base_row + 4
        # Preparation Time is at base_row + 5
        
        if travel_sum > 0:
            ws.cell(row=base_row + 1, column=4).value = travel_sum
            st.write(f"Writing Travel Time {travel_sum} to row {base_row + 1}, col 4")
        if nt_sum > 0:
            ws.cell(row=base_row + 2, column=4).value = nt_sum
            st.write(f"Writing Normal Time {nt_sum} to row {base_row + 2}, col 4")
        if ot_sum > 0:
            ws.cell(row=base_row + 3, column=4).value = ot_sum
            st.write(f"Writing Overtime {ot_sum} to row {base_row + 3}, col 4")
        if waiting_sum > 0:
            ws.cell(row=base_row + 4, column=4).value = waiting_sum
            st.write(f"Writing Waiting Time {waiting_sum} to row {base_row + 4}, col 4")
        if prep_sum > 0:
            ws.cell(row=base_row + 5, column=4).value = prep_sum
            st.write(f"Writing Preparation Time {prep_sum} to row {base_row + 5}, col 4")
    
    # Handle Local Transport in Expenses section
    if l_trpt_sum > 0:
        # Find the expenses section and local transport row
        for r in range(1, ws.max_row + 1):
            cell_val = str(ws.cell(row=r, column=2).value or "").lower()
            if "local transport" in cell_val:
                ws.cell(row=r, column=3).value = l_trpt_sum  # Quantity column
                st.write(f"Writing Local Transport {l_trpt_sum} to row {r}, col 3")
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
                
                # Check if this is a placeholder row
                if "ADD RELEVANT EXPENSES" in c2_val or "ADD RELEVANT EXPENSES" in c3_val:
                    exp_to_inject = expense_queue.pop(0)
                    
                    # Write description to column 2 or 3 depending on which has placeholder
                    if "ADD RELEVANT EXPENSES" in c2_val:
                        safe_write(ws, r, 2, exp_to_inject['desc'])
                    else:
                        safe_write(ws, r, 3, exp_to_inject['desc'])
                        
                    if exp_to_inject['qty'] > 0:
                        safe_write(ws, r, 3, exp_to_inject['qty'])  # Quantity is column 3
                    if exp_to_inject['price'] > 0:
                        safe_write(ws, r, 5, exp_to_inject['price'])  # Price is column 5
    
    # Export Final Invoice
    invoice_output = io.BytesIO()
    wb.save(invoice_output)
    invoice_output.seek(0)
    
    return invoice_output

# STREAMLIT UI
st.set_page_config(page_title="Invoice Generator", layout="wide")

st.title("Final Invoice Generation")
st.write("Upload your timesheet and generate the invoice.")

# File uploads
col1, col2 = st.columns(2)
with col1:
    timesheet_file = st.file_uploader("Upload Timesheet", type=["xlsx", "csv"], key="ts_upload")
with col2:
    template_file = st.file_uploader("Upload Invoice Template", type=["xlsx"], key="inv_upload")

st.markdown("### Enter Information")
c1, c2 = st.columns(2)
with c1:
    cust_name = st.text_input("Customer name", key="cust_name")
    inv_address = st.text_input("Invoicing address", key="inv_addr")
    del_address = st.text_input("Delivery address", key="del_addr")
    reference = st.text_input("Reference", key="ref")
with c2:
    cust_po = st.text_input("Customer PO", key="po")
    proj_no = st.text_input("Project No", key="proj")
    svc_type = st.text_input("Service Type", key="svc")
    vessel_name = st.text_input("Vessel Name", key="vessel")
    vessel_no = st.text_input("Vessel No (if applicable)", key="vessel_no")
    engineer_name = st.text_input("Engineer Name (For Expenses)", key="eng_name")

st.markdown("### Service & Role Details")
c3, c4, c5 = st.columns(3)
with c3:
    currency = st.selectbox("Select Currency:", ["SG", "CN", "KR", "EUR", "USD"], key="curr")
with c4:
    include_admin_fee = st.radio("Include 10% Admin Fee?", ["Yes", "No"], key="admin_fee")
with c5:
    position = st.selectbox("Assign Hours to Position:", [
        "Service Technician", "Service Engineer", "Senior Service Engineer", "Specialist Service Engineer"
    ], key="pos")

user_expenses = render_expense_ui("main")

if st.button("Generate Final Invoice", type="primary"):
    if not timesheet_file or not template_file:
        st.error("Please upload both the Timesheet AND the Invoice Template.")
    else:
        try:
            # Read timesheet
            if timesheet_file.name.lower().endswith('.csv'):
                df = pd.read_csv(timesheet_file)
            else:
                df = pd.read_excel(timesheet_file, sheet_name=0)
            
            st.write("Timesheet columns:", df.columns.tolist())
            st.write("Timesheet data preview:")
            st.dataframe(df.head())
            
            # Process invoice
            output = process_invoice_logic(
                df, template_file, 
                cust_name, inv_address, del_address, reference, cust_po, 
                proj_no, svc_type, vessel_name, vessel_no, engineer_name, 
                include_admin_fee, position, currency, user_expenses
            )
            
            st.success("Invoice Generated Successfully!")
            st.download_button(
                label="Download Final Invoice",
                data=output,
                file_name=f"Invoice_{cust_name or 'Completed'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl"
            )
        except Exception as e:
            st.error(f"Error occurred while processing: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
