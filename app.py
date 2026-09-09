import streamlit as st
import pandas as pd
import io
import zipfile
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
        "Flight Ticket", "Misc"
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
            if st.button("Remove", key=f"del_{tab_key}_{i}"):
                st.session_state[f"expenses_{tab_key}"].pop(i)
                st.rerun()
                
    if st.button("+ Add Expense", key=f"add_{tab_key}"):
        st.session_state[f"expenses_{tab_key}"].append({'desc': 'Shipyard Pass', 'qty': 0.0, 'price': 0.0})
        st.rerun()
        
    return st.session_state[f"expenses_{tab_key}"]

def process_invoice_logic(
    eng_df, client_df, template_file, 
    cust_name, inv_address, del_address, reference, cust_po, 
    proj_no, svc_type, vessel_name, vessel_no, engineer_name, 
    include_admin_fee, position, currency, user_expenses
):
    # Clean column names to remove accidental trailing spaces
    eng_df.columns = eng_df.columns.str.strip()
    client_df.columns = client_df.columns.str.strip()

    # --- PROCESS ENGINEER TIMESHEET (Work Hours) ---
    if 'Date' in eng_df.columns:
        eng_df = eng_df[eng_df['Date'].astype(str) != 'Total']
        
    travel = pd.to_numeric(eng_df["Travel"], errors="coerce").fillna(0).sum() if "Travel" in eng_df.columns else 0.0
    travel_ot = pd.to_numeric(eng_df["Travel OT"], errors="coerce").fillna(0).sum() if "Travel OT" in eng_df.columns else 0.0
    travel_sum = travel + travel_ot
    
    nt_col = "Normal Time" if "Normal Time" in eng_df.columns else ("NT" if "NT" in eng_df.columns else None)
    nt_sum = pd.to_numeric(eng_df[nt_col], errors="coerce").fillna(0).sum() if nt_col and nt_col in eng_df.columns else 0.0
    ot_sum = pd.to_numeric(eng_df["OT"], errors="coerce").fillna(0).sum() if "OT" in eng_df.columns else 0.0
    
    waiting_sum = pd.to_numeric(eng_df["Waiting time"], errors="coerce").fillna(0).sum() if "Waiting time" in eng_df.columns else 0.0
    prep_sum = pd.to_numeric(eng_df["Preparation"], errors="coerce").fillna(0).sum() if "Preparation" in eng_df.columns else 0.0
    
    # --- PROCESS CLIENT TIMESHEET (Local Transport) ---
    if 'Date' in client_df.columns:
        client_df = client_df[client_df['Date'].astype(str) != 'Total']
        
    l_trpt_sum = pd.to_numeric(client_df["L.Trpt"], errors="coerce").fillna(0).sum() if "L.Trpt" in client_df.columns else 0.0
    
    # --- LOAD AND FILL INVOICE TEMPLATE ---
    if template_file.name.lower().endswith('.csv'):
        raise ValueError("The Invoice Template must be an Excel file (.xlsx) to preserve formulas and formatting.")
        
    wb = load_workbook(template_file)
    
    sheet_map = {"SG": "SG", "CN": "CN", "KR": "KR ", "EUR": "EUR", "USD": "USD"}
    target_sheet = sheet_map.get(currency, currency)
    
    if target_sheet not in wb.sheetnames:
        if currency in wb.sheetnames:
            target_sheet = currency
        else:
            raise ValueError(f"The uploaded template does not contain a tab for {currency}.")
            
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
    
    # Determine Row Offset based on Engineer Role
    r_offset = 20
    if position == "Service Engineer": r_offset = 30
    elif position == "Senior Service Engineer": r_offset = 40
    elif position == "Specialist Service Engineer": r_offset = 50
        
    # Inject Hours into Invoice Table
    safe_write(ws, r_offset + 1, 4, travel_sum if travel_sum > 0 else "")
    safe_write(ws, r_offset + 2, 4, nt_sum if nt_sum > 0 else "")
    safe_write(ws, r_offset + 3, 4, ot_sum if ot_sum > 0 else "")
    safe_write(ws, r_offset + 4, 4, waiting_sum if waiting_sum > 0 else "")
    safe_write(ws, r_offset + 5, 4, prep_sum if prep_sum > 0 else "")
    
    expense_row = 59
    local_transport_row = None
    
    # Dynamically locate Expenses and Local Transport rows
    for row_idx in range(50, 75):
        col_b_val = str(ws.cell(row=row_idx, column=2).value).strip()
        col_c_val = str(ws.cell(row=row_idx, column=3).value).strip()
        
        if "Expenses" in col_b_val:
            expense_row = row_idx
            
        if "local transport" in col_c_val.lower() or "transportation" in col_c_val.lower():
            local_transport_row = row_idx
            break
            
    if not local_transport_row:
        local_transport_row = 63
        safe_write(ws, local_transport_row, 3, "Local Transport")

    safe_write(ws, expense_row + 2, 3, engineer_name)
        
    if l_trpt_sum > 0:
        safe_write(ws, local_transport_row, 4, l_trpt_sum)

    # Inject User Custom Expenses
    if user_expenses:
        # Scan the rows immediately below the 'Expenses' header
        for row_idx in range(expense_row + 1, expense_row + 20):
            cell_desc = str(ws.cell(row=row_idx, column=3).value).strip()
            
            for exp in user_expenses:
                if cell_desc == exp['desc']:
                    if exp['qty'] > 0:
                        safe_write(ws, row_idx, 4, exp['qty'])
                    if exp['price'] > 0:
                        safe_write(ws, row_idx, 6, exp['price'])

    # --- DYNAMIC INVOICE TOTALS & TAX LOGIC ---
    if currency == "CN":
        if include_admin_fee == "No":
            safe_write(ws, 82, 7, "-")
        else:
            safe_write(ws, 82, 7, "=0.1*G67")
        safe_write(ws, 84, 7, "=0.06*(SUM(G28,G38,G48,G58,G67,G74,G82))")
        safe_write(ws, 86, 7, "=SUM(G28,G38,G48,G58,G67,G74,G77:G80,G82,G84)")

    elif currency == "KR":
        if include_admin_fee == "No":
            safe_write(ws, 82, 3, "-")
        else:
            safe_write(ws, 82, 3, "=0.1*(SUM(G63:G66))")
        safe_write(ws, 84, 3, "=SUM(G41:G47,G61:G63,C82,G31:G37, G21:G27, G51:G57)")

    elif currency == "SG":
        if include_admin_fee == "No":
            safe_write(ws, 78, 3, "-")
        else:
            safe_write(ws, 78, 3, "=SUM(G61:G62)*0.1")
        safe_write(ws, 80, 3, "=SUM(G41:G47,G61:G63,C78,G31:G37, G21:G27, G51:G57)")

    elif currency == "EUR":
        if include_admin_fee == "No":
            safe_write(ws, 82, 3, "-")
        else:
            safe_write(ws, 82, 3, "=0.1*(SUM(G61:G66))")
        safe_write(ws, 84, 3, "=SUM(G41:G47,G61:G66,C82,G31:G37, G21:G27,G51:G57,G70:G73,G77:G80)")

    elif currency == "USD":
        if include_admin_fee == "No":
            safe_write(ws, 82, 3, "-")
        else:
            safe_write(ws, 82, 3, "=0.1*(SUM(G63:G66))")
        safe_write(ws, 84, 3, "=SUM(G21:G27,G31:G37,G41:G47,G51:G57,G61:G66,G70:G73,G77:G80,C82)")

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

    # Render dynamic expenses UI
    user_expenses_t1 = render_expense_ui("t1")

    if st.button("Generate Final Invoice", type="primary", key="btn_t1"):
        if not timesheet_excel_t1 or not template_excel_t1:
            st.error("Please upload the Processed Timesheet AND the Invoice Template.")
        else:
            try:
                xls = pd.ExcelFile(timesheet_excel_t1)
                sheet_names = xls.sheet_names
                
                # Dynamically locate sheets based on keywords, fallback to index
                eng_sheet = next((s for s in sheet_names if 'engineer' in s.lower()), sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])
                client_sheet = next((s for s in sheet_names if 'client' in s.lower()), sheet_names[0])
                
                eng_df = pd.read_excel(timesheet_excel_t1, sheet_name=eng_sheet, skiprows=2)
                client_df = pd.read_excel(timesheet_excel_t1, sheet_name=client_sheet, skiprows=2)
                
                output = process_invoice_logic(
                    eng_df, client_df, template_excel_t1, 
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
            except KeyError as e:
                st.error(f"Missing expected column in timesheet: {str(e)}. Please check the uploaded file format.")
            except ValueError as e:
                st.error(f"Error encountered: {str(e)}")
            except zipfile.BadZipFile:
                st.error("One of the uploaded files is not a valid Excel file or is corrupted.")
            except Exception as e:
                st.error(f"An unexpected error occurred while processing the invoice: {str(e)}")

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

    # Render dynamic expenses UI
    user_expenses_t2 = render_expense_ui("t2")

    if st.button("Generate Final Invoice", type="primary", key="btn_t2"):
        if not client_timesheet_t2 or not template_excel_t2:
            st.error("Please upload the Client Timesheet AND the Invoice Template.")
        else:
            try:
                if client_timesheet_t2.name.lower().endswith('.csv'):
                    client_df = pd.read_csv(client_timesheet_t2)
                else:
                    client_df = pd.read_excel(client_timesheet_t2, sheet_name=0)
                
                eng_df = client_df.copy()
                
                output = process_invoice_logic(
                    eng_df, client_df, template_excel_t2, 
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
            except KeyError as e:
                st.error(f"Missing expected column in timesheet: {str(e)}. Please check the uploaded file format.")
            except ValueError as e:
                st.error(f"Error encountered: {str(e)}")
            except zipfile.BadZipFile:
                st.error("One of the uploaded files is not a valid Excel file or is corrupted.")
            except Exception as e:
                st.error(f"An unexpected error occurred while processing the invoice: {str(e)}")
