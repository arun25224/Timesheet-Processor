import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# ============================================================
# UTILITY FUNCTIONS
# ============================================================
def safe_write(ws, row_idx, col_idx, value):
    try:
        ws.cell(row=row_idx, column=col_idx).value = value if value != "" else None
    except AttributeError:
        coord = ws.cell(row=row_idx, column=col_idx).coordinate
        for merged_range in ws.merged_cells.ranges:
            if coord in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value if value != "" else None
                break

def clean_and_find_headers(df):
    """Automatically hunts down the actual header row in any uploaded timesheet."""
    current_cols = df.columns.astype(str).str.lower().str.strip()
    valid_keywords = ['date', 'travel', 'nt', 'normal time', 'ot', 'overtime', 'waiting time']
    
    if any(k in current_cols for k in valid_keywords):
        df.columns = current_cols
        return df
        
    for idx, row in df.head(10).iterrows():
        row_str = row.astype(str).str.lower().str.strip().tolist()
        if any(k in row_str for k in valid_keywords):
            df.columns = row_str
            return df.iloc[idx+1:].reset_index(drop=True)
            
    df.columns = current_cols
    return df

def get_sum(df, possible_cols):
    """Safely extracts sums from columns regardless of slight naming variations."""
    for c in possible_cols:
        if c in df.columns:
            return pd.to_numeric(df[c], errors='coerce').fillna(0).sum()
    return 0.0

def extract_metadata(df):
    """Scans the raw dataframe to extract metadata for auto-filling the Streamlit UI."""
    metadata = {
        "wo": "", "cust_name": "", "proj_no": "", "po": "", 
        "svc": "", "vessel": "", "vessel_no": "", "eng_name": "", "place": ""
    }
    for index, row in df.iterrows():
        row_str = row.astype(str).str.lower().str.strip().tolist()
        
        if 'work order number' in row_str:
            idx = row_str.index('work order number')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['wo'] = val
                
        if 'customer name' in row_str:
            idx = row_str.index('customer name')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['cust_name'] = val
                
        if 'project number' in row_str:
            idx = row_str.index('project number')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['proj_no'] = val
                
        if 'customer po number' in row_str:
            idx = row_str.index('customer po number')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['po'] = val
                
        if 'service type' in row_str:
            idx = row_str.index('service type')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['svc'] = val
                
        if 'vessel name' in row_str:
            idx = row_str.index('vessel name')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['vessel'] = val
                
        if 'imo or vessel number' in row_str:
            idx = row_str.index('imo or vessel number')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['vessel_no'] = val
                
        if 'engineer name shown in activity table' in row_str:
            idx = row_str.index('engineer name shown in activity table')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['eng_name'] = val
                
        if 'place of attendance' in row_str:
            idx = row_str.index('place of attendance')
            if idx + 1 < len(row):
                val = str(row.iloc[idx+1]).strip()
                if val.lower() != 'nan' and val != '': metadata['place'] = val
                
    return metadata

def render_expense_ui(tab_key):
    if f"expenses_{tab_key}" not in st.session_state:
        st.session_state[f"expenses_{tab_key}"] = []
        
    st.markdown("### 4. Additional Expenses (Optional)")
    
    expense_options = [
        "Shipyard Pass", "Taxi Overseas", "Ferry Fare", "Visa", "Hotel", 
        "Laundry", "Agent Fee", "Excess Baggage Fee", "Airport Tax", 
        "Flight Ticket", "Misc", "Local transport", "Allowance"
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
            exp['price'] = st.number_input("Price", min_value=0.0, value=float(exp.get('price', 0.0)), step=0.01, key=f"price_{tab_key}_{i}")
        with col4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("X", key=f"del_{tab_key}_{i}"):
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
    eng_df = clean_and_find_headers(eng_df)
    client_df = clean_and_find_headers(client_df)

    if 'date' in eng_df.columns:
        eng_df = eng_df[eng_df['date'].astype(str).str.lower() != 'total']
    if 'date' in client_df.columns:
        client_df = client_df[client_df['date'].astype(str).str.lower() != 'total']
        
    travel_sum = get_sum(eng_df, ["travel", "travel time"]) + get_sum(eng_df, ["travel ot", "travel time ot"])
    nt_sum = get_sum(eng_df, ["nt", "normal time", "normal"])
    ot_sum = get_sum(eng_df, ["ot", "overtime"])
    waiting_sum = get_sum(eng_df, ["waiting time", "waiting"]) + get_sum(eng_df, ["waiting time ot", "waiting ot"])
    prep_sum = get_sum(eng_df, ["preparation", "preparation time", "prep"])
    
    l_trpt_sum = get_sum(client_df, ["l.trpt", "local transport", "transport"])
    
    extracted_info = {
        "travel": travel_sum,
        "nt": nt_sum,
        "ot": ot_sum,
        "waiting": waiting_sum,
        "prep": prep_sum,
        "transport": l_trpt_sum
    }
    
    if template_file.name.lower().endswith('.csv'):
        raise ValueError("The Invoice Template must be an Excel file (.xlsx).")
        
    wb = load_workbook(template_file)
    
    currency_keywords = {
        "SG": ["sg", "singapore", "sgd"],
        "CN": ["cn", "china", "cny", "rmb"],
        "KR": ["kr", "korea", "krw", "won"],
        "EUR": ["eur", "europe", "euro"],
        "USD": ["usd", "us", "america", "dollar"]
    }
    
    target_sheet = None
    for s in wb.sheetnames:
        if s.strip().upper() == currency.upper():
            target_sheet = s
            break
            
    if not target_sheet:
        for s in wb.sheetnames:
            if any(k in s.lower() for k in currency_keywords.get(currency, [])):
                target_sheet = s
                break
                
    if not target_sheet:
        target_sheet = wb.sheetnames[0]
            
    ws = wb[target_sheet]
    
    safe_write(ws, 7, 3, cust_name)
    safe_write(ws, 8, 3, inv_address)
    safe_write(ws, 9, 3, del_address)
    safe_write(ws, 10, 3, reference)
    safe_write(ws, 11, 3, cust_po)
    safe_write(ws, 12, 3, proj_no)
    safe_write(ws, 13, 3, svc_type)
    safe_write(ws, 14, 3, vessel_name)
    safe_write(ws, 15, 3, vessel_no)
    
    role_base_row_map = {
        "Service Technician": 20,
        "Service Engineer": 30,
        "Senior Service Engineer": 40,
        "Specialist Service Engineer": 50
    }
    
    base_row = role_base_row_map.get(position)
    
    if base_row:
        safe_write(ws, base_row + 1, 4, travel_sum)
        safe_write(ws, base_row + 2, 4, nt_sum)
        safe_write(ws, base_row + 3, 4, ot_sum)
        safe_write(ws, base_row + 4, 4, waiting_sum)
        safe_write(ws, base_row + 5, 4, prep_sum)
    
    # Inject Engineer Name strictly into C61
    if engineer_name:
        safe_write(ws, 61, 3, engineer_name)

    if l_trpt_sum > 0:
        for r in range(50, 80):
            c3_val = str(ws.cell(row=r, column=3).value).strip().lower()
            if c3_val == "local transport":
                safe_write(ws, r, 4, l_trpt_sum)
                break

    if user_expenses:
        for exp in user_expenses:
            if exp['desc'] == "Allowance":
                if exp['qty'] > 0:
                    safe_write(ws, 61, 4, exp['qty'])
                if exp['price'] > 0:
                    safe_write(ws, 61, 6, exp['price'])
                continue
                
            for r in range(50, 80):
                c3_val = str(ws.cell(row=r, column=3).value).strip().lower()
                if exp['desc'].lower() == c3_val:
                    if exp['qty'] > 0:
                        safe_write(ws, r, 4, exp['qty'])
                    if exp['price'] > 0:
                        safe_write(ws, r, 6, exp['price'])
                    break

    if include_admin_fee == "No":
        for r in range(80, 100):
            val = str(ws.cell(row=r, column=2).value).strip().lower()
            if "admin" in val and "fee" in val:
                for c in range(3, 10):
                    if ws.cell(row=r, column=c).value is not None:
                        safe_write(ws, r, c, None)
                safe_write(ws, r, 7, 0)
                break

    invoice_output = io.BytesIO()
    wb.save(invoice_output)
    invoice_output.seek(0)
    
    return {"file": invoice_output, "info": extracted_info}

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="Invoice Generator", layout="wide")

st.title("Final Invoice Generation")
st.write("Select your timesheet and generate the invoice.")

tab1, tab2 = st.tabs(["Single Timesheet Upload (Combined Excel)", "SANA Timesheet Upload"])

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
        
    if st.button("Extract Information", key="extract_t1"):
        if timesheet_excel_t1:
            try:
                xls_temp = pd.ExcelFile(timesheet_excel_t1)
                sheet_names_temp = xls_temp.sheet_names
                client_sheet_temp = next((s for s in sheet_names_temp if 'client' in s.lower()), sheet_names_temp[0])
                df_temp = pd.read_excel(timesheet_excel_t1, sheet_name=client_sheet_temp)
                
                meta_t1 = extract_metadata(df_temp)
                
                st.session_state.wo_t1 = meta_t1.get("wo", "")
                st.session_state.cust_name_t1 = meta_t1.get("cust_name", "")
                st.session_state.po_t1 = meta_t1.get("po", "")
                st.session_state.proj_t1 = meta_t1.get("proj_no", "")
                st.session_state.svc_t1 = meta_t1.get("svc", "")
                st.session_state.vessel_t1 = meta_t1.get("vessel", "")
                st.session_state.vessel_no_t1 = meta_t1.get("vessel_no", "")
                st.session_state.eng_name_t1 = meta_t1.get("eng_name", "")
                st.session_state.del_addr_t1 = meta_t1.get("place", "")
                
                st.success("Information extracted.")
            except Exception as e:
                st.error(f"Failed to extract information: {str(e)}")
        else:
            st.warning("Please upload a timesheet first.")

    st.markdown("### 2. Enter Information")
    c1_t1, c2_t1 = st.columns(2)
    with c1_t1:
        work_order_t1 = st.text_input("Work Order Number", key="wo_t1")
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
                client_sheet = next((s for s in sheet_names if 'client' in s.lower()), sheet_names[0])
                
                eng_df = pd.read_excel(timesheet_excel_t1, sheet_name=eng_sheet)
                client_df = pd.read_excel(timesheet_excel_t1, sheet_name=client_sheet)
                
                output = process_invoice_logic(
                    eng_df, client_df, template_excel_t1, 
                    cust_name_t1, inv_address_t1, del_address_t1, reference_t1, cust_po_t1, 
                    proj_no_t1, svc_type_t1, vessel_name_t1, vessel_no_t1, engineer_name_invoice_t1, 
                    include_admin_fee_t1, position_t1, currency_t1, user_expenses_t1
                )
                
                info = output["info"]
                st.info(f" **Data Extracted:**\n"
                        f"- **Travel Time:** {info['travel']} hours\n"
                        f"- **Normal Time:** {info['nt']} hours\n"
                        f"- **Overtime:** {info['ot']} hours\n"
                        f"- **Waiting Time:** {info['waiting']} hours\n"
                        f"- **Preparation Time:** {info['prep']} hours\n"
                        f"- **Local Transport:** {info['transport']} units")
                
                st.success("Invoice Generated.")
                st.download_button(
                    label="Download Final Invoice",
                    data=output["file"],
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
        
    if st.button("Extract Information", key="extract_t2"):
        if client_timesheet_t2:
            try:
                if client_timesheet_t2.name.lower().endswith('.csv'):
                    df_temp2 = pd.read_csv(client_timesheet_t2)
                else:
                    df_temp2 = pd.read_excel(client_timesheet_t2, sheet_name=0)
                
                meta_t2 = extract_metadata(df_temp2)
                
                st.session_state.wo_t2 = meta_t2.get("wo", "")
                st.session_state.cust_name_t2 = meta_t2.get("cust_name", "")
                st.session_state.po_t2 = meta_t2.get("po", "")
                st.session_state.proj_t2 = meta_t2.get("proj_no", "")
                st.session_state.svc_t2 = meta_t2.get("svc", "")
                st.session_state.vessel_t2 = meta_t2.get("vessel", "")
                st.session_state.vessel_no_t2 = meta_t2.get("vessel_no", "")
                st.session_state.eng_name_t2 = meta_t2.get("eng_name", "")
                st.session_state.del_addr_t2 = meta_t2.get("place", "")
                
                st.success("Information extracted.")
            except Exception as e:
                st.error(f"Failed to extract information: {str(e)}")
        else:
            st.warning("Please upload a timesheet first.")

    st.markdown("### 2. Enter Information")
    c1_t2, c2_t2 = st.columns(2)
    with c1_t2:
        work_order_t2 = st.text_input("Work Order Number", key="wo_t2")
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
                    df_t2, df_t2, template_excel_t2, 
                    cust_name_t2, inv_address_t2, del_address_t2, reference_t2, cust_po_t2, 
                    proj_no_t2, svc_type_t2, vessel_name_t2, vessel_no_t2, engineer_name_invoice_t2, 
                    include_admin_fee_t2, position_t2, currency_t2, user_expenses_t2
                )
                
                info = output["info"]
                st.info(f" **Data Extracted:**\n"
                        f"- **Travel Time:** {info['travel']} hours\n"
                        f"- **Normal Time:** {info['nt']} hours\n"
                        f"- **Overtime:** {info['ot']} hours\n"
                        f"- **Waiting Time:** {info['waiting']} hours\n"
                        f"- **Preparation Time:** {info['prep']} hours\n"
                        f"- **Local Transport:** {info['transport']} units")
                
                st.success("Invoice Generated.")
                st.download_button(
                    label="Download Final Invoice",
                    data=output["file"],
                    file_name=f"Invoice_{cust_name_t2 or 'Completed'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_t2"
                )
            except Exception as e:
                st.error(f"Error occurred while processing: {str(e)}")
