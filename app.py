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

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="Invoice Generator", layout="wide")

st.title("Final Invoice Generation")
st.write("Select your timesheet format and generate the final invoice template.")

tab1, tab2 = st.tabs(["Single Timesheet Upload (Combined)", "Dual Timesheet Upload (Separate)"])

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
        cust_name_t1 = st.text_input("Customer name", key="cust_name_t1")
        inv_address_t1 = st.text_input("Invoicing address", key="inv_addr_t1")
        del_address_t1 = st.text_input("Delivery address", key="del_addr_t1")
        reference_t1 = st.text_input("Reference", key="ref_t1")
        cust_po_t1 = st.text_input("Customer PO", key="po_t1")
    with c2_t1:
        proj_no_t1 = st.text_input("Project No", key="proj_t1")
        svc_type_t1 = st.text_input("Service Type", key="svc_t1")
        vessel_name_t1 = st.text_input("Vessel Name", key="vessel_t1")
        vessel_no_t1 = st.text_input("Vessel No (if applicable)", key="vessel_no_t1")
        engineer_name_invoice_t1 = st.text_input("Engineer Name (For Expenses)", key="eng_name_t1")

    st.markdown("### 3. Type of Service")
    service_category_t1 = st.radio("Select Type of Service:", ["Internal", "External"], key="svc_cat_t1")
        
    st.markdown("### 4. Select Engineer Role")
    position_t1 = st.selectbox("Assign Hours to Position:", [
        "Service Technician", 
        "Service Engineer", 
        "Senior Service Engineer", 
        "Specialist Service Engineer"
    ], key="pos_t1")

    if st.button("Generate Final Invoice", type="primary", key="btn_t1"):
        if not timesheet_excel_t1 or not template_excel_t1:
            st.error("Please upload the Processed Timesheet AND the Invoice Template.")
        else:
            try:
                client_df = pd.read_excel(timesheet_excel_t1, sheet_name="Client", skiprows=2)
                
                travel_sum = pd.to_numeric(client_df["Travel"], errors="coerce").sum()
                nt_sum = pd.to_numeric(client_df["NT"], errors="coerce").sum()
                ot_sum = pd.to_numeric(client_df["OT"], errors="coerce").sum()
                waiting_sum = pd.to_numeric(client_df["Waiting time"], errors="coerce").sum()
                prep_sum = pd.to_numeric(client_df["Preparation"], errors="coerce").sum()
                l_trpt_sum = pd.to_numeric(client_df["L.Trpt"], errors="coerce").sum()
                
                wb = load_workbook(template_excel_t1)
                if "SG" not in wb.sheetnames:
                    st.error("The uploaded template does not contain an 'SG' tab.")
                    st.stop()
                    
                ws = wb["SG"]
                
                safe_write(ws, 7, 3, cust_name_t1)
                safe_write(ws, 8, 3, inv_address_t1)
                safe_write(ws, 9, 3, del_address_t1)
                safe_write(ws, 10, 3, reference_t1)
                safe_write(ws, 11, 3, cust_po_t1)
                safe_write(ws, 12, 3, proj_no_t1)
                safe_write(ws, 13, 3, svc_type_t1)
                safe_write(ws, 14, 3, vessel_name_t1)
                safe_write(ws, 15, 3, vessel_no_t1)
                
                r_offset = 20
                if position_t1 == "Service Engineer":
                    r_offset = 30
                elif position_t1 == "Senior Service Engineer":
                    r_offset = 40
                elif position_t1 == "Specialist Service Engineer":
                    r_offset = 50
                    
                safe_write(ws, r_offset + 1, 4, travel_sum if travel_sum > 0 else "")
                safe_write(ws, r_offset + 2, 4, nt_sum if nt_sum > 0 else "")
                safe_write(ws, r_offset + 3, 4, ot_sum if ot_sum > 0 else "")
                safe_write(ws, r_offset + 4, 4, waiting_sum if waiting_sum > 0 else "")
                safe_write(ws, r_offset + 5, 4, prep_sum if prep_sum > 0 else "")
                
                expense_row = None
                local_transport_row = None
                
                for row_idx in range(1, 150):
                    col_b_val = ws.cell(row=row_idx, column=2).value
                    col_c_val = ws.cell(row=row_idx, column=3).value
                    
                    if col_b_val and str(col_b_val).strip() == "Expenses":
                        expense_row = row_idx
                        
                    if col_c_val and "local transport" in str(col_c_val).lower():
                        local_transport_row = row_idx
                        
                if expense_row:
                    safe_write(ws, expense_row + 2, 3, engineer_name_invoice_t1)
                    
                if local_transport_row and l_trpt_sum > 0:
                    safe_write(ws, local_transport_row, 4, l_trpt_sum)

                if service_category_t1 == "Internal":
                    safe_write(ws, 78, 3, "-")
                else:
                    safe_write(ws, 78, 3, "=SUM(G41:G47,G61:G63,G31:G37, G21:G27, G51:G57)*0.1")

                safe_write(ws, 80, 3, "=SUM(G41:G47,G61:G63,C78,G31:G37, G21:G27, G51:G57)")
                
                invoice_output = io.BytesIO()
                wb.save(invoice_output)
                invoice_output.seek(0)
                
                st.success("Invoice Generated Successfully")
                st.download_button(
                    label="Download Final Invoice",
                    data=invoice_output,
                    file_name=f"Invoice_{cust_name_t1 or 'Completed'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_t1"
                )
            except Exception as e:
                st.error(f"An error occurred while processing the invoice: {str(e)}")

# ------------------------------------------------------------
# TAB 2: DUAL TIMESHEET UPLOAD
# ------------------------------------------------------------
with tab2:
    st.markdown("### 1. Upload Required Files")
    col1_t2, col2_t2, col3_t2 = st.columns(3)
    with col1_t2:
        engineer_timesheet_t2 = st.file_uploader("Upload Engineer Timesheet", type=["xlsx"], key="eng_upload_t2")
    with col2_t2:
        client_timesheet_t2 = st.file_uploader("Upload Client Timesheet", type=["xlsx"], key="cli_upload_t2")
    with col3_t2:
        template_excel_t2 = st.file_uploader("Upload Blank Invoice Template", type=["xlsx"], key="inv_upload_t2")
        
    st.markdown("### 2. Enter Information")
    c1_t2, c2_t2 = st.columns(2)
    with c1_t2:
        cust_name_t2 = st.text_input("Customer name", key="cust_name_t2")
        inv_address_t2 = st.text_input("Invoicing address", key="inv_addr_t2")
        del_address_t2 = st.text_input("Delivery address", key="del_addr_t2")
        reference_t2 = st.text_input("Reference", key="ref_t2")
        cust_po_t2 = st.text_input("Customer PO", key="po_t2")
    with c2_t2:
        proj_no_t2 = st.text_input("Project No", key="proj_t2")
        svc_type_t2 = st.text_input("Service Type", key="svc_t2")
        vessel_name_t2 = st.text_input("Vessel Name", key="vessel_t2")
        vessel_no_t2 = st.text_input("Vessel No (if applicable)", key="vessel_no_t2")
        engineer_name_invoice_t2 = st.text_input("Engineer Name (For Expenses)", key="eng_name_t2")

    st.markdown("### 3. Type of Service")
    service_category_t2 = st.radio("Select Type of Service:", ["Internal", "External"], key="svc_cat_t2")
        
    st.markdown("### 4. Select Engineer Role")
    position_t2 = st.selectbox("Assign Hours to Position:", [
        "Service Technician", 
        "Service Engineer", 
        "Senior Service Engineer", 
        "Specialist Service Engineer"
    ], key="pos_t2")

    if st.button("Generate Final Invoice", type="primary", key="btn_t2"):
        if not engineer_timesheet_t2 or not client_timesheet_t2 or not template_excel_t2:
            st.error("Please upload the Engineer Timesheet, Client Timesheet, AND the Invoice Template.")
        else:
            try:
                ts_df = pd.read_excel(engineer_timesheet_t2, sheet_name=0, skiprows=3)
                
                if 'Date' in ts_df.columns:
                    ts_df = ts_df[ts_df['Date'] != 'Total']
                    
                travel = pd.to_numeric(ts_df["Travel"], errors="coerce").sum() if "Travel" in ts_df.columns else 0.0
                travel_ot = pd.to_numeric(ts_df["Travel OT"], errors="coerce").sum() if "Travel OT" in ts_df.columns else 0.0
                travel_sum = travel + travel_ot
                
                nt_sum = pd.to_numeric(ts_df["Normal Time"], errors="coerce").sum() if "Normal Time" in ts_df.columns else pd.to_numeric(ts_df["NT"], errors="coerce").sum() if "NT" in ts_df.columns else 0.0
                ot_sum = pd.to_numeric(ts_df["OT"], errors="coerce").sum() if "OT" in ts_df.columns else 0.0
                
                waiting_sum = pd.to_numeric(ts_df["Waiting time"], errors="coerce").sum() if "Waiting time" in ts_df.columns else 0.0
                prep_sum = pd.to_numeric(ts_df["Preparation"], errors="coerce").sum() if "Preparation" in ts_df.columns else 0.0
                
                client_df = pd.read_excel(client_timesheet_t2, sheet_name=0, skiprows=3)
                
                if 'Date' in client_df.columns:
                    client_df = client_df[client_df['Date'] != 'Total']
                    
                l_trpt_sum = pd.to_numeric(client_df["L.Trpt"], errors="coerce").sum() if "L.Trpt" in client_df.columns else 0.0
                
                wb = load_workbook(template_excel_t2)
                if "SG" not in wb.sheetnames:
                    st.error("The uploaded template does not contain an 'SG' tab.")
                    st.stop()
                    
                ws = wb["SG"]
                
                safe_write(ws, 7, 3, cust_name_t2)
                safe_write(ws, 8, 3, inv_address_t2)
                safe_write(ws, 9, 3, del_address_t2)
                safe_write(ws, 10, 3, reference_t2)
                safe_write(ws, 11, 3, cust_po_t2)
                safe_write(ws, 12, 3, proj_no_t2)
                safe_write(ws, 13, 3, svc_type_t2)
                safe_write(ws, 14, 3, vessel_name_t2)
                safe_write(ws, 15, 3, vessel_no_t2)
                
                r_offset = 20
                if position_t2 == "Service Engineer":
                    r_offset = 30
                elif position_t2 == "Senior Service Engineer":
                    r_offset = 40
                elif position_t2 == "Specialist Service Engineer":
                    r_offset = 50
                    
                safe_write(ws, r_offset + 1, 4, travel_sum if travel_sum > 0 else "")
                safe_write(ws, r_offset + 2, 4, nt_sum if nt_sum > 0 else "")
                safe_write(ws, r_offset + 3, 4, ot_sum if ot_sum > 0 else "")
                safe_write(ws, r_offset + 4, 4, waiting_sum if waiting_sum > 0 else "")
                safe_write(ws, r_offset + 5, 4, prep_sum if prep_sum > 0 else "")
                
                expense_row = None
                local_transport_row = None
                
                for row_idx in range(1, 150):
                    col_b_val = ws.cell(row=row_idx, column=2).value
                    col_c_val = ws.cell(row=row_idx, column=3).value
                    
                    if col_b_val and str(col_b_val).strip() == "Expenses":
                        expense_row = row_idx
                        
                    if col_c_val and "local transport" in str(col_c_val).lower():
                        local_transport_row = row_idx
                        
                if expense_row:
                    safe_write(ws, expense_row + 2, 3, engineer_name_invoice_t2)
                    
                if local_transport_row and l_trpt_sum > 0:
                    safe_write(ws, local_transport_row, 4, l_trpt_sum)

                if service_category_t2 == "Internal":
                    safe_write(ws, 78, 3, "-")
                else:
                    safe_write(ws, 78, 3, "=SUM(G41:G47,G61:G63,G31:G37, G21:G27, G51:G57)*0.1")

                safe_write(ws, 80, 3, "=SUM(G41:G47,G61:G63,C78,G31:G37, G21:G27, G51:G57)")
                
                invoice_output = io.BytesIO()
                wb.save(invoice_output)
                invoice_output.seek(0)
                
                st.success("Invoice Generated Successfully")
                st.download_button(
                    label="Download Final Invoice",
                    data=invoice_output,
                    file_name=f"Invoice_{cust_name_t2 or 'Completed'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_t2"
                )
            except Exception as e:
                st.error(f"An error occurred while processing the invoice: {str(e)}")
