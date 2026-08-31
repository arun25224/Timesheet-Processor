import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

st.set_page_config(page_title="Invoice Generator", layout="wide")

st.title("Final Invoice Generation")
st.write("Fill in the customer details and generate the final invoice template.")

st.markdown("### 1. Upload Required Files")
col1, col2, col3 = st.columns(3)
with col1:
    engineer_timesheet = st.file_uploader("Upload Engineer Timesheet", type=["xlsx"], key="eng_upload")
with col2:
    client_timesheet = st.file_uploader("Upload Client Timesheet", type=["xlsx"], key="cli_upload")
with col3:
    template_excel = st.file_uploader("Upload Blank Invoice Template", type=["xlsx"], key="inv_upload")
    
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
    if not engineer_timesheet or not client_timesheet or not template_excel:
        st.error("Please upload the Engineer Timesheet, Client Timesheet, AND the Invoice Template.")
    else:
        try:
            # --- PROCESS ENGINEER TIMESHEET (Work Hours) ---
            ts_df = pd.read_excel(engineer_timesheet, sheet_name=0, skiprows=3)
            
            # Remove the "Total" row to prevent double-counting
            if 'Date' in ts_df.columns:
                ts_df = ts_df[ts_df['Date'] != 'Total']
                
            # Safely extract and calculate hours from Engineer
            travel = pd.to_numeric(ts_df["Travel"], errors="coerce").sum() if "Travel" in ts_df.columns else 0.0
            travel_ot = pd.to_numeric(ts_df["Travel OT"], errors="coerce").sum() if "Travel OT" in ts_df.columns else 0.0
            travel_sum = travel + travel_ot
            
            nt_sum = pd.to_numeric(ts_df["Normal Time"], errors="coerce").sum() if "Normal Time" in ts_df.columns else pd.to_numeric(ts_df["NT"], errors="coerce").sum() if "NT" in ts_df.columns else 0.0
            ot_sum = pd.to_numeric(ts_df["OT"], errors="coerce").sum() if "OT" in ts_df.columns else 0.0
            
            waiting_sum = pd.to_numeric(ts_df["Waiting time"], errors="coerce").sum() if "Waiting time" in ts_df.columns else 0.0
            prep_sum = pd.to_numeric(ts_df["Preparation"], errors="coerce").sum() if "Preparation" in ts_df.columns else 0.0
            
            # --- PROCESS CLIENT TIMESHEET (Local Transport) ---
            client_df = pd.read_excel(client_timesheet, sheet_name=0, skiprows=3)
            
            if 'Date' in client_df.columns:
                client_df = client_df[client_df['Date'] != 'Total']
                
            l_trpt_sum = pd.to_numeric(client_df["L.Trpt"], errors="coerce").sum() if "L.Trpt" in client_df.columns else 0.0
            
            # --- LOAD AND FILL INVOICE TEMPLATE ---
            wb = load_workbook(template_excel)
            if "SG" not in wb.sheetnames:
                st.error("The uploaded template does not contain an 'SG' tab.")
                st.stop()
                
            ws = wb["SG"]
            
            # Inject Customer Information
            ws["C7"] = cust_name
            ws["C8"] = inv_address
            ws["C9"] = del_address
            ws["C10"] = reference
            ws["C11"] = cust_po
            ws["C12"] = proj_no
            ws["C13"] = svc_type
            ws["C14"] = vessel_name
            ws["C15"] = vessel_no
            
            # Determine Row Offset based on Engineer Role
            r_offset = 20
            if position == "Service Engineer":
                r_offset = 30
            elif position == "Senior Service Engineer":
                r_offset = 40
            elif position == "Specialist Service Engineer":
                r_offset = 50
                
            # Inject Hours into Invoice Table
            ws[f"D{r_offset + 1}"] = travel_sum if travel_sum > 0 else ""
            ws[f"D{r_offset + 2}"] = nt_sum if nt_sum > 0 else ""
            ws[f"D{r_offset + 3}"] = ot_sum if ot_sum > 0 else ""
            ws[f"D{r_offset + 4}"] = waiting_sum if waiting_sum > 0 else ""
            ws[f"D{r_offset + 5}"] = prep_sum if prep_sum > 0 else ""
            
            expense_row = None
            local_transport_row = None
            
            # Dynamically locate Expenses and Local Transport rows
            for row_idx in range(1, 150):
                col_b_val = ws.cell(row=row_idx, column=2).value
                col_c_val = ws.cell(row=row_idx, column=3).value
                
                if col_b_val and str(col_b_val).strip() == "Expenses":
                    expense_row = row_idx
                    
                if col_c_val and "local transport" in str(col_c_val).lower():
                    local_transport_row = row_idx
                    
            if expense_row:
                ws.cell(row=expense_row + 2, column=3).value = engineer_name_invoice
                
            if local_transport_row and l_trpt_sum > 0:
                ws.cell(row=local_transport_row, column=5).value = l_trpt_sum
            
            # Export Final Invoice
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
