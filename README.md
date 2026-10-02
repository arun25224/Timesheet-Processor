# Timesheet-Processor

A two-stage operational pipeline designed to bridge the gap between field service reports and finalized, client-ready financial documents. 

This system eliminates manual data entry, enforces strict labor law compliance for overtime and holiday billing, and injects calculated hours and custom expenses into complex, multi-currency Excel invoice templates without breaking downstream formulas.

## How the System Works

The pipeline is split into an AI-driven extraction phase and a Python-based formatting and injection phase.

### Phase 1: Data Extraction
Field engineers often submit unstructured timesheets—ranging from scanned PDFs to handwritten images. The first half of the system relies on an LLM-driven OCR agent (Sana AI, which uses OpenAI models) to interpret these documents. 
* **Duration Calculation:** The AI calculates exact shift durations (End Time minus Start Time).
* **Strict Labor Compliance:** It evaluates every logged hour against strict `08:00 to 16:00` boundaries. Hours within this window are classified as Normal Time; anything outside is flagged as Overtime.
* **Public Holiday Cross-Referencing:** The agent cross-references the dates against a programmed list of Singapore Ministry of Manpower (MOM) public holidays, automatically applying weekend/overtime multipliers to relevant shifts (including Sunday off-in-lieu rules).
* **Output:** It generates two highly structured, clean datasets: an Engineer Timesheet and a Client Timesheet.

### Phase 2: Processing
Once the data is extracted, `app.py` takes over. It ingests the structured timesheet data alongside the company's blank Excel invoice template.
* **Auto-Header Detection & Aggregation:** The app scans the uploaded datasets to find the true header rows, bypassing messy formatting or empty top rows. It then calculates the final sums for Travel, Normal Time, Overtime, Waiting Time, Preparation, and Local Transport.
* **Currency Routing:** Based on user selection, the script scans the Excel workbook and automatically routes the data to the correct currency worksheet tab (SG, CN, KR, EUR, USD) using keyword matching.
* **Absolute Role Mapping:** The app locates the selected engineering role (e.g., Senior Service Engineer) and maps the aggregated hours directly into the corresponding absolute rows in the spreadsheet.
* **Dynamic Expense Handling:** Users can log custom incidentals (Visas, Shipyard Passes, Hotels) via the UI. The script scans the Excel file for exact description matches and injects the quantities and prices. 

## Core Capabilities

* **Template-Driven Excel Injection:** Utilises `openpyxl` to map data to where it belongs. It uses anchor-text scanning to dynamically inject custom expenses, ensuring the code doesn't break even if the template length varies.
* **Formula Preservation:** One of the biggest challenges with automated Excel generation is protecting existing logic. This system injects raw numeric values while strictly preserving downstream taxation and summation arrays. 
* **Targeted Formula-Wiping:** If a user opts out of the standard 10% Administrative Fee, the script removes the specific formula array for that fee, clears the row, and forces a `0` to prevent phantom calculations.

## Repository Structure

* `app.py`: The primary application script containing the Streamlit frontend UI, session state management, data aggregation algorithms, and the Pandas/Openpyxl injection logic.
* `requirements.txt`: The dependency mapping file required for environment replication.
* `sana_prompt`: The prompt that was added to the internal company platform for the reading and extraction of the timesheet.

## System Maintenance & Protocols

To ensure long-term system stability, two administrative rules must be adhered to:

### 1. Annual Public Holiday Updates
Because the AI extraction prompt relies on a hardcoded list of dates to prevent date hallucinations, the system prompt on the Sana AI platform must be updated annually with the upcoming year's official public holidays.

### 2. Template Anchor Preservation
The `app.py` script dynamically locates injection zones based on specific text strings within the uploaded Excel template. If the business redesigns the blank template, the following exact textual strings must remain unchanged in Column B and Column C:
* **Role Routing:** `Service Technician`, `Service Engineer`, `Senior Service Engineer`, `Specialist Service Engineer`. (Altering these will break the vertical hour placement logic).
* **Section Locators:** `Expenses` and `Local transport`.
* **Expense Targets:** `[ADD RELEVANT EXPENSES HERE]` and `[ADD DESCRIPTION HERE]`. 
