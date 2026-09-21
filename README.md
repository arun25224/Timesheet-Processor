# Timesheet-Processor
Timesheet-Processor

A Python-based web application built with Streamlit that processes structured timesheet data and dynamically injects calculated billable hours and custom expenses into complex, multi-currency Excel invoice templates.

This application serves as the final formatting and generation layer of the invoicing pipeline, eliminating manual data entry while preserving downstream native spreadsheet formulas.

## Core Capabilities

* **Template-Driven Excel Injection:** Utilizes `openpyxl` to map extracted hours to absolute rows based on engineering roles, while using anchor-text scanning to dynamically inject custom expenses regardless of template length variations.
* **Multi-Currency Routing:** Automatically locates and writes to the correct worksheet tab (SG, CN, KR, EUR, USD) within the master template using flexible keyword matching.
* **Dynamic Expense Logging:** Provides a user interface to append itemized incidentals (e.g., visas, hotel, shipyard passes), which are programmatically written into designated template placeholders.
* **Formula Preservation:** Injects raw numeric values without breaking existing downstream taxation or summation arrays. Includes a targeted formula-wipe algorithm for conditional administrative fees.
* **Auto-Header Detection:** Scans uploaded timesheet datasets to dynamically locate the correct header row, ensuring stability against files with variable row-skipping or formatting inconsistencies.

## Repository Structure

* `app.py`: The primary application script containing the Streamlit frontend UI, session state management, and the Pandas/Openpyxl injection logic.
* `requirements.txt`: The dependency mapping file required for environment replication.
* `Invoice Template.xlsx`: (To be provided by the user) The master corporate spreadsheet containing necessary text anchors and billing formulas.

## Technology Stack

* **Frontend Framework:** Streamlit
* **Data Processing:** Pandas
* **Spreadsheet Manipulation:** Openpyxl

## Installation and Setup

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/yourusername/timesheet-processor.git](https://github.com/yourusername/timesheet-processor.git)
   cd timesheet-processor
