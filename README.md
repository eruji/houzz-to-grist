# Houzz to Grist Converter

A tool to convert Houzz Pro proposal exports (Excel) into a clean CSV format ready for import into Grist.

## Features
- **Drag & Drop Interface:** Simple batch file for instant conversion.
- **Web App:** Modern local web interface with data preview.
- **Smart Filtering:** Automatically removes parent rows if detailed child rows exist.
- **Vendor Extraction:** Parses vendor names from website URLs.
- **Dynamic Schema:** Adapts to your `ORDERS.csv` column layout automatically.

## How to Use

### Option 1: Drag-and-Drop (Fastest)
1.  Drag your Excel file onto `Drag_Excel_Here.bat`.
2.  Enter the Project Name and Proposal # when prompted.
3.  The converted file `houzz_import.csv` will be created instantly.

### Option 2: Web App (Visual)
1.  Run `First_Time_Setup.bat` once to install requirements.
2.  Run `Start_Web_App.bat`.
3.  Upload your file, review the data, and click **Copy to Clipboard** or **Download**.

## Requirements
- Python 3.x
- Pandas, Openpyxl, Streamlit (installed via setup script)
