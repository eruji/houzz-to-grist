# Houzz to Grist Converter

A tool to convert Houzz Pro proposal exports (Excel) into a clean CSV format ready for import into Grist.

## Features
- **Web App:** Modern web interface with data preview.
- **Smart Filtering:** Automatically removes parent rows if detailed child rows exist.
- **Vendor Extraction:** Parses vendor names from website URLs.
- **Dynamic Schema:** Adapts to your `ORDERS.csv` column layout automatically.

## How to Use

## How to Use

1.  Run the application (e.g. `streamlit run app.py`).
2.  Enter the Proposal #.
3.  Upload your file, review the data, and click **Copy for Grist** or **Download**.

## Requirements
- Python 3.x
- Pandas, Openpyxl, Streamlit (installed via setup script)
