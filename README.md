# Houzz to Grist Converter

Convert a **Houzz Pro proposal export (Excel)** into a clean CSV/TSV matching
the `ORDERS.csv` Grist import template. Includes smart parent-row filtering,
vendor extraction from URLs, and currency/markup parsing.

There are two versions:

| Version | Where | Notes |
|---------|-------|-------|
| **Static web app** (recommended) | `site/` | 100% client-side — no server, no timeout, deployable to Netlify |
| Python / Streamlit (legacy) | `app.py`, `houzz_to_grist.py` | runs locally with `streamlit run app.py` |

## Static web app (Netlify)

Everything runs in the browser using [SheetJS](https://sheetjs.com). Your file
never leaves the page — there is no backend, no cold start, and no idle
timeout (unlike Streamlit).

### Usage

1. Open the site.
2. Enter the **Proposal #**.
3. Drop the Houzz `.xlsx`/`.xls` file (or click to browse).
4. Review the preview, then **Download CSV** or **Copy for Grist** (TSV, no headers).

### Deploy to Netlify

**Option A — Git import (recommended, auto-deploys on push):**

1. Push this repo to GitHub (the `site/` + `netlify.toml` are committed).
2. On [app.netlify.com](https://app.netlify.com) → **Add new site → Import an existing project** → GitHub → pick `houzz-to-grist`.
3. Build settings are already handled by `netlify.toml` (publish directory = `site`, no build command). Click **Deploy**.

**Option B — Netlify Drop (no account setup, drag-and-drop):**

1. Go to [app.netlify.com/drop](https://app.netlify.com/drop).
2. Drag the `site/` folder onto the page. Done.

### Local preview

```bash
cd site
python3 -m http.server 8000
# open http://127.0.0.1:8000
```

## Python version (legacy)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or headless conversion from the command line:

```bash
python houzz_to_grist.py "input.xlsx" "houzz_import.csv" "Art_12_16"
```

## Column mapping

The output columns are fixed to the `ORDERS.csv` header:

`Project, Proposal, Ordered, Item, Vendor, QTY, Unit COST, Markup%, Markup, Subtotal, Pre-Tax, Shipping, Total, Prepaid Tax?, Project Tax Rate, Tax, Prepaid Tax Amt, Created, Notes, Project_Project Tax Rate, Received, URL`

If your Grist template columns change, edit the `TARGET_COLUMNS` array at the
top of `site/app.js` (and `ORDERS.csv` for the Python version).
