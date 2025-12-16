"""
Houzz to Grist Import Script
Converts Houzz Pro proposal exports (Excel) to Grist ORDERS.csv format
"""
import pandas as pd
from datetime import datetime
import re
import sys

from urllib.parse import urlparse

def clean_currency(value):
    """Convert currency string to float"""
    if pd.isna(value) or value == '':
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    # Remove $, commas, and extract number
    cleaned = re.sub(r'[\$,]', '', str(value))
    # Handle formats like "$1,566.68\n($1,706.68 including shipping)"
    match = re.search(r'([\d,]+\.?\d*)', cleaned)
    if match:
        return float(match.group(1))
    return 0.0

def parse_markup_percent(value):
    """Extract markup percentage from string like '$617.18 (65%)'"""
    if pd.isna(value) or value == '':
        return 0.0
    match = re.search(r'\((\d+(?:\.\d+)?)%\)', str(value))
    if match:
        return float(match.group(1)) / 100
    return 0.0

def extract_vendor_from_url(url):
    """
    Extract vendor name from URL
    e.g. 'https://www.target.com/...' -> 'Target'
         'https://juniperprintshop.com/...' -> 'Juniperprintshop'
    """
    if pd.isna(url) or url == '':
        return ''
    
    try:
        parsed = urlparse(str(url))
        hostname = parsed.netloc
        # Remove www.
        if hostname.startswith('www.'):
            hostname = hostname[4:]
        
        # Get the main domain part (e.g. 'target' from 'target.com')
        if '.' in hostname:
            vendor = hostname.split('.')[0]
        else:
            vendor = hostname
            
        return vendor.title()
    except:
        return ''

def get_row_data(project_name, proposal_number, houzz_row):
    """
    Extract and process data for a single row
    Returns a dictionary of all available data points
    """
    # Extract raw data
    item_name = houzz_row.get('Item', '')
    description = houzz_row.get('Description', '')
    qty = houzz_row.get('Qty', 1)
    cost = clean_currency(houzz_row.get('Cost', 0))
    markup_pct = parse_markup_percent(houzz_row.get('Markup', ''))
    markup_str = houzz_row.get('Markup', '') # Raw markup string
    unit_price = clean_currency(houzz_row.get('Unit Price', 0))
    room = houzz_row.get('Room', '')
    shipping = clean_currency(houzz_row.get('Shipping', 0))
    taxable = houzz_row.get('Taxable', '')
    website_link = houzz_row.get('Website Link', '')
    total = clean_currency(houzz_row.get('Total', 0))
    
    # helper for formatting
    fmt_currency = lambda x: f"${x:,.2f}" if x != 0 else ""
    
    # Calculate derived values
    markup_amt = cost * markup_pct if markup_pct > 0 else (unit_price - cost)
    subtotal = cost * qty
    pre_tax_amt = subtotal + (markup_amt * qty)
    
    # Determine prepaid tax (assume true if taxable is "Item")
    prepaid_tax = 'true' if taxable == 'Item' else 'false'
    
    # Create dictionary of available data
    # We map potential Grist column names to our values here
    data = {
        # Identity
        'Project': project_name,
        'Proposal': proposal_number,
        'imported project': project_name,
        
        # Core Item Data
        'Item': item_name,
        'Product': description, # Keeping for backward compat if column exists
        'Room': room if pd.notna(room) else '',
        'Vendor': extract_vendor_from_url(website_link), # Extracted from URL
        'URL': website_link,
        'Notes': description, # Redundant but useful if Product column is removed
        
        # Quantities & Costs
        'QTY': int(qty) if pd.notna(qty) else 1,
        'Unit COST': fmt_currency(cost),
        'Unit MSRP': fmt_currency(unit_price),
        'Cost': fmt_currency(cost),
        
        # Financials
        'Markup%': f"{markup_pct*100:.2f}%" if markup_pct > 0 else "",
        'Markup': fmt_currency(markup_amt * qty), # Total markup for the line
        'Shipping': fmt_currency(shipping),
        
        # Calculated Fields (Left blank for Grist to calculate)
        'Subtotal': '', 
        'Pre-Tax': '',
        'Total': '',
        'Tax': '',
        'Prepaid Tax?': '', # User requested to leave blank
        'Project Tax Rate': '', # User requested to leave blank
        'Prepaid Tax Amt': '',
        'Owed State Tax Amt': '',
        
        # Meta
        'Ordered': datetime.now().strftime('%m/%d/%Y'), # Default to today
        'Created': datetime.now().strftime('%Y-%m-%d %I:%M%p').lower(),
        'Created By': 'houzz_import',
        'Modified': datetime.now().strftime('%Y-%m-%d %I:%M%p').lower(),
        'Modified By': 'houzz_import',
        'Received': '',
        'In Houzz': 'true',
        
        # Fallback for removed columns if needed
        'Description': description
    }
    
    return data

def convert_houzz_to_grist(houzz_file, output_file=None, project_name="", proposal_number="", template_csv="ORDERS.csv"):
    """
    Convert Houzz export to Grist format using template CSV for column structure.
    If output_file is None, simply returns the DataFrame.
    """
    # Read Houzz export
    print(f"Reading Houzz export from: {houzz_file}")
    df_houzz = pd.read_excel(houzz_file)
    print(f"Found {len(df_houzz)} items in Houzz export")

    # Filter out parent rows if children exist
    # Rule: if we see 7.1, 7.2 etc, remove 7.
    # The '#' column contains these IDs.
    
    if '#' in df_houzz.columns:
        print("Applying parent row filtering...")
        # Get all IDs as strings
        ids = df_houzz['#'].astype(str).tolist()
        
        # Identify parents that have children
        parents_to_remove = set()
        for id_val in ids:
            if '.' in id_val:
                try:
                    # '7.1' -> parent is '7' or '7.0'
                    parts = id_val.split('.')
                    if len(parts) >= 2:
                        parent_id = parts[0]
                        parents_to_remove.add(parent_id)
                        # Also add with .0 just in case (e.g. 7.0)
                        parents_to_remove.add(f"{parent_id}.0") 
                except:
                    pass
        
        print(f"Parents to remove (because children found): {parents_to_remove}")
        
        # Filter the dataframe
        # Keep row IF its ID is NOT in the parents_to_remove list
        # We need to be careful with type matching (float vs str)
        
        original_count = len(df_houzz)
        df_houzz['#_str'] = df_houzz['#'].astype(str).apply(lambda x: x.replace('.0', '') if x.endswith('.0') and '.' not in x[:-2] else x)
        
        # More robust check: convert everything to standardized string
        # 7.0 -> "7"
        # 7.1 -> "7.1"
        def standardize_id(val):
            s = str(val)
            if s.endswith('.0'):
                return s[:-2]
            return s

        df_houzz['temp_id'] = df_houzz['#'].apply(standardize_id)
        
        # Re-calculate parents to remove based on standardized IDs
        parents_with_children = set()
        all_std_ids = df_houzz['temp_id'].tolist()
        
        for pid in all_std_ids:
             if '.' in pid:
                 parent = pid.split('.')[0]
                 parents_with_children.add(parent)
                 
        print(f"Detected parents with children: {parents_with_children}")

        # Filter
        # Remove if ID is in parents_with_children AND ID doesn't contain '.' (is a parent)
        df_houzz = df_houzz[~((df_houzz['temp_id'].isin(parents_with_children)) & (~df_houzz['temp_id'].str.contains(r'\.')))]
        
        new_count = len(df_houzz)
        print(f"Filtered {original_count - new_count} parent rows. New count: {new_count}")

    # Determine Output Columns
    target_columns = []
    try:
        if os.path.exists(template_csv):
            print(f"Reading column structure from {template_csv}...")
            # Read only header
            df_template = pd.read_csv(template_csv, nrows=0)
            target_columns = df_template.columns.tolist()
            print(f"Detected {len(target_columns)} columns: {target_columns}")
        else:
            print(f"Template file {template_csv} not found. Using default schema.")
            target_columns = [
                'Project', 'Proposal', 'Ordered', 'Item', 'Vendor', 'QTY', 'Unit COST', 
                'Markup%', 'Markup', 'Subtotal', 'Pre-Tax', 'Shipping', 'Total', 
                'Prepaid Tax?', 'Project Tax Rate', 'Tax', 'Prepaid Tax Amt', 
                'Created', 'Notes', 'Project_Project Tax Rate', 'Received', 'URL'
            ]
    except Exception as e:
        print(f"Error reading template: {e}. using default.")
        # Fallback to the full known schema so the user still gets a usable result
        # even if ORDERS.csv is locked or missing.
        target_columns = [
            'Project', 'Proposal', 'Ordered', 'Item', 'Vendor', 'QTY', 'Unit COST', 
            'Markup%', 'Markup', 'Subtotal', 'Pre-Tax', 'Shipping', 'Total', 
            'Prepaid Tax?', 'Project Tax Rate', 'Tax', 'Prepaid Tax Amt', 
            'Created', 'Notes', 'Project_Project Tax Rate', 'Received', 'URL'
        ]

    # Map Rows
    rows = []
    for idx, houzz_row in df_houzz.iterrows():
        # Get all potential data matches
        row_data = get_row_data(project_name, proposal_number, houzz_row)
        
        # Build the ordered row based on target columns
        ordered_row = {}
        for col in target_columns:
            # Smart mapping: Check direct match, then specific overrides
            if col in row_data:
                ordered_row[col] = row_data[col]
            else:
                # Handle special cases where column names might differ slightly
                ordered_row[col] = ''
                
        rows.append(ordered_row)
    
    df_grist = pd.DataFrame(rows)
    
    # Save to CSV
    if output_file:
        print(f"\nSaving {len(df_grist)} rows to: {output_file}")
        df_grist.to_csv(output_file, index=False)
        print("Conversion complete!")
    
    return df_grist

if __name__ == "__main__":
    import os
    
    # Default values
    houzz_file = "Paxman DDS_ Art_12_16_2025.xlsx"
    output_file = "houzz_import.csv"
    project_name = "Paxman DDS"
    proposal_number = "Art_12_16_2025"
    template_file = "ORDERS.csv"
    
    # Allow command-line arguments
    if len(sys.argv) > 1:
        houzz_file = sys.argv[1]
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    if len(sys.argv) > 3:
        project_name = sys.argv[3]
    if len(sys.argv) > 4:
        proposal_number = sys.argv[4]
    
    print("=" * 60)
    print("Houzz to Grist Converter (Dynamic Schema)")
    print("=" * 60)
    print(f"Input file: {houzz_file}")
    print(f"Output file: {output_file}")
    print(f"Template: {template_file}")
    print("=" * 60)
    
    result = convert_houzz_to_grist(houzz_file, output_file, project_name, proposal_number, template_file)

