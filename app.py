import streamlit as st
import pandas as pd
import os
import sys
# Import our logic (assuming it's in the same directory)
from houzz_to_grist import convert_houzz_to_grist

def main():
    st.set_page_config(page_title="Houzz to Grist Converter", page_icon="📝", layout="wide")

    st.title("Houzz to Grist Converter")
    st.markdown("Convert Houzz Excel Proposals to Grist Import Format instantly.")

    # Input Fields
    proposal_number = st.text_input("Proposal #", placeholder="e.g. Art_12_16")

    # File Uploader
    uploaded_file = st.file_uploader("Upload Houzz Excel File", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        if not proposal_number:
            st.warning("⚠️ Please enter a Proposal Number first.")
        else:
            try:
                # Run conversion logic
                with st.spinner("Converting..."):
                     # Save processed structure to session state to prevent reload loops
                     df_result = convert_houzz_to_grist(
                         uploaded_file, 
                         output_file=None, # Don't save to file automatically
                         proposal_number=proposal_number
                     )
                
                st.success("Conversion Successful!")
                
                # Preview
                st.subheader("Data Preview")
                st.dataframe(df_result, use_container_width=True)

                # --- Download Section ---
                col_dl, col_copy = st.columns(2)

                with col_dl:
                    # Convert DF to CSV text for downloading
                    csv_data = df_result.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="⬇️ Download CSV",
                        data=csv_data,
                        file_name='houzz_import.csv',
                        mime='text/csv',
                        type="primary"
                    )

                # --- Copy Section ---
                with col_copy:
                    st.markdown("### Copy for Grist")
                    st.markdown("Click the copy icon in the top right of the box below to copy data without headers.")
                    
                    # Convert to TSV (Tab Separated Values) string without index and header
                    # This is usually best for pasting into spreadsheet-like views
                    tsv_data = df_result.to_csv(index=False, header=False, sep='\t')
                    st.code(tsv_data, language="text")

            except Exception as e:
                st.error(f"Error during conversion: {e}")

if __name__ == "__main__":
    main()
