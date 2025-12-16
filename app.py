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

    col1, col2 = st.columns(2)
    
    # Input Fields
    with col1:
        project_name = st.text_input("Project Name", placeholder="e.g. Paxman DDS")
    
    with col2:
        proposal_number = st.text_input("Proposal #", placeholder="e.g. Art_12_16")

    # File Uploader
    uploaded_file = st.file_uploader("Upload Houzz Excel File", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        if not project_name:
            st.warning("⚠️ Please enter a Project Name first.")
        elif not proposal_number:
            st.warning("⚠️ Please enter a Proposal Number first.")
        else:
            try:
                # Run conversion logic
                # We assume convert_houzz_to_grist can accept a file-like object from streamlit
                with st.spinner("Converting..."):
                     # Save processed structure to session state to prevent reload loops
                     df_result = convert_houzz_to_grist(
                         uploaded_file, 
                         output_file=None, # Don't save to file automatically
                         project_name=project_name, 
                         proposal_number=proposal_number
                     )
                
                st.success("Conversion Successful!")
                
                # Preview
                st.subheader("Data Preview")
                st.dataframe(df_result, use_container_width=True)
                
                # Convert DF to CSV text for downloading/copying
                csv_data = df_result.to_csv(index=False).encode('utf-8')
                csv_text = df_result.to_csv(index=False)
                
                st.download_button(
                    label="⬇️ Download CSV",
                    data=csv_data,
                    file_name='houzz_import.csv',
                    mime='text/csv',
                    type="primary"
                )

            except Exception as e:
                st.error(f"Error during conversion: {e}")

if __name__ == "__main__":
    main()
