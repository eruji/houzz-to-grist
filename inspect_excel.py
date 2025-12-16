import pandas as pd

file_path = "c:/GIT/houzz to grist/Paxman DDS_ Art_12_16_2025.xlsx"
try:
    df = pd.read_excel(file_path)
    print("Columns:")
    print(df.columns.tolist())
    print("\nFirst 3 rows:")
    print(df.head(3).to_string())
except Exception as e:
    print(f"Error reading excel file: {e}")
