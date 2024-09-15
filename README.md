# Roobini
# To install required libraries: pip3 install pandas fastpunct openpyxl
import pandas as pd
from fastpunct import FastPunct
import logging

# Setup logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Load your Excel file
file_path = 'punctuation_corrections.xlsx'
df = pd.read_excel(file_path)

# Initialize fastPunct
fastpunct = FastPunct()

# Function to correct text
def correct_text(text):
    return fastpunct.punct(text)

# Apply correction to all cells in the DataFrame
total_rows = df.shape[0]  # Total number of rows
logging.info(f"Starting correction for {total_rows} rows.")

for i, (index, row) in enumerate(df.iterrows()):
    df.iloc[index] = row.apply(lambda cell: correct_text(str(cell)) if isinstance(cell, str) else cell)
    if (i + 1) % 10 == 0:  # Log progress every 10 rows
        logging.info(f"Processed {i + 1} out of {total_rows} rows.")

# Save the corrected DataFrame back to an Excel file
output_file_path = 'corrected_excel_file.xlsx'
df.to_excel(output_file_path, index=False)

logging.info(f"Corrected file saved to {output_file_path}")

# To run: python3 correct_excel.py

