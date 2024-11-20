import sys
from tabula import read_pdf
import pandas as pd

def extract_tables(pdf_path, output_path):
    try:
        # Extract tables from PDF
        tables = read_pdf(pdf_path, pages="all", multiple_tables=True, lattice=True)

        # Write tables to an Excel file
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            for i, table in enumerate(tables):
                table.to_excel(writer, index=False, sheet_name=f"Table_{i+1}")

        print(f"Success: Tables extracted to {output_path}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    pdf_path = sys.argv[1]
    output_path = sys.argv[2]
    extract_tables(pdf_path, output_path)


