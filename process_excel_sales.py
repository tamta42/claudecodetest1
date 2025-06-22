import pandas as pd
import openpyxl
import re
from datetime import datetime
import os
import platform
from pathlib import Path

def process_sales_excel(excel_file_path):
    """
    Process the sales Excel file according to the specific rules (cross-platform compatible):
    - Row 1: Report title with dates (ignore)
    - Row 2: Company info (ignore) 
    - Row 3: Column headers
    - Row 4+: Data rows
    """
    
    print(f"Processing on {platform.system()} platform")
    excel_path = Path(excel_file_path)
    print(f"Excel file path: {excel_path.absolute()}")
    
    # Load workbook to extract dates from row 1
    wb = openpyxl.load_workbook(str(excel_path))
    ws = wb.active
    
    # Extract dates from row 1
    title_cell = ws.cell(row=1, column=1).value
    print(f"Title cell content: {title_cell}")
    
    # Extract dates using regex pattern dd/MM/yyyy
    date_pattern = r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})'
    date_match = re.search(date_pattern, str(title_cell))
    
    if not date_match:
        raise ValueError("Could not find date range in the expected format dd/MM/yyyy - dd/MM/yyyy")
    
    start_date_str = date_match.group(1)
    end_date_str = date_match.group(2)
    
    print(f"Found dates: {start_date_str} to {end_date_str}")
    
    # Parse dates
    start_date = datetime.strptime(start_date_str, '%d/%m/%Y')
    end_date = datetime.strptime(end_date_str, '%d/%m/%Y')
    
    # Read the Excel file starting from row 3 (which becomes the header)
    df = pd.read_excel(str(excel_path), header=2)  # 0-indexed, so row 3 becomes header
    
    print(f"Original dataframe shape: {df.shape}")
    print(f"Column names: {list(df.columns)}")
    
    # Insert two new columns at the beginning
    df.insert(0, 'Period_Start', start_date_str)
    df.insert(1, 'Period_End', end_date_str)
    
    # Generate output filename using pathlib for cross-platform compatibility
    start_formatted = start_date.strftime('%Y%m%d')
    end_formatted = end_date.strftime('%Y%m%d')
    output_filename = f"sales_{start_formatted}_{end_formatted}.csv"
    output_path = Path(output_filename)
    
    print(f"Output will be saved to: {output_path.absolute()}")
    
    # Save to CSV using pathlib
    df.to_csv(str(output_path), index=False)
    
    print(f"Processed file saved as: {output_filename}")
    print(f"Final dataframe shape: {df.shape}")
    print(f"First few rows:")
    print(df.head())
    
    return output_filename

def main():
    # Use pathlib for cross-platform path handling
    excel_file = Path("extracted_attachments") / "AHEAD Supplier Sales v2.0.xlsx"
    
    print(f"Running on {platform.system()}")
    print(f"Current working directory: {Path.cwd()}")
    print(f"Looking for Excel file at: {excel_file.absolute()}")
    
    if not excel_file.exists():
        print(f"Error: {excel_file} not found")
        # Try to find any Excel files in the directory
        extracted_dir = Path("extracted_attachments")
        if extracted_dir.exists():
            excel_files = list(extracted_dir.glob("*.xlsx")) + list(extracted_dir.glob("*.xls"))
            if excel_files:
                print(f"Available Excel files: {[f.name for f in excel_files]}")
        return
    
    try:
        output_file = process_sales_excel(str(excel_file))
        print(f"\n✅ Successfully processed Excel file!")
        print(f"📄 Output: {output_file}")
    except Exception as e:
        print(f"❌ Error processing file: {str(e)}")

if __name__ == "__main__":
    main()