import email
import os
import pandas as pd
import openpyxl
import re
import platform
from pathlib import Path
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

def get_unique_filename(file_path):
    """
    Generate a unique filename by adding _n suffix if file already exists
    """
    path = Path(file_path)
    if not path.exists():
        return str(path)
    
    base_name = path.stem
    extension = path.suffix
    parent = path.parent
    counter = 1
    
    while True:
        new_name = f"{base_name}_{counter}{extension}"
        new_path = parent / new_name
        if not new_path.exists():
            return str(new_path)
        counter += 1

def extract_excel_from_eml(eml_file_path, output_dir="temp_attachments"):
    """
    Extract Excel attachments from an EML file (cross-platform compatible)
    """
    # Use pathlib for cross-platform path handling
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    print(f"  Extracting attachments to: {output_path}")
    
    with open(eml_file_path, 'rb') as f:
        msg = email.message_from_bytes(f.read())
    
    excel_files = []
    
    for part in msg.walk():
        if part.get_content_disposition() == 'attachment':
            filename = part.get_filename()
            if filename:
                # Decode base64 encoded filename if needed
                if filename.startswith('=?') and filename.endswith('?='):
                    decoded_header = email.header.decode_header(filename)[0]
                    if decoded_header[1]:
                        filename = decoded_header[0].decode(decoded_header[1])
                    else:
                        filename = decoded_header[0]
                
                if filename.endswith('.xlsx') or filename.endswith('.xls'):
                    # Use pathlib for cross-platform path handling
                    file_path = output_path / filename
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    excel_files.append(str(file_path))
                    print(f"  Extracted Excel file: {filename}")
    
    return excel_files

def process_sales_excel(excel_file_path):
    """
    Process the sales Excel file according to the specific rules (cross-platform compatible)
    """
    excel_path = Path(excel_file_path)
    
    # Load workbook to extract dates from row 1
    wb = openpyxl.load_workbook(str(excel_path))
    ws = wb.active
    
    # Extract dates from row 1
    title_cell = ws.cell(row=1, column=1).value
    print(f"  Title cell content: {title_cell}")
    
    # Extract dates using regex pattern dd/MM/yyyy
    date_pattern = r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})'
    date_match = re.search(date_pattern, str(title_cell))
    
    if not date_match:
        raise ValueError("Could not find date range in the expected format dd/MM/yyyy - dd/MM/yyyy")
    
    start_date_str = date_match.group(1)
    end_date_str = date_match.group(2)
    
    print(f"  Found dates: {start_date_str} to {end_date_str}")
    
    # Parse dates
    start_date = datetime.strptime(start_date_str, '%d/%m/%Y')
    end_date = datetime.strptime(end_date_str, '%d/%m/%Y')
    
    # Read the Excel file starting from row 3 (which becomes the header)
    df = pd.read_excel(str(excel_path), header=2)  # 0-indexed, so row 3 becomes header
    
    print(f"  Original dataframe shape: {df.shape}")
    
    # Insert two new columns at the beginning
    df.insert(0, 'Period_Start', start_date_str)
    df.insert(1, 'Period_End', end_date_str)
    
    # Generate filename with date format
    start_formatted = start_date.strftime('%Y%m%d')
    end_formatted = end_date.strftime('%Y%m%d')
    output_filename = f"sales_{start_formatted}_{end_formatted}.csv"
    
    return df, output_filename

def process_eml_file(eml_file_path, csv_output_dir, temp_dir="temp_processing"):
    """
    Process a single EML file: extract Excel, process it, and save as CSV
    """
    eml_path = Path(eml_file_path)
    csv_dir = Path(csv_output_dir)
    temp_path = Path(temp_dir)
    
    print(f"\nProcessing: {eml_path.name}")
    
    # Create temporary directory for this file
    file_temp_dir = temp_path / eml_path.stem
    file_temp_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # Extract Excel attachments
        excel_files = extract_excel_from_eml(str(eml_path), str(file_temp_dir))
        
        if not excel_files:
            print(f"  ❌ No Excel attachments found in {eml_path.name}")
            return False
        
        processed_files = []
        
        # Process each Excel file
        for excel_file in excel_files:
            try:
                print(f"  Processing Excel: {Path(excel_file).name}")
                
                # Process the Excel file
                df, csv_filename = process_sales_excel(excel_file)
                
                # Create output path in csv directory
                csv_output_path = csv_dir / csv_filename
                
                # Get unique filename if file already exists
                unique_csv_path = get_unique_filename(str(csv_output_path))
                
                # Save to CSV
                df.to_csv(unique_csv_path, index=False)
                
                print(f"  ✅ Saved: {Path(unique_csv_path).name}")
                print(f"  📊 Shape: {df.shape}")
                
                processed_files.append(unique_csv_path)
                
            except Exception as e:
                print(f"  ❌ Error processing {Path(excel_file).name}: {str(e)}")
        
        return len(processed_files) > 0
        
    except Exception as e:
        print(f"  ❌ Error processing {eml_path.name}: {str(e)}")
        return False
    
    finally:
        # Clean up temporary files
        try:
            import shutil
            if file_temp_dir.exists():
                shutil.rmtree(file_temp_dir)
        except Exception as e:
            print(f"  ⚠️ Warning: Could not clean up temp directory: {e}")

def main():
    """
    Main function to process all EML files in the eml subfolder
    """
    print(f"🚀 Batch Email Processor - Running on {platform.system()}")
    print(f"📁 Working directory: {Path.cwd()}")
    
    # Define directories
    eml_dir = Path("eml")
    csv_dir = Path("csv")
    temp_dir = Path("temp_processing")
    
    # Check if directories exist
    if not eml_dir.exists():
        print(f"❌ Error: {eml_dir} directory not found")
        return
    
    # Create csv directory if it doesn't exist
    csv_dir.mkdir(exist_ok=True)
    print(f"📤 Output directory: {csv_dir.absolute()}")
    
    # Find all EML files
    eml_files = list(eml_dir.glob("*.eml"))
    
    if not eml_files:
        print(f"❌ No EML files found in {eml_dir}")
        return
    
    print(f"📧 Found {len(eml_files)} EML file(s) to process:")
    for eml_file in eml_files:
        print(f"   - {eml_file.name}")
    
    # Process each EML file
    successful = 0
    failed = 0
    
    for eml_file in eml_files:
        if process_eml_file(str(eml_file), str(csv_dir), str(temp_dir)):
            successful += 1
        else:
            failed += 1
    
    # Clean up main temp directory
    try:
        import shutil
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
    except Exception as e:
        print(f"⚠️ Warning: Could not clean up main temp directory: {e}")
    
    # Summary
    print(f"\n📋 Processing Summary:")
    print(f"   ✅ Successful: {successful}")
    print(f"   ❌ Failed: {failed}")
    print(f"   📁 Output directory: {csv_dir.absolute()}")
    
    if successful > 0:
        csv_files = list(csv_dir.glob("*.csv"))
        print(f"   📄 Generated CSV files: {len(csv_files)}")
        for csv_file in csv_files:
            print(f"      - {csv_file.name}")

if __name__ == "__main__":
    main()