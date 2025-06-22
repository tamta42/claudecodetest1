import email
import os
import pandas as pd
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

def extract_excel_from_eml(eml_file_path, output_dir="extracted_attachments"):
    """
    Extract Excel attachments from an EML file
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
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
                    file_path = os.path.join(output_dir, filename)
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    excel_files.append(file_path)
                    print(f"Extracted Excel file: {filename}")
    
    return excel_files

def excel_to_csv(excel_file_path, csv_output_path=None):
    """
    Convert Excel file to CSV (first sheet)
    """
    if csv_output_path is None:
        base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
        csv_output_path = f"{base_name}.csv"
    
    # Read the Excel file (first sheet by default)
    df = pd.read_excel(excel_file_path)
    
    # Save as CSV
    df.to_csv(csv_output_path, index=False)
    print(f"Saved CSV file: {csv_output_path}")
    
    return csv_output_path

def main():
    eml_file = "email1.eml"
    
    if not os.path.exists(eml_file):
        print(f"Error: {eml_file} not found in current directory")
        return
    
    print(f"Processing {eml_file}...")
    
    # Extract Excel attachments
    excel_files = extract_excel_from_eml(eml_file)
    
    if not excel_files:
        print("No Excel attachments found in the email")
        return
    
    # Convert each Excel file to CSV
    for excel_file in excel_files:
        try:
            csv_file = excel_to_csv(excel_file)
            print(f"Successfully converted {excel_file} to {csv_file}")
        except Exception as e:
            print(f"Error converting {excel_file}: {str(e)}")

if __name__ == "__main__":
    main()