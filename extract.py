import os
import re
import PyPDF2
import pandas as pd
import fitz  # PyMuPDF
import pdfplumber
import time

# Function to update Excel file
def update_excel(excel_path, data):
    # Load existing Excel file or create a new DataFrame
    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path)
    else:
        df = pd.DataFrame(columns=["Company Name", "City", "State", "Merchant Name", "Email ID", "Phone No", "Net Amount INR", "Sample Name", "Test Name" ])

    # Append new data
    new_data = pd.DataFrame(data)
    df = pd.concat([df, new_data], ignore_index=True)
    
    # Save back to Excel
    df.to_excel(excel_path, index=False)


# Function to extract text from a specific region using coordinates
def extract_text_from_coordinates(pdf_path, page_number, x0, top, x1, bottom):
    with pdfplumber.open(pdf_path) as pdf:
        # Get the page by page_number (zero-indexed)
        page = pdf.pages[page_number]
        
        # Define the bounding box (x0, top, x1, bottom)
        crop_box = (x0, top, x1, bottom)
        
        # Crop the page to the region and extract text
        cropped_page = page.within_bbox(crop_box)
        text = cropped_page.extract_text()
        
        return text

def extract_personal_details(pdf_file):

        with open(pdf_file, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            
            # Concatenate text from all pages
            text = ""
            for page in reader.pages:
                text += page.extract_text() or ""

        # Define regex patterns for personal details
        patterns = {
            "Company Name": r"\b([A-Z\s]+(?:\s+(?:COMPANY|ENTERPRISES|LIMITED|PRIVATE LIMITED))?)([\s\S]{0,6})\b",
            "City":r'\b([A-Za-z]+)\s+([A-Za-z]+)\s+\(\d{2}\)\s+\d{6}\b',
            "State":r'\b([A-Za-z]+)\s+\(\d{2}\)\s+\d{6}\b',
            "Merchant Name": r"(?:Merchant Name*)[:\s]+(.*)",
            "Email ID": r"(?:Email ID)[:\s]+([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})",
            "Phone No": r"(?:Phone No)[:\s]+(?:\+91[-\s]?)?(\d{10})",
            "Sample Name": r"(?:SAMPLE NAME)[:\s]+(.*)",
            "Test Name": r"(?:TEST NAME|TEST NAME-)[:\s]+(.*)",
            "Net Amount INR":r"(?:Net Amount INR|Net Amount  INR|Net Amount  INR|Net Amount   INR|Net Amount   INR )[:\s]+(.*)",
        }

        # Extract details using regex patterns
        details = {}
        for label, pattern in patterns.items():
            if label == "Sample Name":
                # details[label] = "None"
                matching_records = []
                matches = re.findall(pattern, text)
                if matches:
                    # Append each match as a dictionary to the list, labeled by the pattern label
                    matching_records.extend([match.strip() for match in matches])
                    details[label] = ", ".join(matching_records)
            elif label == "Test Name":
                # details[label] = "None"
                matching_records = []
                matches = re.findall(pattern, text)
                if matches:
                    # Append each match as a dictionary to the list, labeled by the pattern label
                    matching_records.extend([match.strip() for match in matches])
                    details[label] = ", ".join(matching_records)
            else:     
                match = re.search(pattern, text)        
                if match and label != "Company Name" and label != "Net Amount INR":
                    details[label] = match.group(1)
                elif label == "Net Amount INR" and match:
                    details[label] = re.sub(r'\b(Sum|of|Tax|INR)\b', '', match.group(1).strip())    
                elif match and label == "Company Name":
                    details[label] = extract_text_from_coordinates(pdf_file, 0, 300, 100, 580, 300)    
                else:
                    details[label] = None  # Fill with None if not found

        return details

# Main function to process PDF batch
def process_pdfs_in_batch(pdf_folder, excel_path, batch_size=20):
    all_data = []
    batch_count = 0

    for filename in os.listdir(pdf_folder):
        if filename.endswith(".PDF"):
            pdf_path = os.path.join(pdf_folder, filename)
            data = extract_personal_details(pdf_path)
            all_data.append(data)
            batch_count += 1
            
            # Process the batch if it reaches the specified batch size
            if batch_count == batch_size:
                # Update Excel with the current batch
                print(all_data)
                update_excel(excel_path, all_data)
                all_data = []  # Clear the batch data after updating
                batch_count = 0  # Reset the batch count for the next batch
                time.sleep(5)
    
    # Process any remaining data if the total files are not a multiple of batch_size
    if all_data:
        update_excel(excel_path, all_data)
# Usage
pdf_folder = os.getcwd()
excel_path = "output.xlsx"
process_pdfs_in_batch(pdf_folder, excel_path)
