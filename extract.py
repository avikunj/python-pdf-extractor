import os
import re
import PyPDF2
import pandas as pd
import fitz  # PyMuPDF
import pdfplumber

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

def extract_personal_details(text, pdf_path):
    # Define regex patterns for personal details
    patterns = {
        "Company Name": r"\b([A-Z\s]+(?:\s+(?:COMPANY|ENTERPRISES|LIMITED|PRIVATE LIMITED))?)([\s\S]{0,6})\b",
        "Merchant Name": r"(?:Merchant Name)[:\s]+(\b[A-Z]+(?:\s[A-Z]+)+\b)",
        "Email ID": r"(?:Email ID)[:\s]+([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})",
        "Phone No": r"(?:Phone No)[:\s]+(?:\+91[-\s]?)?(\d{10})",
        "Sample Name": r"(?:SAMPLE NAME)[:\s]+(.*)",
        "Test Name": r"(?:TEST NAME|TEST NAME-)[:\s]+(.*)",
    }

    # Extract details using regex patterns
    details = {}
    for label, pattern in patterns.items():
        match = re.search(pattern, text)        
        if match and label != "Company Name":
            details[label] = match.group(1)
        elif match and label == "Company Name":
            details[label] = extract_text_from_coordinates(pdf_path, 0, 300, 100, 580, 300)    
        else:
            details[label] = None  # Fill with None if not found

    return details

def process_pdfs(folder_path, output_excel):
    # List to hold data for each PDF
    data = []
    
    # Loop over all PDF files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".PDF"):
            pdf_path = os.path.join(folder_path, filename)
            with open(pdf_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                
                # Concatenate text from all pages
                text = ""
                for page in reader.pages:
                    text += page.extract_text() or ""
                
                # Extract details from the text
                details = extract_personal_details(text, pdf_path)
                print(details)
                details["File"] = filename  # Add filename for reference
                data.append(details)
    
    # Create a DataFrame from the collected data
    df = pd.DataFrame(data)
    
    # Save DataFrame to Excel
    df.to_excel(output_excel, index=False)
    print(f"Data saved to {output_excel}")

# Usage
folder_path = os.getcwd()  # Replace with the path to your PDF folder
output_excel = "data.xlsx"     # Name of the output Excel file
process_pdfs(folder_path, output_excel)
