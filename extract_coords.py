import os
from PIL import Image
import pytesseract
import re
import openpyxl

# Path to your Tesseract executable (modify as needed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Chirayu Chawande\AppData\Local\Programs\Tesseract-OCR"
# Regex pattern for extracting coordinates
lat_lon_pattern = r"Lat\s+([0-9]+\.[0-9]+)°?\s+Long\s+([0-9]+\.[0-9]+)°?"

def extract_coordinates(image_path):
    try:
        image = Image.open(image_path)
        text = pytesseract.image_to_string(image)
        match = re.search(lat_lon_pattern, text)
        if match:
            latitude = match.group(1)
            longitude = match.group(2)
            return f"{latitude}, {longitude}"
        else:
            return "Coordinates not found"
    except Exception as e:
        return f"Error: {str(e)}"

def process_images_in_folder(root_folder):
    for root, dirs, files in os.walk(root_folder):
        if files:
            folder_name = os.path.basename(root_folder)
            subfolder_name = os.path.basename(root)
            excel_filename = f"{subfolder_name}_{folder_name}.xlsx"
            
            # Create workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Photo No.", "Lat, Long"])
            
            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg', '.png')):
                    image_path = os.path.join(root, file)
                    coords = extract_coordinates(image_path)
                    ws.append([file, coords])
                    print(f"Processed: {file} → {coords}")
            
            # Save Excel file
            output_path = os.path.join(root, excel_filename)
            wb.save(output_path)
            print(f"✅ Saved Excel: {output_path}")