import os
from PIL import Image
import pytesseract
import re
import openpyxl

# Path to your Tesseract executable (modify as needed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Chirayu Chawande\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# Regex pattern for extracting coordinates
lat_lon_pattern = r"Lat\s+([0-9]+\.[0-9]+)°?\s+Long\s+([0-9]+\.[0-9]+)°?"

def extract_coordinates(image_path):
    """Extract latitude and longitude text from an image using OCR."""
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
    """Process all images in a folder (and subfolders) and save results to Excel."""
    # Create output directory
    output_dir = os.path.join(root_folder, "output")
    os.makedirs(output_dir, exist_ok=True)

    for root, dirs, files in os.walk(root_folder):
        if files:
            folder_name = os.path.basename(root_folder)
            subfolder_name = os.path.basename(root)
            excel_filename = f"{subfolder_name}_{folder_name}.xlsx"
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Coordinates"
            ws.append(["Photo No.", "Lat, Long"])
            
            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg', '.png')):
                    image_path = os.path.join(root, file)
                    coords = extract_coordinates(image_path)
                    ws.append([file, coords])
                    print(f"Processed: {file} → {coords}")
            
            # Save Excel to output folder
            output_path = os.path.join(output_dir, excel_filename)
            wb.save(output_path)
            print(f"✅ Saved Excel: {output_path}")

    print("\n🎉 All done! Check your 'output' folder for results.")


if __name__ == "__main__":
    root_folder = input("📁 Enter the path to your image folder: ").strip()
    if os.path.isdir(root_folder):
        process_images_in_folder(root_folder)
    else:
        print("❌ Invalid folder path. Please check and try again.")
