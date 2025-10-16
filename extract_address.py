import os
from PIL import Image
import pytesseract
import openpyxl

# Path to your Tesseract executable (modify as needed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Chirayu Chawande\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

def extract_text_from_region(image_path, crop_box=None):
    """
    Extract text from a specific region of an image using OCR.
    
    Args:
        image_path: Path to the image file
        crop_box: Tuple of (left, top, right, bottom) in pixels or percentage
                  If None, uses default region for address extraction
    
    Returns:
        Extracted text as string
    """
    try:
        image = Image.open(image_path)
        width, height = image.size
        
        # Default crop box - middle-right portion where address appears
        # (above "Agra" text, to the right of map)
        if crop_box is None:
            # Extract from middle-right region
            left = int(width * 0.35)   # Start from 35% from left
            top = int(height * 0.55)   # Start from 55% from top
            right = width              # Full right side
            bottom = int(height * 0.75) # End at 75% down
            crop_box = (left, top, right, bottom)
        
        # Crop the image to the specified region
        cropped_image = image.crop(crop_box)
        
        # Optional: Save cropped image for debugging
        # cropped_image.save("debug_crop.png")
        
        # Extract text from cropped region
        text = pytesseract.image_to_string(cropped_image, config='--psm 6')
        
        # Clean up the text
        text = text.strip()
        text = ' '.join(text.split())  # Remove extra whitespace
        
        return text if text else "No text found"
        
    except Exception as e:
        return f"Error: {str(e)}"

def process_images_in_folder(root_folder, crop_box=None):
    """
    Process all images in a folder (and subfolders) and save results to Excel.
    
    Args:
        root_folder: Root directory containing images
        crop_box: Optional tuple of (left, top, right, bottom) for custom crop region
    """
    # Create output directory
    output_dir = os.path.join(root_folder, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    for root, dirs, files in os.walk(root_folder):
        # Filter image files
        image_files = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
        
        if image_files:
            folder_name = os.path.basename(root_folder)
            subfolder_name = os.path.basename(root)
            excel_filename = f"{subfolder_name}_{folder_name}.xlsx"
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Extracted Text"
            ws.append(["Photo No.", "Extracted Text"])
            
            for file in image_files:
                image_path = os.path.join(root, file)
                extracted_text = extract_text_from_region(image_path, crop_box)
                ws.append([file, extracted_text])
                print(f"Processed: {file}")
                print(f"  → {extracted_text}\n")
            
            # Save Excel to output folder
            output_path = os.path.join(output_dir, excel_filename)
            wb.save(output_path)
            print(f"✅ Saved Excel: {output_path}\n")
    
    print("\n🎉 All done! Check your 'output' folder for results.")

def adjust_crop_region():
    """
    Interactive function to help adjust the crop region.
    Tests on a single image to find the right crop box.
    """
    test_image = input("📷 Enter path to a test image: ").strip()
    
    if not os.path.isfile(test_image):
        print("❌ Invalid file path.")
        return None
    
    image = Image.open(test_image)
    width, height = image.size
    print(f"\n📐 Image dimensions: {width}x{height} pixels")
    
    print("\n💡 Crop region options:")
    print("1. Middle-right area (default) - for address above 'Agra' text")
    print("2. Center middle-right area (Working)")
    print("3. Full middle section")
    print("4. Custom percentage")
    print("5. Custom pixel values")
    
    choice = input("\nSelect option (1-5): ").strip()
    
    if choice == "1":
        # Middle-right: 35-100% horizontal, 55-75% vertical
        crop_box = (int(width * 0.35), int(height * 0.55), width, int(height * 0.75))
    elif choice == "2":
        # center middle-right: 35-100% horizontal, 75-80% vertical
        crop_box = (int(width * 0.35), int(height * 0.75), width, int(height * 0.80))
    elif choice == "3":
        # Full middle: 0-100% horizontal, 50-75% vertical
        crop_box = (0, int(height * 0.50), width, int(height * 0.75))
    elif choice == "4":
        percentage = float(input("Enter percentage from bottom (e.g., 15): "))
        crop_box = (0, int(height * (1 - percentage/100)), width, height)
    elif choice == "5":
        left = int(input(f"Left (0-{width}): "))
        top = int(input(f"Top (0-{height}): "))
        right = int(input(f"Right (0-{width}): "))
        bottom = int(input(f"Bottom (0-{height}): "))
        crop_box = (left, top, right, bottom)
    else:
        print("Invalid choice, using default.")
        crop_box = None
    
    # Test extraction
    print("\n🔍 Testing extraction...")
    result = extract_text_from_region(test_image, crop_box)
    print(f"\n📝 Extracted text:\n{result}")
    
    # Save cropped image for visual verification
    img = Image.open(test_image)
    if crop_box:
        cropped = img.crop(crop_box)
        debug_path = "debug_cropped_region.png"
        cropped.save(debug_path)
        print(f"\n💾 Saved cropped region to: {debug_path}")
    
    return crop_box

if __name__ == "__main__":
    print("=" * 60)
    print("📸 Image Text Extraction Tool")
    print("=" * 60)
    
    print("\nOptions:")
    print("1. Process images with default settings")
    print("2. Adjust crop region first (recommended for first use)")
    
    option = input("\nSelect option (1-2): ").strip()
    
    crop_box = None
    
    if option == "2":
        crop_box = adjust_crop_region()
        proceed = input("\n✅ Use this crop region? (y/n): ").strip().lower()
        if proceed != 'y':
            print("Exiting...")
            exit()
    
    root_folder = input("\n📁 Enter the path to your image folder: ").strip()
    
    if os.path.isdir(root_folder):
        process_images_in_folder(root_folder, crop_box)
    else:
        print("❌ Invalid folder path. Please check and try again.")