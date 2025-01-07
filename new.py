import fitz  
import pandas as pd 
from PIL import Image
import io
import pytesseract

# Define paths
pdf_path = 'sample-pdf-with-images.pdf'  # Ensure this path is correct
output_csv = 'form_data.csv'
output_excel = 'form_data.xlsx'

# Ensure Tesseract-OCR is installed and configured
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'  # Update for your installation path

def extract_form_data(pdf_path):
    """Extract form data from the PDF."""
    doc = fitz.open(pdf_path)
    data = {}

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        print(f"Processing page {page_num + 1}")

        for field in page.widgets():
            field_name = field.field_name
            field_value = field.field_value
            data[field_name] = field_value
            print(f"Field Name: {field_name}, Field Value: {field_value}")

    return data

def extract_images_from_pdf(pdf_path):
    """Extract images from the PDF."""
    doc = fitz.open(pdf_path)
    images = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        image_list = page.get_images(full=True)
        print(f"Found {len(image_list)} image(s) on page {page_num + 1}")

        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image = Image.open(io.BytesIO(image_bytes))
            images.append(image)

    return images

def ocr_image(image, lang='eng'):
    """Perform OCR on an image."""
    return pytesseract.image_to_string(image, lang=lang)

def save_data_to_csv(data, output_file):
    """Save extracted data to a CSV file."""
    if data:
        df = pd.DataFrame([data])
        df.to_csv(output_file, index=False)
    else:
        print("No data to save to CSV.")

def save_data_to_excel(data, output_file):
    """Save extracted data to an Excel file."""
    if data:
        df = pd.DataFrame([data])
        df.to_excel(output_file, index=False)
    else:
        print("No data to save to Excel.")

def main():
    """Main function to run the extraction process."""
    # Extract form data
    form_data = extract_form_data(pdf_path)
    if form_data:
        print("Form Data Extracted:")
        print(form_data)
    else:
        print("No form data found.")

    # Extract and process images
    images = extract_images_from_pdf(pdf_path)
    if images:
        ocr_results = [ocr_image(img) for img in images]
        for idx, text in enumerate(ocr_results):
            print(f"OCR Result for Image {idx + 1}:")
            print(text)
    else:
        print("No images found.")

    # Save extracted data to CSV and Excel
    save_data_to_csv(form_data, output_csv)
    save_data_to_excel(form_data, output_excel)
    print(f"Data saved to {output_csv} and {output_excel}")

if __name__ == "__main__":
    main()
    



