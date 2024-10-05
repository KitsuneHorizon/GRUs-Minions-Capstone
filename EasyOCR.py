import warnings
warnings.filterwarnings('ignore')

import os
import time
import easyocr
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from tqdm import tqdm  # Import tqdm for progress bar
from spellchecker import SpellChecker
import string
import re

# Preprocessing function to clean and prepare words for spell checking
def preprocess_words(text):
    # Remove punctuation using regex
    text = re.sub(r'[^\w\s]', '', text)
    # Convert to lowercase
    words = text.lower().split()
    return words

# Function to process images and extract text with verification
def extract_text_from_images(directory):
    # Initialize the OCR readers
    reader_simplified = easyocr.Reader(['en', 'ch_sim'], verbose=False)
    reader_traditional = easyocr.Reader(['en', 'ch_tra'], verbose=False)

    # Initialize spell checker
    spell = SpellChecker(language='en')
    custom_words = ['cas']
    spell.word_frequency.load_words(custom_words)

    # Initialize list to store data for the excel file
    data = []
    total_images = 0
    images_with_text = 0
    images_without_text = 0
    failed_extractions = 0
    low_confidence_count = 0
    spell_errors_count = 0
    total_extracted_elements = 0  # Counter for total extracted elements

    # Supported image extensions
    supported_extensions = ('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')

    # Loop through the directory with tqdm progress bar
    for filename in tqdm(os.listdir(directory), desc="Processing Images", unit="image"):
        if filename.lower().endswith(supported_extensions):
            total_images += 1  # Count total images
            file_path = os.path.join(directory, filename)

            # Open and read the image using EasyOCR
            try:
                # Try using the Simplified Chinese reader first
                results = reader_simplified.readtext(file_path, detail=1)

                if not results:  # If no text found, try Traditional Chinese reader
                    results = reader_traditional.readtext(file_path, detail=1)

                if results:
                    images_with_text += 1  # Text was found
                    total_extracted_elements += len(results)  # Count extracted elements
                    extracted_text = ' '.join([res[1] for res in results])

                    # Calculate average confidence
                    confidences = [res[2] for res in results]
                    avg_confidence = sum(confidences) / len(confidences)

                    # Count low confidence text elements
                    low_confidence_elements = [c for c in confidences if c < 0.5]
                    low_confidence_count += len(low_confidence_elements)

                    # Determine verification status based on average confidence
                    verification_status = "Low Confidence" if avg_confidence < 0.6 else "Verified"

                    # Spell Checking for English Text
                    misspelled_words = []
                    detected_language = 'en'  # Default language since we are focusing on English
                    if detected_language == 'en':
                        words = preprocess_words(extracted_text)
                        misspelled = spell.unknown(words)
                        if misspelled:
                            misspelled_words = list(misspelled)
                            spell_errors_count += len(misspelled)
                            verification_status += f" | {len(misspelled)} Spell Errors"

                else:
                    images_without_text += 1  # No text found
                    extracted_text = "No text found"
                    avg_confidence = 0
                    verification_status = "No Text"
                    misspelled_words = []

                # Append to data
                data.append([file_path, filename, extracted_text, avg_confidence, verification_status, ', '.join(misspelled_words)])

            except Exception as e:
                failed_extractions += 1  # OCR failed
                print(f"Error processing {filename}: {e}")
                data.append([file_path, filename, "Error in OCR", 0, "Failed", ""])

    return data, total_images, images_with_text, images_without_text, failed_extractions, low_confidence_count, spell_errors_count, total_extracted_elements
# Function to save images and text to Excel
def save_to_excel(directory, data, excel_filename):
    # Create a workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "OCR Results"

    # Set column widths
    ws.column_dimensions['A'].width = 30  # Adjust the width of the column for images
    ws.column_dimensions['B'].width = 20  # Image name column width
    ws.column_dimensions['C'].width = 50  # Extracted text column width
    ws.column_dimensions['D'].width = 15  # Average Confidence column
    ws.column_dimensions['E'].width = 30  # Verification Status column
    ws.column_dimensions['F'].width = 50  # Misspelled Words column width

    # Write headers
    ws.append(["Image", "Image Name", "Extracted Text", "Avg Confidence", "Verification Status", "Misspelled Words"])

    for i, (image_path, image_name, extracted_text, avg_confidence, verification_status, misspelled_words) in enumerate(data, start=2):
        # Insert image into the Excel sheet
        try:
            img = ExcelImage(image_path)

            # Optionally, adjust image dimensions (optional)
            img.width = 200  # Adjust the width of the image
            img.height = 150  # Adjust the height of the image
            ws.add_image(img, f'A{i}')  # Add image in column A
        except Exception as e:
            print(f"Error adding image {image_name} to Excel: {e}")

        # Write the image name, extracted text, average confidence, verification status, and misspelled words into the sheet
        ws[f'B{i}'] = image_name
        ws[f'C{i}'] = extracted_text
        ws[f'D{i}'] = round(avg_confidence, 2)  # Round to two decimal places
        ws[f'E{i}'] = verification_status
        ws[f'F{i}'] = misspelled_words

        # Set row height to match the image height
        row_height = 150 * 0.75  # Approximate conversion from pixels to Excel row height
        ws.row_dimensions[i].height = row_height

    # Save the Excel file in the images directory
    excel_path = os.path.join(directory, excel_filename)
    wb.save(excel_path)
    print(f"Data has been successfully exported to {excel_path}")

# Function to generate stats report
def generate_stats_report(directory, total_images, images_with_text, images_without_text, failed_extractions, time_taken, low_confidence_count, spell_errors_count, total_extracted_elements):
    report_file_path = os.path.join(directory, "stats_report.txt")

    with open(report_file_path, "w") as report_file:
        report_file.write("EasyOCR Stats Report\n")
        report_file.write("====================\n\n")
        report_file.write(f"Time Taken: {time_taken:.2f} seconds\n")
        report_file.write(f"Total Images: {total_images}\n")
        report_file.write(f"Images with Text Extracted: {images_with_text} ({images_with_text / total_images * 100:.2f}%)\n")
        report_file.write(f"Images without Text: {images_without_text} ({images_without_text / total_images * 100:.2f}%)\n")
        report_file.write(f"Failed Extractions: {failed_extractions} ({failed_extractions / total_images * 100:.2f}%)\n")
        report_file.write(f"Low Confidence Text Elements: {low_confidence_count} ({low_confidence_count} / {total_extracted_elements} = {low_confidence_count / total_extracted_elements * 100:.2f}%)\n")
        report_file.write(f"Spell Errors (English): {spell_errors_count} ({spell_errors_count} / {total_extracted_elements} = {spell_errors_count / total_extracted_elements * 100:.2f}%)\n")

    print(f"Stats report saved to {report_file_path}")

# Main function to handle user input and export to Excel
def main():
    # Start the timer
    start_time = time.time()

    # Ask the user for the directory of images
    directory = input("Enter the directory of your images: ").strip()

    if not os.path.isdir(directory):
        print("Invalid directory. Please try again.")
        return

    # Ask the user for the excel file name
    excel_filename = input("Enter the name of your Excel file (without extension): ").strip()
    if not excel_filename:
        excel_filename = "OCR_Results.xlsx"
    else:
        excel_filename += ".xlsx"

    # Extract text from images
    data, total_images, images_with_text, images_without_text, failed_extractions, low_confidence_count, spell_errors_count, total_extracted_elements = extract_text_from_images(directory)

    # Save data to Excel
    save_to_excel(directory, data, excel_filename)

    # End the timer
    end_time = time.time()
    time_taken = end_time - start_time

    # Generate stats report
    generate_stats_report(directory, total_images, images_with_text, images_without_text, failed_extractions, time_taken, low_confidence_count, spell_errors_count, total_extracted_elements)

if __name__ == "__main__":
    main()