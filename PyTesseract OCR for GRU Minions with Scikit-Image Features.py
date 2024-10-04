#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import numpy as np
import cv2
import pytesseract
from PIL import Image
from skimage import exposure
from skimage import filters, img_as_ubyte
from skimage.restoration import denoise_bilateral
import pandas as pd


# In[2]:


pytesseract.pytesseract.tesseract_cmd = r'/opt/homebrew/bin/tesseract'


# In[3]:


def preprocess_image(img):
    """
    Preprocesses an image for improved OCR results.

    Args:
        img (PIL.Image): The image to be preprocessed.

    Returns:
        PIL.Image: The preprocessed image.
    """
    # Convert to grayscale
    img = img.convert('L')

    # Apply adaptive thresholding for better contrast
    thresh = cv2.adaptiveThreshold(np.array(img), 255, 
                                   cv2.ADAPTIVE_THRESH_MEAN_C, 
                                   cv2.THRESH_BINARY, 11, 2)

    return Image.fromarray(thresh)


# In[4]:


def create_scikit_image_version(img):
    """
    Applies Scikit-Image filters for image sharpening and manipulation.

    Args:
        img (PIL.Image): The original image.

    Returns:
        PIL.Image: The manipulated image.
    """
    # Convert to grayscale (if necessary)
    img = img.convert('L')
    
    # Convert to NumPy array
    img_array = np.array(img)

    # Apply bilateral denoising (preserves edges while reducing noise)
    denoised_img = denoise_bilateral(img_array, sigma_color=0.05, sigma_spatial=15)

    # Normalize denoised image to range [0, 1]
    denoised_img = exposure.rescale_intensity(denoised_img, out_range=(0, 1))

    # Apply edge enhancement filter
    edges = filters.sobel(denoised_img)

    # Combine the denoised image and edges
    enhanced_img = denoised_img + edges

    # Normalize the combined image to range [0, 255] and convert to 8-bit unsigned integer
    final_img = img_as_ubyte(exposure.rescale_intensity(enhanced_img, out_range=(0, 1)))
    
    return Image.fromarray(final_img)

def process_images(image_directory, excel_output_path):
    # Normalize path (expands ~ to the full user path)
    image_directory = os.path.expanduser(image_directory)

    image_files = [f for f in os.listdir(image_directory) 
                   if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]

    if not image_files:
        print("No images found in the directory.")
        return

    # List to store results for Excel export
    results = []

    for image_file in image_files:
        image_path = os.path.join(image_directory, image_file)
        try:
            img = Image.open(image_path)

            # Extract text from original image
            original_img = preprocess_image(img.copy())
            original_text = pytesseract.image_to_string(original_img)

            # Create manipulated image
            manipulated_img = create_scikit_image_version(img.copy())

            # Extract text from manipulated image
            manipulated_text = pytesseract.image_to_string(manipulated_img)

            # Append the results to the list (image file, original text, manipulated text)
            results.append({
                'Image File': image_file,
                'Original Text': original_text,
                'Manipulated Text': manipulated_text
            })

            # Print results (optional)
            print(f"Text extracted from {image_file} (Original):\n{original_text}\n")
            print(f"Text extracted from {image_file} (Scikit-Image):\n{manipulated_text}\n{'-'*40}")
        
        except Exception as e:
            print(f"Failed to process {image_file}: {e}")

    # Create a DataFrame from the results and export to Excel
    df = pd.DataFrame(results)
    df.to_excel(excel_output_path, index=False)

    print(f"Results successfully exported to {excel_output_path}")

# Set the path using expanduser
image_directory = '~/Downloads/FilesImage'
excel_output_path = '~/Downloads/extracted_text_results.xlsx'  # Path to save Excel file

process_images(image_directory, excel_output_path)

