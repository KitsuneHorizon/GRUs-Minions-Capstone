#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import numpy as np
import cv2
import pytesseract
from PIL import Image
import pandas as pd


# In[2]:


pytesseract.pytesseract.tesseract_cmd = r'/opt/homebrew/bin/tesseract'


# In[3]:


def preprocess_image(img):
  """
  Preprocesses an image for improved OCR results.

  Args:
      img (np.ndarray): The image to be preprocessed as a NumPy array.

  Returns:
      np.ndarray: The preprocessed image as a NumPy array.
  """

  # Convert to grayscale using NumPy
  gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

  # Apply adaptive thresholding with NumPy functions
  thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                                cv2.THRESH_BINARY, 11, 2)

  # Reduce noise with Gaussian filtering using NumPy
  kernel = np.ones((5, 5), np.float32) / 25
  blur = cv2.filter2D(thresh, -1, kernel)

  return blur


# In[7]:


def process_images(image_directory, output_excel_path):
    # Normalize path (expands ~ to the full user path)
    image_directory = os.path.expanduser(image_directory)

    image_files = [f for f in os.listdir(image_directory)
                   if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]

    if not image_files:
        print("No images found in the directory.")
        return
      
    data = []

    for image_file in image_files:
        image_path = os.path.join(image_directory, image_file)
        try:
            img = cv2.imread(image_path)
            img = preprocess_image(img)  
            pil_image = Image.fromarray(img)  
            text = pytesseract.image_to_string(pil_image)
            
            # Append the filename and extracted text to the list
            data.append({'Filename': image_file, 'Extracted Text': text})
            
            print(f"Text extracted from {image_file}:\n{text}\n{'-'*40}")
        except Exception as e:
            print(f"Failed to process {image_path}: {e}")

    df = pd.DataFrame(data)

    df.to_excel(output_excel_path, index=False)

# Set the path using expanduser
image_directory = '~/Downloads/FilesImage'
output_excel_path = '~/Downloads/extracted_texts.xlsx'

process_images(image_directory, output_excel_path)

