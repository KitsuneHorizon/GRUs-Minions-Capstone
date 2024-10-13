# GRUs-Minions-Capstone
This repository is dedicated to Team GRU's collaborative efforts in researching and developing solutions for the Fentanyl Epidemic OCR/NLP Project. Our focus is to use Optical Character Recognition (OCR) and Natural Language Processing (NLP) techniques to analyze data related to the fentanyl epidemic, with the goal of providing valuable insights and improving data extraction methods.

## Table of Contents
- [Description](#description)
- [Getting Started](#getting-started)
  - [Repository Directory](#repository-directory)
  - [Dependencies](#dependencies)
  - [Installing](#installing)
  - [Executing Program](#executing-program)
- [Authors](#authors)
- [Acknowledgements](#acknowledgements)
- [Future Work](#future-work)

## Description
Our goal is to build upon the work completed during the Spring 2024 phase of the project. Team GRU's Minions are tasked with developing solutions to automate the extraction of entities from images and the processing of unstructured text. This will reduce manual effort and minimize errors in handling complex data. Additionally, we aim to fully leverage the evidence and data sources from Spring 2024, which were not entirely utilized in the previous phase.

## Getting Started
### Repository Directory
* **GRU's-Minions-Capstone/**  
  Main project folder containing all files and directories for the capstone project.

  * **.idea/**  
    Directory containing configuration files for development environments (specific to IDEs like PyCharm).  
    _Note: This is specific to development environments and won't affect project execution._

  * **README.md**  
    This file provides an overview of the project, detailed setup instructions, and other relevant information.

  * **miniProject.py**  
    Python script responsible for processing images using OCR, extracting text, and saving the results to an Excel file.
  
  * **EasyOCR.py**  
    Script that utilizes the EasyOCR library to perform Optical Character Recognition (OCR) on images. It allows for text extraction with minimal setup.

  * **Pillow_preprocessing.py**  
    This script uses the Pillow library for image preprocessing, such as sharpening and enhancing images before performing OCR.

  * **PyTesseract OCR for GRU Minions with NumPy (OpenCV).py**  
    Python script combining Tesseract OCR and NumPy/OpenCV for advanced image preprocessing and text extraction.
  
 * **PyTesseract OCR for GRU Minions with Scikit-Image Features.py**  
    A script that leverages Tesseract OCR with Scikit-Image features to manipulate images and improve text extraction.

  * **pyTorch.py**  
    This script demonstrates the use of PyTorch for image manipulation, aiming to sharpen images and enhance the quality of text extraction via OCR.
  
  * **NLP_test.ipynb**  
    Jupyter notebook used to experiment with various NLP techniques such as tokenization, named entity recognition (NER), and more. This is an interactive environment for exploring text analysis workflows.

  * **NLP_Test.py**  
    Python script that automates the NLP processes from `NLP_test.ipynb` and is designed for larger datasets. It performs tasks such as tokenization, stopword removal, and NER for unstructured text data.

  * **Translate Team 1.py**  
    Script that handles the translation of text from Chinese to English using the `googletrans` package. It's part of the preprocessing step for multilingual text sources.
  
     _Additional Python scripts will be developed and added as the project progresses._

### Dependencies
Ensure that the following dependencies are installed before running the project:

1. **[Tesseract-OCR executable](https://tesseract-ocr.github.io/tessdoc/Installation.html)**  
   * Follow the installation instructions for your operating system. Make sure `tesseract` is added to your system path.
   * _**Minimum version:** 5.0.0 (to support advanced image recognition features)._

2. **Python Packages:**
   * **pillow**  
     _**Minimum version:** 8.0.0_
     * Install via pip:
     ```bash
     pip install pillow==8.0.0
     ```
   * **openpyxl**  
     _**Minimum version:** 3.0.0_
     * Install via pip:
     ```bash
     pip install openpyxl==3.0.0
     ```

3. **Libraries for NLP:**
   * **pandas**  
     Install via pip:
     ```bash
     pip install pandas
     ```
   
   * **nltk**  
     Install via pip:
     ```bash
     pip install nltk
     ```
     Additionally, ensure you download the necessary nltk resources:
     ```bash
     python -m nltk.downloader punkt stopwords
     ```

   * **collections.Counter**  
     This is part of Python's standard library, so no installation is needed.

   * **wordcloud**  
     Install via pip:
     ```bash
     pip install wordcloud
     ```

   * **matplotlib**  
     Install via pip:
     ```bash
     pip install matplotlib
     ```

   * **re**  
     This is part of Python's standard library, so no installation is needed.

   * **unicodedata**  
     This is part of Python's standard library, so no installation is needed.
   
   * **googletrans**  
     Install via pip:
     ```bash
     pip install googletrans==4.0.0-rc1
     ```

   To use `Translator` from `googletrans`:
   ```python
   from googletrans import Translator

### Installing
1. Clone the Repository
   * Using the terminal:
     * Navigate to your desired directory: 
     ```bash
     cd /path/to/your/directory
     ```
     * Clone the repository using HTTPS or SSH: 
     ```bash
     git clone <Insert HTTPS/SSH link here>
     ```
     
2. Using GitHub Desktop:
  * Open GitHub Desktop
  * Click **File > Clone Repository**
  * Paste the repository URL (HTTPS/SSH)
  * Choose the directory where you want to clone it. 


### Executing program
To run the miniProject.py script, follow these steps:
1. **Ensure Dependencies are Installed:**
   * Before running the script, confirm that the necessary dependencies are installed:
     * Tesseract-OCR
     * Pillow (Python Image Library)
     * Openpyxl
   * Refer to the **Dependencies** section above for installation instructions.
2. **Run the Script:**
   * Open your terminal or command prompt.
   * Navigate to the directory where the script is located:
   ```bash
   cd /path/to/your/script
   ```
   * Execute the script by running the following command:
   ```bash
   python NLP_Test.py
   ```
   ```bash
   python EasyOCR.py
   ```
   ```bash
   python KerasOCR.py
   ```
   ```bash
   python miniProject.py
   ```
   ```bash
   python Pillow_preprocessing.py
   ```
   ```bash
   python pyTorch.py
   ```
 
3. **User Input Prompts:**
   - The script will prompt you to provide the path to the image directory and the desired output file name.
   - Input validation is performed to ensure the specified image directory exists and the output file name contains valid characters (alphanumeric only).

4. **Expected Output:**
   * The program generates an Excel file containing:
     * Images processed from the specified directory
     * File names
     * Image dimensions
     * Formats
     * Extracted text
     * Status of Processing (e.g., "OK" for success, "Corrupted" for errors)

5. Error Handling:
   * The script includes error handling for various scenarios, such as:
     * Image files that cannot be opened or processed
     * Permission issues when accessing the directory
     * Existing output files, prompting for overwrite or new name
   * Users will receive feedback in the terminal for any issues encountered during execution. 

## Authors
* KitsuneHorizon (Shayra Amador)
* ryanjeakins (Ryan Eakins)
* bgriffin (Bethany Griffin)
* beckyHuck (Rebecca Huck)
* k-longe (Kristina Longe)

## Acknowledgements
* **Dr. Howard** – Thank you for your guidance and support throughout the project.
* **Project Sponsors** – We appreciate your valuable feedback and the resources provided, which helped shape the direction of this project.

## Future Work
The project is a work in progress, and the following features or improvements are planned in future iterations:

1. **Enhancing the NLP Pipeline for Text Analysis:**  
   The initial NLP pipeline has been implemented to process text and perform tokenization, stopword removal, and Named Entity Recognition (NER). Future work will focus on improving entity extraction accuracy and expanding the pipeline to include sentiment analysis and topic modeling.

2. **Integration with Existing Data Sources:**  
   We'll integrate data sources from the Spring 2024 phase to maximize insights and ensure consistency in data processing. Additionally, we will continue refining the data cleaning and processing techniques to improve overall data quality.