# GRUs-Minions-Capstone
This repository is dedicated to Team GRU's collaborative efforts in researching and developing solutions for the Fentanyl Epidemic OCR/NLP Project. Our focus is to use Optical Character Recognition (OCR) and Natural Language Processing (NLP) techniques to analyze data related to the fentanyl epidemic, with the goal of providing valuable insights and improving data extraction methods.

## Table of Contents
- [Description](#description)
- [Getting Started](#getting-started)
- [Executing Program](#executing-program)
- [Authors](#authors)
- [Acknowledgements](#acknowledgements)
- [Future Work](#future-work)
- 
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
   python miniProject.py
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
* bgriffin (Bethany Griffin)
* beckyHuck (Rebecca Huck)
* k-longe (Kristina Longe)
* ryanjeakins (Ryan Eakins)

## Acknowledgements
* **Dr. Howard** – Thank you for your guidance and support throughout the project.
* **Project Sponsors** – We appreciate your valuable feedback and the resources provided, which helped shape the direction of this project.

## Future Work
The project is a work in progress, and the following features or improvements are planned in future iterations:

1. **NLP Pipeline for Text Analysis:**  
   We will integrate a Natural Language Processing (NLP) pipeline to analyze the extracted text and identify entities related to the fentanyl epidemic.

2. **Integration with Existing Data Sources:**  
   We'll integrate data sources from the Spring 2024 phase to maximize insights and ensure consistency in data processing.