# Army Individual Medical Readiness (IMR) Extractor

## Overview

The **Army-IMR-Extractor** automates the extraction of specific fields from one or multiple Individual Medical Readiness (IMR) PDFs into a neatly formatted Excel spreadsheet. This tool simplifies Army administrative tasks by reducing manual data entry and organizing medical readiness data for easy tracking.

**Note:** This script is designed to operate in **Windows 11 environments**.

## Prerequisites

1. **Install Visual Studio Build Tools**:  
   Install the Build Tools for Visual Studio to use the `fitz` library (part of the PyMuPDF package) for extracting data from PDF files.

   - Download from [Visual Studio Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/).
   - During installation, select the **Desktop development with C++** workload.
   - After installation, restart your computer.

2. **Install Required Python Libraries**:  
   Ensure the following Python packages are installed:
   ```bash
   pip install -r requirments.txt

## Usage

Follow these steps to use the Army-IMR-Extractor:

1. **Prepare IMR Files**:  
   Collect all the IMR PDF files you wish to extract data from and place them in the same directory as the `imr.py` script.

2. **Run the Script**:  
   Execute the script using Python and follow the on-screen prompts to specify the PDF files and output location for the Excel file.
   ```bash
   python ./imr.py

3. **Follow the Prompts**:  
   The script will prompt you to:
   - Enter the paths to the IMR PDFs.
   - Specify the output folder where the Excel file will be saved.

4. **View the Extracted Data**:  
   Once the extraction is complete, an Excel file with the extracted IMR data will be saved in your specified folder. The file will be named with the current date, e.g., `IMR_PDFs_YYYYMMDD.xlsx`.
