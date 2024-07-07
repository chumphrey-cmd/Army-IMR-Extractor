import fitz
import pandas as pd
import re
import os
from datetime import date

def get_pdf_paths():
    """
    Prompts the user to enter multiple PDF file paths and stores them in a list.
    The user can type 'cancel' to stop the input process or 'complete' to finish and return the list of paths.
    """
    pdf_paths = []

    while True:
        user_input = input("Enter the path to an IMR PDF file (or 'cancel' or 'complete'): ")

        if user_input.lower() == 'cancel':
            return None  # User canceled
        elif user_input.lower() == 'complete':
            if pdf_paths:
                return pdf_paths  # User finished and provided paths
            else:
                print("No PDF paths provided.")  # User finished but didn't provide any paths
        else:
            user_input = user_input.strip('"')  # Remove quotes if present
            if os.path.isfile(user_input):
                pdf_paths.append(user_input)
            else:
                print("Invalid file path. Please enter a valid path.")


def get_output_folder_path():
    """
    Prompts the user to enter a valid output folder path and returns it.
    """
    while True:
        folder_path = input("Enter the path to the folder where you want to save the Excel file: ")
        folder_path = folder_path.strip('"')  # Remove quotes from path provided if present

        if os.path.isdir(folder_path):
            return folder_path
        else:
            print("Invalid folder path. Please enter a valid path.")

def extract_imr_data(pdf_paths, field_locations):
    """
    Extracts specific data fields from a list of IMR PDFs based on predefined field locations.

    Args:
        pdf_paths (list): A list of file paths to the IMR PDF files.
        field_locations (dict): A dictionary mapping field names (keys) to their rectangular locations 
                                (x0, y0, x1, y1) on the PDF page (values).

    Returns:
        list: A list of dictionaries, where each dictionary contains the extracted data for one IMR PDF.
    """

    # Initialize an empty list to store the extracted data from each PDF.
    all_data = []

    # Iterate over each PDF file path in the provided list
    for pdf_path in pdf_paths:  
        # Create an empty dictionary to store the extracted data for the current PDF.
        data = {}  

        # Use a try-except block to handle potential file not found errors.
        try:
            # Open the PDF file using PyMuPDF's `fitz.open` function.
            # The `with` statement ensures the file is closed properly after processing.
            with fitz.open(pdf_path) as doc:  
                # Get the first page of the PDF (index 0). This assumes that all relevant data is on this page.
                page = doc[0]  

                # Iterate through each field defined in the `field_locations` dictionary.
                for field_name, location in field_locations.items():  
                    # Convert the field location tuple into a PyMuPDF rectangle object.
                    rect = fitz.Rect(location)  

                    # Extract text from the specified rectangular region on the page.
                    # The `clip=rect` argument limits the extraction to the defined rectangle.
                    # The output `words` is a list of tuples, each representing a word and its properties.
                    words = page.get_text("words", clip=rect)  

                    # Join the extracted words into a single string (field value) by concatenating the text 
                    # part of each word tuple (index 4).
                    field_value = " ".join(word[4] for word in words)  

                    # Clean up the extracted text by removing any characters that are not alphanumeric, 
                    # spaces, or hyphens. Then, strip any leading or trailing whitespace.
                    field_value = re.sub(r'[^\w\s-]', '', field_value).strip()  

                    # Store the cleaned field value in the `data` dictionary with the field name as the key.
                    data[field_name] = field_value 
        
        # If the PDF file is not found, catch the `FileNotFoundError` exception.
        except FileNotFoundError:
            # Print an error message indicating that the file was not found.
            print(f"Error: Could not find the IMR PDF at '{pdf_path}'. Skipping...")

        # If no error occurred (i.e., the PDF was processed successfully), 
        # append the extracted data dictionary for this PDF to the `all_data` list.
        else:
            all_data.append(data)

    # After processing all PDFs, return the list of dictionaries containing the extracted data.
    return all_data


'''
Field locations dictionary
    32.0: The x0-coordinate of the left edge of the box.
    70.0: The y0-coordinate of the top edge of the box.
    250: The x1-coordinate of the right edge of the box
    72.0: The y1-coordinate of the bottom edge of the box.
'''
field_locations = {
    "Name": (32.0, 70.0, 250.0, 72.0), # x1 coordinate of 250 extends to include the entire field of the "Personnel" section
    "SSN": (22.0, 73.0, 250.0, 85.0), 
    "Rank": (22.0, 87.0, 250.0, 98.0),
    "DOB": (22.0, 100.0, 250.0, 112.0), 
    "Sex": (22.0, 113.0, 250.0, 125.0), 
    "UIC": (22.0, 126.0, 250.0, 138.0), 
    "Description": (22.0, 139.0, 250.0, 151.0), 
    "Compo": (22.0, 153.0, 250.0, 164.0), 
    "Arrival Date": (22.0, 166.0, 250.0, 178.0), 
    "Location": (22.0, 179.0, 250.0, 191.0), 
    "Command": (22.0, 192.0, 250.0, 204.0), 
    "Duty Title": (22.0, 205.0, 250.0, 217.0), 
    "Duty AOC": (22.0, 219.0, 250.0, 231.0), 
    "VA Disability Rating": (22.0, 232.0, 250.0, 244.0), 
    "VA Disability Rating Date": (22.0, 254.0, 250.0, 266.0),

    "PULHES Code": (308.0, 60.0, 550.0, 72.0),
    "PULHES Source": (308.0, 73.0, 550.0, 85.0),
    "PHA Date": (308.0, 87.0, 600.0, 98.0),
    "Current Physical Exam Date": (308.0, 100.0, 550.0, 112.0),
    "Physical Category": (308.0, 113.0, 550.0, 125.0),
    "Height": (308.0, 126.0, 550.0, 138.0),
    "Weight": (308.0, 139.0, 550.0, 151.0),
    "Flight Status": (308.0, 153.0, 550.0, 164.0),

    "Dental Class": (22.0, 303.0, 300.0, 315.0),
    "Panograph": (22.0, 316.0, 300.0, 328.0),
    "Last Dental Exam": (22.0, 329.0, 300.0, 341.0),

    "Blood Type": (308.0, 258.0, 600.0, 270.0),
    "HIV Test Date": (308.0, 272.0, 600.0, 284.0),
    "DNA": (308.0, 285.0, 600.0, 297.0),
    "Sickle Cell Screen": (308.0, 298.0, 600.0, 310.0),
    "Sickle Cell Screen Date": (308.0, 311.0, 600.0, 323.0),
    "Pregnant": (308.0, 324.0, 600.0, 336.0),
    "G6PD Date": (308.0, 338.0, 600.0, 350.0),
    "G6PD Status": (308.0, 351.0, 600.0, 363.0),

    "Vision Class": (22.0, 369.0, 300.0, 381.0),
    "Vision Screening Date": (22.0, 382.0, 300.0, 394.0),
    "Two Pair of Glasses": (22.0, 396.0, 300.0, 408.0),
    "Mask Inserts": (22.0, 409.0, 300.0, 421.0),
    "Mission Required Contact Lenses (MRCL)": (22.0, 422.0, 350.0, 434.0),
    "Military Combat Eye Protection": (22.0, 444.0, 300.0, 456.0),
    "Military Combat Eye Protection Inserts": (22.0, 458.0, 300.0, 480.0),
    "Last Prescription Date On File": (22.0, 480.0, 300.0, 492.0),

    "IMM Profile": (308.0, 390.0, 600.0, 402.0),
    "180 Day Meds": (308.0, 404.0, 600.0, 416.0),

    "Medication": (308.0, 457.0, 600.0, 469.0),
    "Medical Warning Tags": (22.0, 621.0, 600.0, 633.0),
    "Immunization Record": (308.0, 483.0, 600.0, 495.0),
    "Summary Sheet of Medical Problems": (308.0, 496.0, 600.0, 508.0),
    "Corrective Lens Prescription": (308.0, 509.0, 600.0, 521.0),

    "Hearing Class": (22.0, 520.0, 300.0, 532.0),
    "Hearing Readiness Status": (22.0, 533.0, 300.0, 545.0),
    "Audiogram Date": (22.0, 546.0, 300.0, 558.0),
    "Triple or Single Flange Earplugs Issued?": (22.0, 559.0, 300.0, 585.0),

    "Latest Date for Pre": (308.0, 549.0, 600.0, 561.0),
    "Latest Date for Post": (308.0, 562.0, 500.0, 574.0),
    "Latest Date for PDHRA": (308.0, 575.0, 500.0, 587.0),

    "Hearing Aid": (22.0, 608.0, 300.0, 620.0),
    "Allergy / Conditions": (22.0, 635.0, 300.0, 647.0),
    "Respiratory": (22.0, 674.0, 300.0, 686.0)
}

# Main Script Execution:

pdf_paths = get_pdf_paths()  # Get PDF paths from user
if not pdf_paths:  # Handle the case where the user cancels
    exit("Exiting script: No PDF paths provided.")

output_folder = get_output_folder_path()  # Get output folder path from user

# Extract data from all PDFs
all_data = extract_imr_data(pdf_paths, field_locations)  

if all_data:  # Check if any data was extracted
    df = pd.DataFrame(all_data)

    # Create output Excel file
    today = date.today()
    default_file_name = f"IMR_PDFs_{today.strftime('%Y%m%d')}.xlsx"
    output_file_path = os.path.join(output_folder, default_file_name)

    df.to_excel(output_file_path, engine='openpyxl', index=False)
    print(f"Excel file saved to: {output_file_path}")
else:
    print("No data was extracted from the provided PDFs.")  # Error message if no data