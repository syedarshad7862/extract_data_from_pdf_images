
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
import os


# Specify the correct path to Poppler
poppler_path = r"C:\Program Files (x86)\poppler-24.08.0\Library\bin"  # Your Poppler installation path
def pdf_to_excel(pdf_path, output_file, poppler_path):
    try:
        images = convert_from_path(pdf_path, poppler_path=poppler_path)
        print(f"Successfully converted PDF to {len(images)} images.")
    except Exception as e:
        print(f"Error reading PDF file: {e}")
        return

    data = []

    for i, image in enumerate(images):
        try:
            text = pytesseract.image_to_string(image)
            lines = text.strip().split('\n')
            table_data = [
                row.strip()
                for row in lines
                if row.strip() and not row.strip().lower().startswith("list of all engineering colleges in telangana")
            ]
            four_letter_list, other_list = separate_four_letter_strings(table_data)
            college_name = process_list(other_list)

            for code, name in zip(four_letter_list, college_name):
                parts = code.split()
                if parts:
                    c = parts[0]
                    n = name[0:]
                    data.append({"Code": c, "College Name": n})

        except Exception as e:
            print(f"Error processing image {i}: {e}")

    # Create a DataFrame and add the custom header
    df = pd.DataFrame(data)
    header_row = {"Code": "", "College Name": "List of all engineering colleges in Telangana"}
    df = pd.concat([pd.DataFrame([header_row]), df], ignore_index=True)

    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"Data successfully saved to {output_file}")

def separate_four_letter_strings(table_data):
    four_letter_list = []
    other_list = []
    
    for item in table_data:
        if len(item) == 4:
            four_letter_list.append(item)
        else:
            other_list.append(item)
    
    return four_letter_list, other_list

def process_list(data):
    new_list = []
    i = 0
    while i < len(data):
        if ("," in data[i] or "-" in data[i]) and i + 1 < len(data):
            combined = data[i].strip() + " " + data[i + 1].strip()
            new_list.append([combined])
            i += 2
        else:
            new_list.append([data[i].strip()])
            i += 1
    return new_list

# Example usage
pdf_to_excel('Telengana_Engineering_Colleges.pdf', 'colleges_with_header.xlsx',poppler_path)

# from pdf2image import convert_from_path
# import pytesseract

# # Specify the correct path to Poppler
# poppler_path = r"C:\Program Files (x86)\poppler-24.08.0\Library\bin"  # Your Poppler installation path

# # Convert the PDF to images
# try:
#     images = convert_from_path('Telengana_Engineering_Colleges.pdf', poppler_path=poppler_path)
#     print(f"Converted {len(images)} pages to images.")
# except Exception as e:
#     print(f"Error reading PDF file: {e}")


