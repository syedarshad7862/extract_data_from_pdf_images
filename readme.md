# PDF to Excel and Image Extractor

This Python project extracts text from images in a PDF file and saves it to an Excel file. Additionally, it extracts images embedded in the PDF and saves them to a specified folder. It uses **Poppler** to convert PDF pages to images, **Tesseract** for Optical Character Recognition (OCR), and **PyMuPDF (fitz)** for image extraction.

## Requirements

Before running the script, ensure you have the following installed:

1. **Poppler** - used for PDF-to-image conversion
   - Download Poppler from [Poppler](https://poppler.freedesktop.org/) and install it on your machine.
   - Set the `poppler_path` variable in the script to your Poppler installation directory.

2. **Python Libraries**:
   - Install the required Python libraries using pip:
     ```bash
     pip install pytesseract pdf2image pandas pymupdf
     ```

## Features

1. **Extract Text and Save to Excel**:
   - Converts PDF pages to images using **pdf2image**.
   - Uses **Tesseract OCR** to extract text from images.
   - Processes and saves the extracted data into an Excel file.

2. **Extract Images from PDF**:
   - Extracts all images embedded in a PDF using **PyMuPDF (fitz)**.
   - Saves the extracted images to a specified folder.
   - Avoids re-extracting images if they already exist in the folder.

## Steps to Run

### 1. Extract Text and Save to Excel

Use the `pdf_to_excel` function to extract text from a PDF and save it as an Excel file.

```python
pdf_to_excel('Telengana_Engineering_Colleges.pdf', 'colleges_with_header.xlsx', 'C:\\Program Files (x86)\\poppler-24.08.0\\Library\\bin')
```
- `pdf_path`: Path to the PDF file.
- `output_file`: Path to save the output Excel file.
- `poppler_path`: Path to the Poppler installation folder.

### 2. Extract Images from PDF

Use the following code snippet to extract images from a PDF:

```python
import fitz  # PyMuPDF
import os

def extract_images_from_pdf(pdf_path, output_folder):
    """
    Extracts images from a PDF file and saves them to the specified folder.
    Skips extraction if the images are already in the folder.

    Args:
        pdf_path: Path to the PDF file.
        output_folder: Path to the folder where images will be saved.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    doc = fitz.open(pdf_path)

    for page_index in range(len(doc)):
        page = doc[page_index]
        image_list = page.get_images(full=True)

        for img_index, img in enumerate(image_list):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            image_filename = os.path.join(output_folder, f'image_{page_index}_{img_index}.png')

            if not os.path.exists(image_filename):
                pix.save(image_filename)
                print(f"Saved: {image_filename}")
            else:
                print(f"Skipped (already exists): {image_filename}")

# Example usage
pdf_path = r"Telengana_Engineering_Colleges.pdf"
output_folder = r"extracted_images"
extract_images_from_pdf(pdf_path, output_folder)
```
- `pdf_path`: Path to the PDF file.
- `output_folder`: Folder to save extracted images.

## Output

### Excel Output:
An Excel file is generated with the following columns:
- **Code**: College code.
- **College Name**: Name of the engineering college.

Example:
| Code  | College Name                                                        |
|-------|---------------------------------------------------------------------|
| MHVR  | MAHAVEER INSTITUTE OF SCI. AND TECHNOLOGY, VYASAPURI, BANDLAGUDA P  |
| 5678  | XYZ Institute of Technology                                         |

### Image Output:
Extracted images from the PDF are saved as PNG files in the specified folder.

## Troubleshooting

1. **Poppler Issues**:
   - Ensure Poppler is correctly installed and the path is set in the script.

2. **OCR Accuracy**:
   - Adjust Tesseract OCR settings or check image quality for better results.

3. **Image Extraction**:
   - Ensure the output folder has write permissions.

## License

This project is open source and available under the MIT License.
