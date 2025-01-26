import pytesseract
import pandas as pd
import os

def image_to_excel(image_dir, output_file):
    """
    Extracts text from an image and creates an Excel file.

    Args:
        image_path: Path to the image file.
        output_file: Path to the output Excel file.
    """

    # text = pytesseract.image_to_string(image_path)

    # # Clean and split the text
    # lines = text.strip().split('\n')

    # Create a list to store data
    data = []
    for filename in os.listdir(image_dir):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
            image_path = os.path.join(image_dir, filename) 
            try:
                text = pytesseract.image_to_string(image_path)
                # text = pytesseract.image_to_string(image_path, config='--psm 6')

                # Clean and split the text
                lines = text.strip().split('\n')   
                # line = lines.strip().split('\n')   
                table_data = [
                    row.strip()
                    for row in lines
                    if row.strip() and not row.strip().lower().startswith("list of all engineering colleges in telangana")
                ]
                four_letter_list, other_list = separate_four_letter_strings(table_data)
                # four_letter_list, other_list = separate_strings(table_data)
                num = len(other_list)
                slice_indices = create_slice_indices(num,2)
                college_name = process_list(other_list)
                for code, name in zip(four_letter_list,college_name):
                    print(code,name)
                    parts = code.split()
                    parts2 = name
                    if parts:
                        c = parts[0]
                        n = parts2[0:]
                        for d in n:
                            print(d)
                        data.append({"Code": c, "College Name": d})
                    
            except Exception as e:
                print(f"Error processing {filename}: {e}")
            
           
    # data.append({"Code": four_letter_list, "Institute Name": other_list})
    # Create a DataFrame and save to Excel
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)


def separate_four_letter_strings(table_data):
    four_letter_list = []
    other_list = []
    
    for item in table_data:
        if len(item) == 4:
            four_letter_list.append(item)
        else:
            other_list.append(item)
    
    return four_letter_list, other_list
def create_slice_indices(num_elements, step_size):
    """
    Creates a list of tuples, where each tuple represents the 
    start and end indices for slicing a list into sublists.

    Args:
        num_elements: The number of elements in the list.
        step_size: The size of each slice.

    Returns:
        A list of tuples, where each tuple contains the start and end indices.
    """
    slice_indices = []
    start_index = 0
    while start_index < num_elements:
        end_index = min(start_index + step_size, num_elements)
        slice_indices.append((start_index, end_index))
        start_index = end_index
    return slice_indices

def process_list(data):
    """
    Processes the input list and returns a new list.
    If an element ends with ',' or '-' or contains ',' or '-' anywhere, it combines it with the next element.
    Otherwise, it adds the element as is.

    Args:
        data (list): List of strings to process.

    Returns:
        list: Processed list of lists.
    """
    new_list = []
    i = 0
    while i < len(data):
        # Check if the element ends with ',' or '-' or contains ',' or '-'
        if ("," in data[i] or "-" in data[i]) and i + 1 < len(data):
            # Combine current and next element
            combined = data[i].strip() + " " + data[i + 1].strip()
            new_list.append([combined])
            i += 2  # Skip the next element since it's already added
        else:
            # Add the element as is
            new_list.append([data[i].strip()])
            i += 1
    return new_list
# Example usage
image_to_excel('extracted_images', 'colleges.xlsx')