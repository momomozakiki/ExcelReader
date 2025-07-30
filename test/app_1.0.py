from keyword import kwlist
from pathlib import Path
import pandas as pd
import json

# file_path = r"C:\Users\kbsim\Downloads\ST-2025-02-463_(THONG SIEK FOOD INDUSTRY PTE. LTD.) (1).xlsx"
# file_path = r"R:\Quotation\SIM\2025\ST-2025-03-002_SERVICE(FISCHER BELL PRIVATE LTD).xlsx"
file_path = r"C:\Users\ST-Service\Desktop\ST-2025-03-002_SERVICE(FISCHER BELL PRIVATE LTD).xlsx"
# keywords = 'Product : , Brand : , Model : , Capacity : , Pan Size : '
kw_list = ['Product :', 'Brand :', 'Model :', 'Capacity :', 'Pan Size :', 'Quotation No', 'SCALE-TECH (GLOBAL) PTE. LTD ']


def content_boundaries(content_name, start_keywords, end_keywords, file_path):
    """
    Finds the start and end row indices of a content block in an Excel file.

    Parameters:
        content_name (str): Logical name for the content block
        start_keywords (list): List of keywords that all must be present in a row to mark the start
        end_keywords (list): List of keywords that all must be present in a row to mark the end
        file_path (str): Path to the Excel file

    Returns:
        tuple: (content_name, start_row, end_row)

    Raises:
        ValueError: If any keyword is not found or invalid order
    """
    df = pd.read_excel(file_path, header=None)  # Read without assuming headers

    start_row = None
    end_row = None

    for idx, row in df.iterrows():
        # Join all non-empty cells in the row into a single string for easier search
        row_text = " ".join(str(cell).strip().lower() for cell in row if pd.notna(cell))

        # Check if all start_keywords are in this row
        if start_row is None and all(kw.lower() in row_text for kw in start_keywords):
            start_row = idx

        # Check if all end_keywords are in this row, and stop scanning after found
        if end_row is None and all(kw.lower() in row_text for kw in end_keywords):
            end_row = idx
            break  # No need to keep scanning

    # --- Post-check validations ---
    if start_row is None:
        raise ValueError(f"Start keywords {start_keywords} not found in the file.")

    if end_row is None:
        raise ValueError(f"End keywords {end_keywords} not found in the file.")

    if start_row > end_row:
        raise ValueError(
            f"Invalid block: Start row ({start_row}) comes after End row ({end_row})."
        )

    return content_name, start_row, end_row


def get_file_name(file_path):
    # file name with the file extension
    # file_name = Path(file_path).name
    # print(file_name)

    # file name without file extension
    file_stem = Path(file_path).stem
    # print(file_stem)

    return file_stem


import pandas as pd
from openpyxl.utils import column_index_from_string


def scan_excel_with_pandas(file_path, keywords, start_col=0, end_col=19, start_row=0, end_row=58):
    """
    Scans a specified range in an Excel file for cells containing any of the given keywords.

    Parameters:
        file_path (str): Path to the Excel file
        keywords (list): List of keywords to search for
        start_col (int or str): Starting column (e.g., 0 or 'A')
        end_col (int or str): Ending column (e.g., 19 or 'T')
        start_row (int): Starting row (0-based)
        end_row (int): Ending row (0-based)

    Returns:
        dict: Dictionary of matching cells in format {"A1": {"Product :": "Scale"}, ...}
    """

    # Convert column letters to numeric index if needed (e.g., 'A' -> 0)
    if isinstance(start_col, str):
        start_col = column_index_from_string(start_col) - 1  # 'A' -> 0

    if isinstance(end_col, str):
        end_col = column_index_from_string(end_col) - 1  # 'T' -> 19

    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path, header=None)

    result = {}

    # Loop through each cell in the defined range
    for row_idx, row in df.iterrows():
        if not (start_row <= row_idx < end_row):
            continue  # Skip out-of-range rows

        for col_idx, value in enumerate(row):
            if not (start_col <= col_idx <= end_col):
                continue  # Skip out-of-range columns

            cell_value = str(value).strip() if pd.notna(value) else ""
            if not cell_value:
                continue

            matched_keywords = [
                keyword for keyword in keywords
                if keyword.lower() in cell_value.lower()
            ]

            if matched_keywords:
                column_letter = chr(65 + col_idx)  # 0 -> A, 1 -> B, etc.
                cell_address = f"{column_letter}{row_idx + 1}"  # Excel-style row number

                result[cell_address] = {
                    keyword: cell_value for keyword in matched_keywords
                }

    return result


# --- Main Execution ---
if __name__ == "__main__":

    kw_list = ['Product :', 'Brand :', 'Model :', 'Capacity :', 'Pan Size :', 'Quotation No']
    # User inputs
    # file_path = input("Enter the path to your Excel file (.xlsx): ").strip().strip('"').strip("'")
    # keywords_input = input("Enter keywords separated by commas (e.g., Product, Brand, Model): ")
    # print(keywords_input)

    # keywords = [kw.strip() for kw in keywords_input.split(',')]
    # print(keywords)

    # Run scanner
    # output = scan_excel_with_pandas(file_path, keywords)
    output = scan_excel_with_pandas(file_path, kw_list)

    # Print JSON output
    print(json.dumps(output, indent=4))

    # Optional: Save to JSON file
    with open(get_file_name(file_path), "w") as f:
        json.dump(output, f, indent=4)

    print(f"✅ Result saved to {get_file_name(file_path)}")