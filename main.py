import pandas as pd
import os
import string

# 13 = id
# 35 = RPT_TEXT
# 47 = RPT_ID


def remove_non_ascii (a_str):
    ascii_chars = set(string.printable)

    return ''.join(
        filter(lambda x: x in ascii_chars, a_str)
    )


def create_text_files_from_excel_column(excel_file):
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Create a folder to store the text files (if it doesn't exist)
    output_folder = "text_files"
    os.makedirs(output_folder, exist_ok=True)

    # Iterate through the cells in the specified column
    for index, row in df.iterrows():
        Pid = str(row["id"])
        report_id = str(row["RPT_ID"])
        text = row["RPT_TEXT"]

        if pd.notna(text):  # Check if the cell is not empty
            # Create a new text file with the cell's data
            os.makedirs(output_folder+ "/patient_" + Pid, exist_ok=True)
            file_name = f"{output_folder}/patient_{Pid}/report_{report_id}.txt"
            cleaned_value = remove_non_ascii(text)
            with open(file_name, "w") as f:
                f.write(str(cleaned_value))


if __name__ == "__main__":
    # Replace with the actual path to your Excel file
    excel_file_path = "TestExcelScript.xlsx"

    create_text_files_from_excel_column(excel_file_path)
