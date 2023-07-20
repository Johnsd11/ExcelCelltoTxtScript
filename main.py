import pandas as pd
import os


def create_text_files_from_excel_column(excel_file, column_name):
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Create a folder to store the text files (if it doesn't exist)
    output_folder = "text_files"
    os.makedirs(output_folder, exist_ok=True)

    # Iterate through the cells in the specified column
    for index, value in enumerate(df[column_name]):
        if pd.notna(value):  # Check if the cell is not empty
            # Create a new text file with the cell's data
            file_name = f"{output_folder}/text_{index+2}.txt"
            with open(file_name, "w") as f:
                f.write(str(value))


if __name__ == "__main__":
    # Replace with the actual path to your Excel file
    excel_file_path = "TestExcelScript.xlsx"
    # Replace with the actual column name containing the data
    column_to_extract = "test"

    create_text_files_from_excel_column(excel_file_path, column_to_extract)
