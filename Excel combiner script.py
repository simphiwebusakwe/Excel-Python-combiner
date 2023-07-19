import pandas as pd
import os
from datetime import datetime

def combine_excel_files():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = script_dir  # Set the input folder to the script's directory

    all_data = pd.DataFrame()

    for filename in os.listdir(input_folder):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            if "output" in filename.lower():
                continue  # Skip files with "output" in the filename
            file_path = os.path.join(input_folder, filename)
            df = pd.read_excel(file_path)
            all_data = pd.concat([all_data, df], ignore_index=True)

    # Generate output filename with date and time
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"output_{now}.xlsx"

    output_file = os.path.join(script_dir, output_filename)

    all_data.to_excel(output_file, index=False)
    print("Combined Excel file saved as", output_file)

# Example usage
combine_excel_files()
