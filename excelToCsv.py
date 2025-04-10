import os
import sys
import pandas as pd

def convert_excel_to_csv(input_folder: str, output_folder: str) -> None:
    """
    Convert all Excel files in the input folder to CSV files in the output folder,
    creating one CSV per worksheet. The name of each CSV will be:
        [excel_filename_without_ext]-[worksheet_name].csv
    """

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Loop through all files in the input folder
    for file_name in os.listdir(input_folder):
        # Build full path
        full_file_path = os.path.join(input_folder, file_name)

        # Ignore directories, only process .xlsx or .xls files
        if not os.path.isfile(full_file_path):
            continue
        if not (file_name.lower().endswith(".xlsx") or file_name.lower().endswith(".xls")):
            continue

        # Extract base name without extension
        file_base_name, _ = os.path.splitext(file_name)

        # Read the Excel file using pandas
        try:
            excel_file = pd.ExcelFile(full_file_path)
        except Exception as e:
            print(f"Skipping '{file_name}' due to error: {e}")
            continue

        # Convert each sheet to CSV
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                output_csv_name = f"{file_base_name}-{sheet_name}.csv"
                output_csv_path = os.path.join(output_folder, output_csv_name)
                df.to_csv(output_csv_path, index=False)
                print(f"Created CSV: {output_csv_name}")
            except Exception as e:
                print(f"Error processing sheet '{sheet_name}' in '{file_name}': {e}")


def main():
    """
    Main entry point to parse arguments and trigger the conversion function.
    Usage: python excel_to_csv.py [input_folder] [output_folder]
    """
    if len(sys.argv) != 3:
        print("Usage: python excel_to_csv.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]

    convert_excel_to_csv(input_folder, output_folder)


if __name__ == "__main__":
    main()
