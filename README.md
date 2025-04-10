# Excel to CSV Converter

This project provides a Python script that converts Excel files (with any number of worksheets) to one or more CSV files. Each worksheet within an Excel file is exported as its own CSV file.

## Requirements

The script requires the following Python packages:
- `pandas`
- `openpyxl`

You can install them via either of the following:

```bash
pip install pandas openpyxl

pip install -r requirements.txt
```

## Usage

1. Place your Excel files (in `.xlsx` or `.xls` format) inside an **input folder**.
2. Create an output folder where you want your generated CSV files to be stored.
3. Run the script from the command line by specifying the input and output folder paths:

```bash
python excelToCsv.py <input_folder> <output_folder>
```

### Example

```bash
python excelToCsv.py ./excel_files ./csv_output
```

- This command reads all `.xlsx` or `.xls` files in the `./excel_files` folder.
- The script will convert every worksheet in each Excel file into a separate CSV file.
- CSV files will be named in the format:
  ```
  <original_excel_filename_without_extension>-<worksheet_name>.csv
  ```
- All CSV files will be placed in the `./csv_output` folder.

## Notes

- The script automatically skips files that do not end in `.xlsx` or `.xls`.
- If any error occurs while reading an Excel file or a worksheet, the script logs the error and continues processing the remaining files.
- If the output folder does not exist, the script creates it automatically.

Feel free to modify this script or its dependencies to suit your needs.