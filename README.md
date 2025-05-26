# Excel Input Driver Analyzer

This Python script analyzes Excel files to identify input drivers and their relationships within formulas. It's particularly useful for financial models where you need to track dependencies between cells and identify key input variables.

## Features

- Analyzes Excel files to find input drivers (cells that are referenced in formulas)
- Identifies specific cell references (e.g., CP235, CP243) and their usage in formulas
- Processes multiple sheets within an Excel file
- Generates detailed output including:
  - List of all input drivers
  - Formulas containing specific cell references
  - Sheet-wise analysis of input drivers
- Outputs results to CSV files for easy analysis

## Requirements

- Python 3.x
- Required packages (see requirements.txt):
  - openpyxl
  - pandas
  - numpy

## Installation

1. Clone this repository
2. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

Run the script with the following command:
```bash
python find_input_drivers.py <path_to_excel_file> <output_csv_file>
```

Example:
```bash
python find_input_drivers.py "path/to/your/model.xlsx" "input_drivers_output.csv"
```

## Output

The script generates a CSV file containing:
- Cell references
- Associated formulas
- Sheet names
- Column information
- Additional metadata about the input drivers

## Notes

- The script is designed to work with financial models and Excel files containing formulas
- It can handle multiple sheets and complex formula relationships
- Debug output is available for detailed analysis of the process

## License

MIT License 
