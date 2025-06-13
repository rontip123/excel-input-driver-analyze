# Excel Input Driver Analyzer

This Python tool analyzes Excel files to identify input drivers and their relationships within formulas. It is particularly useful for financial models where it's important to track dependencies between cells and identify key input variables.

## Features

- Analyzes Excel files to find input drivers (cells that are referenced in formulas)
- Attempts to label each input based on surrounding row and column headers
- Supports both `.xlsx` and `.xls` formats
- Processes multiple sheets within an Excel file
- Generates a detailed output including:
  - Sheet name
  - Row and column labels
  - Cell reference and content
  - Adjacent right-hand cell references and their labels
- Outputs results to a CSV file for easy analysis

## Requirements

- Python 3.11 or later
- `uv` (for environment and dependency management)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/your-org/excel-input-driver-analyze.git
   cd excel-input-driver-analyze
   ```

2. Initialize the project:
   ```bash
   uv init
   ```

3. Create a virtual environment:
   ```bash
   uv venv
   ```

4. Install the project in editable mode:
   ```bash
   uv pip install --editable .
   ```

5. Optionally freeze the current environment:
   ```bash
   uv pip freeze > requirements.txt
   ```

## Usage

Run the CLI tool using:

```bash
uv run excel-input-driver-analyze --input-excel <path_to_excel_file> --output-csv <output_csv_file>
```

### Example

```bash
uv run excel-input-driver-analyze --input-excel model.xlsx --output-csv input_drivers.csv
```

## Output

The script generates a CSV file containing:
- Sheet name
- Cell references and content
- Row and column labels associated with each input
- Cell references and labels for neighboring cells to the right

## Notes

- The script is designed to work with financial models and Excel files containing formulas
- It can handle multiple sheets and complex formula relationships
- Supports `.xlsx` files via `openpyxl` and `.xls` files via `xlrd`
- Debug and advanced options may be added in future releases

## License

MIT License
