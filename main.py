# main.py
import click
from find_input_drivers import find_input_drivers


@click.command(
    help="""
Analyze an Excel model to identify input drivers and export results to a CSV file.

This tool scans the given Excel file for formulas and determines which cells
act as inputs to those formulas. It then attempts to label each input with its
corresponding row and column context (e.g., headers or descriptors nearby).

The output CSV includes:
- The sheet name
- Row and column labels for each input
- The cell reference and its content
- References to nearby cells (to the right)
- Useful metadata for review or auditing

Supports both `.xlsx` and legacy `.xls` formats.

Example:
  excel-input-driver-analyze model.xlsx input_drivers.csv
"""
)
@click.argument(
    "input_excel", type=click.Path(exists=True, readable=True), metavar="INPUT_EXCEL"
)
@click.argument("output_csv", type=click.Path(writable=True), metavar="OUTPUT_CSV")
def cli(input_excel, output_csv):
    """Extract input drivers from INPUT_EXCEL and save to OUTPUT_CSV."""
    find_input_drivers(input_excel, output_csv)


if __name__ == "__main__":
    cli()
