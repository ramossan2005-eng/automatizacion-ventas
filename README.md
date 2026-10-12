# Sales Report Automation (Python + Excel)

Turns messy monthly sales spreadsheets into a clean, formatted Excel report with charts — in one command.

Many small businesses receive sales data from different people, each with their own format. This script reads every Excel file in a folder, fixes the inconsistencies automatically, and produces a ready-to-share report.

## What it does

- **Reads** all `.xlsx` files in the `datos/` folder
- **Standardizes** column names that differ between files (e.g. `Cant.` → `cantidad`, `Valor` → `precio`)
- **Cleans** the data:
  - trims extra spaces and fixes inconsistent capitalization (`"  BOGOTÁ "` → `"Bogotá"`)
  - converts prices stored as text (`"$180.000"`) into numbers
  - removes duplicate rows and rows with missing key data
- **Summarizes** sales by month, product, city and salesperson
- **Exports** a formatted Excel report with native Excel charts and a sheet with the cleaned data

## Output

`reporte_ventas.xlsx` contains:

| Sheet | Content |
| --- | --- |
| Resumen | Total sales, units, average sale, top product and city, data-cleaning stats |
| Por mes | Sales per month + chart |
| Por producto | Sales per product + chart |
| Por ciudad | Sales per city + chart |
| Por vendedor | Sales per salesperson + chart |
| Datos limpios | The full cleaned dataset, with the source file of each row |

## How to run

```bash
pip3 install -r requirements.txt
python3 generar_datos.py     # creates sample messy data in datos/
python3 reporte_ventas.py    # generates reporte_ventas.xlsx
```

To use your own data, replace the files in `datos/` with your spreadsheets. New column-name variations can be added to the `NOMBRES_COLUMNAS` dictionary in `reporte_ventas.py`.

## Tech stack

Python · pandas · openpyxl

## Author

Santiago — Python developer specializing in automation and data processing.
