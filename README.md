# SQLite to Excel Converter

A simple Python program that exports data from a SQLite database into an Excel workbook.

## Features

- Convert SQLite tables to Excel
- Export all tables to separate worksheets
- Preserve column names
- Easy to use

## Requirements

- Python 3.x
- pandas
- openpyxl

Install the required packages:

```bash
pip install pandas openpyxl
```

## Usage

1. Place your SQLite database in the project folder.
2. Update the database path in the script if necessary.
3. Run the program:

```bash
python app/main.py
```

The program will generate an Excel file containing the exported database tables.

## License

This project is licensed under the MIT License.
