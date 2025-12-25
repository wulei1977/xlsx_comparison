# Excel File Comparison Tool

A tool for comparing differences between two Excel files, supporting both Web interface and command-line usage.

## Features

- Multi-column composite key matching
- Automatic detection of rows existing only in one file
- Cell-by-cell comparison for common rows
- Generate color-coded Excel files highlighting differences
- User-friendly Web interface
- Command-line support for batch processing
- Can be packaged as standalone executable

## Difference Highlighting

- 🟢 Green background: Rows existing only in the current file
- 🟡 Yellow background + Red font: Cells with differences from the other file
- Hover over cells to see comments showing corresponding values from the compared file

## Installation

```bash
pip install flask pandas openpyxl python-calamine
```

## Usage

### Web Interface

```bash
python app.py
```

After starting, visit `http://127.0.0.1:5000` and follow these steps:

1. Upload two Excel files
2. Select the Worksheet to compare
3. Select key columns for row matching (supports multiple columns)
4. Click "Start Compare"
5. View results and download comparison report or annotated files

### Command Line

```bash
python compare_xlsx.py file1.xlsx file2.xlsx --keys column1 column2 --sheet1 Sheet1 --sheet2 Sheet1
```

Parameters:
- `file1`, `file2`: Two Excel files to compare
- `--keys`: Column names for row matching (required, multiple allowed)
- `--sheet1`: Worksheet name in file1 (default: Sheet1)
- `--sheet2`: Worksheet name in file2 (default: Sheet1)
- `--output`: Output log file path (optional)

## Build Executable

```bash
pip install pyinstaller
pyinstaller build.spec
```

The executable will be in the `dist` directory.

## URL Prefix Configuration

Supports URL prefix configuration via argument or environment variable for reverse proxy deployment:

```bash
python app.py --prefix excel-compare
# or
set SCRIPT_NAME=excel-compare
python app.py
```

## Project Structure

```
├── app.py              # Web service main program
├── compare_xlsx.py     # Command-line comparison tool
├── templates/
│   └── index.html      # Web interface template
├── uploads/            # Temporary upload directory
├── build.spec          # PyInstaller build configuration
└── README.md
```

## License

MIT
