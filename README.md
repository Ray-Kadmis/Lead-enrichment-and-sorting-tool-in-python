# Excel/CSV Sorter Tool

A Python tool that processes Excel and CSV files to sort and organize business data based on website domains and reviews.

## Features

- **Sorts empty website rows to the top** (by highest reviews)
- **Groups repeated business domains** together at the bottom
- **Handles both Excel (.xlsx) and CSV files**
- **Converts CSV to Excel output**
- **Process single or multiple files**
- **Combines multiple files into one** (optional)

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Interactive Mode (Easiest)
Simply run the script without arguments:
```bash
python excel_sorter.py
```

### Command Line Mode

#### Process a single file:
```bash
python excel_sorter.py data.xlsx
```

#### Process multiple files separately:
```bash
python excel_sorter.py file1.xlsx file2.csv file3.xlsx
```

#### Combine multiple files into one:
```bash
python excel_sorter.py --combine file1.xlsx file2.csv --output Combined_Cleaned.xlsx
```

## Required Columns

The tool looks for these columns (case-insensitive):
- **reviews** - Number of reviews
- **website** - Website URLs
- **rating** - Business ratings

## How It Works

1. **Empty Website Sorting**: Rows with empty website cells are moved to the top and sorted by highest reviews
2. **Domain Extraction**: Extracts domain names from URLs (e.g., "abc123" from "www.abc123.com")
3. **Duplicate Grouping**: Groups businesses with the same domain under "Repeated Businesses" section
4. **Output**: Saves as `filename_Cleaned.xlsx`

## Examples

### Input Data:
| Business | Website | Reviews | Rating |
|----------|---------|---------|--------|
| ABC Dental | www.abc123.com/dental | 50 | 4.5 |
| XYZ Store | | 100 | 4.0 |
| ABC Medical | https://abc123.com/medical | 30 | 4.2 |

### Output Data:
| Business | Website | Reviews | Rating |
|----------|---------|---------|--------|
| XYZ Store | | 100 | 4.0 |
| **Repeated Businesses** | | | |
| ABC Dental | www.abc123.com/dental | 50 | 4.5 |
| ABC Medical | https://abc123.com/medical | 30 | 4.2 |

## Error Handling

- Automatically detects column names (case-insensitive)
- Handles missing or invalid data gracefully
- Provides clear error messages for troubleshooting
