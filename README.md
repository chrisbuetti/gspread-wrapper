# gspread-pandas-wrapper

A Python wrapper module that extends the [`gspread`](https://github.com/burnash/gspread) library with convenience methods for seamless integration with pandas. This utility simplifies working with Google Sheets as tabular data structures, adding helpful abstraction, automatic error handling, and pandas DataFrame compatibility.

## Features

- Retry logic with smart exception handling (rate limits, 502s, service unavailability)
- Fetch Google Sheets as pandas DataFrames
- Replace worksheet content with a DataFrame
- Append rows to a worksheet
- Update arbitrary cell ranges
- Access worksheet objects by name or ID
- Convenience methods for clearing data, deleting rows, batch operations, and more

## Installation

Install dependencies if not already installed:

```bash
pip install git+https://github.com/chrisbuetti/gspread-wrapper.git

```

Also ensure you have a Google service account JSON key file and that your target spreadsheet is shared with the service account's email.

## Usage

### Basic Initialization

```python
from gspread_wrapper import GSPREAD  # replace with actual filename if different
g = GSPREAD("My Workbook Name")
```

### Convert a Sheet to DataFrame

```python
df = g.sheet_to_df("Sheet1")
```

### Replace Sheet with a DataFrame

```python
g.replace_worksheet_with_df("Sheet1", df)
```

### Append a Row

```python
g.append_row_to_sheet("Sheet1", ["val1", "val2", "val3"])
```

### Update a Range

```python
g.update_rows_by_range("Sheet1", "A2:C2", [["a", "b", "c"]])
```

### Delete Rows

```python
g.delete_rows("Sheet1", 5, 10)
```

### Get Sheet URL

```python
url = g.get_sheet_url("Sheet1")
```

## Helper Functions and Features

- `gspread_function(f)`: Wraps gspread calls with retry logic for 502s, rate limits, and transient errors.
- `_sheet_check`: Resolves worksheet input whether it's a name, ID, or already a worksheet object.
- `_number_to_column`: Converts column numbers to Excel-style letters.
- `get_worksheet_dict`: Returns a mapping of worksheet names to IDs.
- `batch_clear`, `clear_basic_filter`: Utility methods to manipulate sheet formatting or clear filters.

## Error Handling

The wrapper retries on the following:

- JSON decoding errors
- 502 Server Errors
- Rate limits (read/write quota exceeded)
- General service unavailability

Retries are attempted up to 9 times with a 30-second delay.

## Requirements

- `gspread`
- `pandas`
- `requests`
- Google Service Account with access to target spreadsheets

## License

MIT
