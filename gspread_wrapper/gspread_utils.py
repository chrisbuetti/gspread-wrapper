from time import sleep
import gspread
import traceback
import pandas as pd
import requests

print('V1.0.0')

# Global gspread client initialized with service account
gc = gspread.service_account()

def gspread_function(f):
    """
    Executes a gspread API call with retry logic for transient errors.

    Retries up to 9 times on:
    - JSONDecodeError (request parsing issue)
    - 502 Server Errors
    - Quota exceeded rate limits
    - Service Unavailable issues

    Parameters:
        f (function): Zero-argument function wrapping a gspread operation

    Returns:
        Result of the wrapped function if successful
    """
    counter = 0
    counter_threshold = 9

    while counter < counter_threshold:
        try:
            return f()
        except requests.exceptions.JSONDecodeError:
            print("JSONDecodeError. Sleeping for 30 seconds.")
            sleep(30)
        except gspread.client.APIError as e:
            tb = traceback.format_exc()
            if "502" in str(e) or "Server Error" in tb:
                print(f"502 Server Error. Retrying ({counter + 1}/{counter_threshold})...")
            elif "Quota exceeded" in tb or "Read requests per minute per user" in tb:
                print(f"Rate limit hit. Retrying ({counter + 1}/{counter_threshold})...")
            elif "the service is currently unavailable" in tb.lower():
                print(f"Service unavailable. Retrying ({counter + 1}/{counter_threshold})...")
            else:
                print(tb)
                raise
            sleep(30)
        counter += 1

    print("Final attempt after hitting retry threshold.")
    return f()


class GSPREAD:
    """
    Wrapper class around gspread to simplify common sheet operations
    using pandas DataFrames and provide automatic retrying.
    """

    def __init__(self, workbook_name):
        """
        Initialize the GSPREAD object and open the specified workbook.
        """
        self.gc = gspread.service_account()
        self.sh = gspread_function(lambda: self.open_workbook(workbook_name))

    def open_workbook(self, workbook_name):
        """Open the Google Sheets workbook by name."""
        return gspread_function(lambda: gc.open(workbook_name))

    def get_sheet_by_name(self, sheet_name):
        """
        Get worksheet by name (case-insensitive).
        Raises ValueError if name is not found.
        """
        worksheet_names = {k.lower(): v for k, v in self.get_worksheet_dict().items()}
        if sheet_name.lower() not in worksheet_names:
            raise ValueError(f"Worksheet name must be one of {list(worksheet_names.keys())}")
        return self.get_sheet_by_id(worksheet_names[sheet_name.lower()])

    def get_worksheet_dict(self):
        """
        Get a dictionary mapping sheet titles to their IDs.
        """
        worksheets = gspread_function(lambda: self.sh.worksheets())
        return {ws.title: ws.id for ws in worksheets}

    def get_sheet_by_id(self, sheet_id):
        """
        Get worksheet by internal gspread ID.
        """
        return gspread_function(lambda: self.sh.get_worksheet_by_id(sheet_id))

    def _number_to_column(self, n, start_at_0=False):
        """
        Convert a column index to its corresponding Excel-style letter.
        Example: 1 -> 'A', 27 -> 'AA'
        """
        if start_at_0:
            n += 1
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    
    def replace_worksheet_with_df(self, worksheet_name, df, extra_rows=0):
        """
        Replace the contents of a worksheet with a pandas DataFrame.

        Parameters:
            worksheet_name (str): Sheet name
            df (pd.DataFrame): Data to write
            extra_rows (int): Optional buffer rows to add
        """
        worksheet = self.get_sheet_by_name(worksheet_name)
        gspread_function(lambda: worksheet.clear_basic_filter())
        gspread_function(lambda: worksheet.freeze(rows=0, cols=0))
        
        # Prepare all rows: headers + data
        headers = [list(df.columns)]
        data_rows = gspread_function(lambda: df.values.tolist())
        all_rows = headers + data_rows
        
        # Clear and write everything at once
        gspread_function(lambda: worksheet.clear())
        col_letter = self._number_to_column(df.shape[1])
        range_label = f'A1:{col_letter}{len(all_rows)}'
        gspread_function(lambda: worksheet.update(range_label, all_rows))

        # Format the data rows (skip header row)
        if len(all_rows) > 1:
            gspread_function(lambda: worksheet.format(
                f"A2:{col_letter}{len(all_rows)}",
                {
                    "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "horizontalAlignment": "CENTER",
                    "textFormat": {
                        "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0},
                        "fontSize": 10,
                        "bold": False,
                    },
                },
            ))
        gspread_function(lambda: worksheet.freeze(rows=1))
        return worksheet


    def update_rows_by_range(self, sheet, cell_range, values):
        """
        Update a rectangular cell range with values.

        Parameters:
            sheet: sheet name, ID, or worksheet object
            cell_range (str): Range like 'A1:C3'
            values (List[List[Any]]): New data
        """
        worksheet = self._sheet_check(sheet)
        gspread_function(lambda: worksheet.update(cell_range, values))
        return worksheet

    def _sheet_check(self, sheet):
        """
        Normalize input (name, ID, or object) to a worksheet object.
        """
        if isinstance(sheet, gspread.worksheet.Worksheet):
            return sheet
        elif isinstance(sheet, str):
            return self.get_sheet_by_name(sheet)
        elif isinstance(sheet, int):
            return self.get_sheet_by_id(sheet)
        raise ValueError("Invalid sheet reference. Use name (str), ID (int), or Worksheet object.")

    def sheet_to_df(self, sheet):
        """
        Convert a worksheet to a pandas DataFrame using the first row as headers.
        """
        sheet = self._sheet_check(sheet)
        data = gspread_function(lambda: sheet.get_all_values())

        df = pd.DataFrame(data)
        df.columns = df.iloc[0]  # Set first row as header
        df = df[1:].reset_index(drop=True)  # Drop header and reset index
        return df

    def append_row_to_sheet(self, sheet, row):
        """
        Append a single row to the bottom of a worksheet.
        """
        sheet = self._sheet_check(sheet)
        gspread_function(lambda: sheet.append_row(row))
        return sheet

    def get_sheet_url(self, sheet):
        """
        Get the public URL of a worksheet.
        """
        sheet = self._sheet_check(sheet)
        return sheet.url

    def get_column_values(self, sheet, index):
        """
        Get all values from a specific column (1-based index).
        """
        sheet = self._sheet_check(sheet)
        return gspread_function(lambda: sheet.col_values(index))

    def batch_clear(self, sheet, range):
        """
        Clear a specified range in the worksheet.
        """
        sheet = self._sheet_check(sheet)
        return gspread_function(lambda: sheet.batch_clear([range]))

    def clear_basic_filter(self, sheet):
        """
        Remove basic filter from a worksheet (if applied).
        """
        sheet = self._sheet_check(sheet)
        return gspread_function(lambda: sheet.clear_basic_filter())

    def delete_rows(self, sheet, starting_index, ending_index):
        """
        Delete rows from a worksheet between the specified indexes.
        """
        sheet = self._sheet_check(sheet)
        return gspread_function(lambda: sheet.delete_rows(starting_index, ending_index))
