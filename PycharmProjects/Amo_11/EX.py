import win32com.client
import pythoncom

def num2(self):
    try:
        pythoncom.CoInitialize()
        xl_app = win32com.client.Dispatch("Excel.Application")
        wb1 = xl_app.Workbooks.Open(self.file_path, ReadOnly=True)
        sheet = wb1.ActiveSheet
        last_data_column_index = None
        for col in range(1, sheet.Columns.Count + 1):
            if sheet.Cells(13, col).Value is not None:
                last_data_column_index = col
                break
    except Exception as e:
        print(e)
    finally:
        wb1.Close(False)
        xl_app.Quit()
        pythoncom.CoUninitialize()

    if last_data_column_index is not None:
        last_data_column_letter = chr(64 + last_data_column_index)
        self.las_num = f"{last_data_column_letter}13"
        print("Index of the last data cell in row 14:", self.las_num)
    else:
        print("No data found.")
