import os
import win32com.client

import GlobalLogs


class ClickMacro:
    def __init__(self, file_path, macro_name):
        print("Executing the Macro...")
        GlobalLogs.logging.info("Executing the Macro...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        try:
            wb = excel.Workbooks.Open(os.path.abspath(file_path))
            excel.Run(macro_name)
            wb.Save()
            print("Macro Executed Succesfully")
            GlobalLogs.logging.info("Macro Executed Succesfully")

        except Exception as e:
            print(f"Error: {e}")
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
        finally:
            excel.Quit()


