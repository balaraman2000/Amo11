import os
import sys
import win32com.client
import pywintypes
import GlobalLogs


class ClickMacro:
    def __init__(self, file_path, macro_name, graph_path):

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


        except FileNotFoundError:
            print("Error:Source file not found.")
            GlobalLogs.logging.exception(f"Source file not found: {str(FileNotFoundError)} ")
            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.exception(f"Source file not found: {str(PermissionError)} ")
            sys.exit()

        except pywintypes.com_error as e:
            print(f"Error:Sheet not found")
            GlobalLogs.logging.error(f"Error:Sheet not found")
            sys.exit()

        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred")
            sys.exit()
        finally:
            excel.Quit()
