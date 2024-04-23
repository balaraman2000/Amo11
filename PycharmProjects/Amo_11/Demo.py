import os
import sys
from datetime import datetime
import GlobalLogs
from ChartSheetUpdate import ChartSheetUpdate
import pywintypes

class CharSheet:
    def __init__(self, directory):
        try:
            directory1 = os.path.join(directory, "DB容量貼り付けマクロ.xlsm")
            directory2 = os.path.join(directory, "~$DB容量貼り付けマクロ.xlsm")
            demo = os.path.join(directory, "DB容量使用状況_")

            self.directory = directory

            latest_mod_time = 0
            latest_file_path = ""

            files = os.listdir(self.directory)

            excel_files = [file for file in files if file.endswith(('.xlsx', '.xlsm'))]

            for excel_file in excel_files:
                file_path = os.path.join(self.directory, excel_file)

                mod_time = os.path.getmtime(file_path)

                if mod_time > latest_mod_time and directory1 != file_path and directory2 != file_path:
                    latest_mod_time = mod_time
                    latest_file_path = file_path

            latest_mod_time_formatted = datetime.fromtimestamp(latest_mod_time).strftime('%Y-%m-%d %H:%M:%S')

            # print(f" {latest_file_path}")
            GlobalLogs.logging.info(
                f"Latest modified file: {latest_file_path} - Last Modified: {latest_mod_time_formatted}")

            if latest_file_path.startswith(demo):


                object = ChartSheetUpdate(latest_file_path)

            else:
                print("Error:No Excel files found in the directory.")
                GlobalLogs.logging.error(("No Excel files found in the directory."))
                sys.exit()
        except FileNotFoundError:
            print("Error:Source file not found.")
            GlobalLogs.logging.exception(f"Source file not found: {str(FileNotFoundError)} ")
            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.exception(f"Source file not found: {str(PermissionError)} ")
            sys.exit()

        except pywintypes.com_error as e:
            print(f"Error:Sheet not found:{e}")
            GlobalLogs.logging.error(f"DB容量 Sheet not found")
            sys.exit()

        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()
