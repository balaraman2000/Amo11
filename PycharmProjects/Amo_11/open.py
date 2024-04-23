import sys
import pygetwindow as gw
import GlobalLogs

class OpenExcel:
    def __init__(self):
        file_extensions = ['.xlsx', '.csv', '.txt']
        open_files = self.find_open_files(file_extensions)
        if open_files:

            for file_path in open_files:
                print(f"Please close all the currently opened  files.")
                GlobalLogs.logging.error(f"Please close all the currently opened  files : {file_path}")
                sys.exit()
        else:
            pass

    def find_open_files(self, file_extensions):  # Added self as the first parameter
        open_files = []
        for window_title in gw.getAllTitles():
            for ext in file_extensions:
                if ext in window_title.lower():
                    open_files.append(window_title)
        return open_files

