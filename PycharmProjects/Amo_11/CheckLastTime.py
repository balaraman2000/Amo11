import csv
import os
import GlobalLogs
import shutil
import sys
from datetime import datetime

import pywintypes




class CheckLastTime:
    def __init__(self, first_file, second_file, target_folder):




        self.first_file = first_file
        self.second_file = second_file
        self.target_folder = target_folder
        self.compare_last_modified_times(first_file, second_file, target_folder)

    def compare_last_modified_times(self, first_file, second_file, target_folder):
        try:
            first_file_modification_time = os.path.getmtime(first_file)
            second_file_modification_time = os.path.getmtime(second_file)
            if first_file_modification_time < second_file_modification_time:
                self.move_to_folder(second_file, target_folder)
            elif first_file_modification_time > second_file_modification_time:
                self.move_to_folder(first_file, target_folder)
            else:
                self.read_last_cell_data(first_file, second_file, target_folder)

        except FileNotFoundError:
            print("\n")
            print("Error:FileNotFoundError Or File path wrong.")
            print("The Error occurred due to any of the following reasons :")
            print("1. Required file not exists[dbspacechk.csv] ,please place the required file")
            print("2. Placed [dbspacechk.csv] file is not current week's file, please place the current week's file")
            print("3. please keep correct file path ")
            GlobalLogs.logging.error(f"Source file not found: {str(FileNotFoundError)} ")
            GlobalLogs.logging.error("The Error occurred due to any of the following reasons :")
            GlobalLogs.logging.error("1. Required file not exists[dbspacechk.csv] ,please place the required file")
            GlobalLogs.logging.error("2. Placed [dbspacechk.csv] file is not current week's file, please place the current week's file")

            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.error(f"Source file not found :{str(PermissionError)}.")
            sys.exit()
        except pywintypes.com_error as e:
            print(f"Error:Sheet not found")
            GlobalLogs.logging.error(f"Sheet not found")
            sys.exit()
        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()

    def move_to_folder(self, file, target_folder):
        try:
            shutil.copy2(file, target_folder)

            GlobalLogs.logging.info(f"File '{file}' moved to '{target_folder}' successfully.")
        except FileNotFoundError:
            print("Error:FileNotFoundError Or File path wrong.")
            print("The Error occurred due to any of the following reasons :")
            print("1. Required file not exists[dbspacechk.csv] ,please place the required file")
            print("2. Placed [dbspacechk.csv] file is not current week's file, please place the current week's file")
            print("3. please keep correct file path ")
            GlobalLogs.logging.error(f"Source file not found: {str(FileNotFoundError)} ")
            GlobalLogs.logging.error("The Error occurred due to any of the following reasons :")
            GlobalLogs.logging.error("1. Required file not exists[dbspacechk.csv] ,please place the required file")
            GlobalLogs.logging.error(
                "2. Placed [dbspacechk.csv] file is not current week's file, please place the current week's file")
            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.error(f"Source file not found :{str(PermissionError)}.")
            sys.exit()
        except pywintypes.com_error as e:
            print(f"Error:Sheet not found")
            GlobalLogs.logging.error(f"Sheet not found")
            sys.exit()
        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()


    def read_last_cell_data(self, first_file, second_file, target_folder):
        try:
            lastRowFirstFileData = self.read_first_column(first_file)
            lastRowSecondFileData = self.read_first_column(second_file)
            date_time1 = datetime.strptime(lastRowFirstFileData, "%Y/%m/%d %H:%M:%S")
            date_time2 = datetime.strptime(lastRowSecondFileData, "%Y/%m/%d %H:%M:%S")
            if date_time1 > date_time2:
                self.move_to_folder(first_file, target_folder)
            elif date_time1 < date_time2:
                self.move_to_folder(second_file, target_folder)
            else:
                self.move_to_folder(first_file, target_folder)

        except FileNotFoundError:
            print("Error:FileNotFoundError Or File path wrong.")
            print("The Error occurred due to any of the following reasons :")
            print("1. Required file not exists[dbspacechk.csv] ,please place the required file")
            print("2. Placed [dbspacechk.csv] file is not current week's file, please place the current week's file")
            print("3. please keep correct file path ")
            GlobalLogs.logging.error(f"Source file not found: {str(FileNotFoundError)} ")
            GlobalLogs.logging.error("The Error occurred due to any of the following reasons :")
            GlobalLogs.logging.error("1. Required file not exists[dbspacechk.csv] ,please place the required file")
            GlobalLogs.logging.error(
                "2. Placed [dbspacechk.csv] file is not current week's file, please place the current week's file")
            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.error(f"Source file not found: {str(PermissionError)} ")
            sys.exit()
        except pywintypes.com_error as e:
            print(f"Error:Sheet not found")
            GlobalLogs.logging.error(f"Sheet not found")
            sys.exit()
        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()

    def read_first_column(self, second_file):
        try:
            with open(second_file, newline='') as file:
                reader = csv.reader(file)
                last_data = None
                for row in reader:

                    if row:
                        last_data = row[0]
                return last_data

        except FileNotFoundError:
            print("Error:FileNotFoundError .")
            GlobalLogs.logging.error(f"Source file not found: {str(FileNotFoundError)} ")
            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.error(f"Source file not found :{str(PermissionError)}.")
            sys.exit()
        except pywintypes.com_error as e:
            print(f"Error:Sheet not found")
            GlobalLogs.logging.error(f"Sheet not found")
            sys.exit()
        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()