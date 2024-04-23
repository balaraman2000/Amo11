import csv
import shutil
import sys
from datetime import datetime

import pywintypes

import GlobalLogs


class CompareForuFile:
    def __init__(self, first_File, second_file, third_file, four_file, target_folder):
        self.first_File = first_File
        self.second_file = second_file
        self.third_file = third_file
        self.four_file = four_file
        self.target_folder = target_folder
        self.compare_last_modified_times(first_File, second_file, third_file, four_file, target_folder)

    def compare_last_modified_times(self, first_File, second_file, third_file, four_file, target_folder):
        try:
            last_row_first_file_data = self.read_first_column(first_File)
            last_row_second_file_data = self.read_first_column(second_file)
            last_row_third_file_data = self.read_first_column(third_file)
            last_row_four_file_data = self.read_first_column(four_file)

            date_time1 = datetime.strptime(last_row_first_file_data, "%Y/%m/%d %H:%M:%S")
            date_ime2 = datetime.strptime(last_row_second_file_data, "%Y/%m/%d %H:%M:%S")
            date_time3 = datetime.strptime(last_row_third_file_data, "%Y/%m/%d %H:%M:%S")
            date_time4 = datetime.strptime(last_row_four_file_data, "%Y/%m/%d %H:%M:%S")

            if date_time1 > date_ime2:
                self.check_to_move(date_time1, date_time3, date_time4)
            elif date_time1 < date_ime2:
                self.check_to_move(date_ime2, date_time3, date_time4)
            else:
                self.check_to_move(date_time1, date_time3, date_time4)

        except FileNotFoundError:
            print("Error:FileNotFoundError Or File path wrong.")
            print("The Error occurred due to any of the following reasons :")
            print("1. Required file not exists[dbspacechk.csv] ,please place the required file")
            print("2. Placed [dbspacechk.csv] file is not current week's file, please place the current week's file")
            GlobalLogs.logging.error(f"Source file not found: {str(FileNotFoundError)} ")
            GlobalLogs.logging.error("The Error occurred due to any of the following reasons :")
            GlobalLogs.logging.error("1. Required file not exists[dbspacechk.csv] ,please place the required file")
            GlobalLogs.logging.error(
                "2. Placed [dbspacechk.csv] file is not current week's file, please place the current week's file")
            GlobalLogs.logging.error("3. please keep correct file path ")

            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.error(f"Error:Source file not found .")
            sys.exit()
        except pywintypes.com_error as e:
            print(f"Error:Sheet not found")
            GlobalLogs.logging.error(f"Sheet not found")
            sys.exit()
        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()



    def check_to_move(self, result, date_time3, date_time4):
        try:
            if date_time3 > date_time4:
                if date_time3 == result:
                    self.move_to_folder(self.first_File)
                elif date_time3 > result:
                    self.move_to_folder(self.third_file)
                elif date_time3 < result:
                    self.move_to_folder(self.first_File)

            elif date_time3 < date_time4:
                if date_time4 == result:
                    self.move_to_folder(self.first_File)
                elif date_time4 > result:
                    self.move_to_folder(self.four_file)
                elif date_time4 < result:
                    self.move_to_folder(self.first_File)
            else:
                self.move_to_folder(self.first_File)
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
            GlobalLogs.logging.error("3. please keep correct file path ")

            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.error(f"Source file not found. ")
            sys.exit()
        except pywintypes.com_error as e:
            print(f"Error:Sheet not found.")
            GlobalLogs.logging.error(f"Sheet not found.")
            sys.exit()
        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()


    def move_to_folder(self, file):
        try:
            shutil.copy2(file, self.target_folder)

            GlobalLogs.logging.info(f"File '{file}' moved to '{self.target_folder}' successfully.")
        except FileNotFoundError:
            print("Error:Source file not found.")
            GlobalLogs.logging.error(f"Source file not found.")
            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.error(f"Source file not found.")
            sys.exit()
        except pywintypes.com_error as e:
            print(f"Error:Sheet not found")
            GlobalLogs.logging.error(f"Sheet not found")
            sys.exit()
        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()

    def read_first_column(self, first_file):
        try:
            with open(first_file, newline='') as file:
                reader = csv.reader(file)
                last_data = None
                for row in reader:

                    if row:
                        last_data = row[0]
                return last_data
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
            GlobalLogs.logging.error("3. please keep correct file path ")
            sys.exit()
        except PermissionError:
            print("Error:Permission denied.")
            GlobalLogs.logging.error(f"Source file not found. ")
            sys.exit()
        except pywintypes.com_error as e:
            print(f"Error:Sheet not found")
            GlobalLogs.logging.error(f"Sheet not found.")
            sys.exit()
        except Exception as e:
            print("An error occurred:", e)
            GlobalLogs.logging.error(f"An error occurred: {str(e)}")
            sys.exit()
