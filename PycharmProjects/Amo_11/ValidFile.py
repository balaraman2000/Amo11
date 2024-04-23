import os
import sys

import GlobalLogs


def Vaild_file(directory):
    macro = "DB容量貼り付けマクロ"
    last_week_file = "DB容量使用状況_"

    files = os.listdir(directory)
    excel_files = [file for file in files if file.endswith(('.xlsx', '.xlsm'))]
    tem1 = 0
    my_list = []
    my_list1 = []
    for excel_file in excel_files:
        # print(excel_file)
        if excel_file.startswith(macro):
            tem = excel_file

        elif excel_file.startswith(last_week_file):
            tem1 += 1

            value_to_add = excel_file
            my_list.append(value_to_add)
            if tem1 == 1:
                pass
            elif tem1 == 2:
                print("\n")
                print(f"The existing files in build path:{my_list}")

                print("The Error occurred due to any of the following reasons :")

                print("1. Required file not exists[DB容量使用状況_YYYYMMDD.xlsx] ,please place the required file")
                print(
                    "2. Placed [DB容量使用状況_YYYYMMDD.xlsx] file is not current week's file, please place the current week's file")
                print(
                    "3. Current week's output file already exists,please move the existing output file to bk & execute")
                print("4. please keep correct file path ")

                GlobalLogs.logging.error(f"The existing files in build path:{my_list}")
                GlobalLogs.logging.error("The Error occurred due to any of the following reasons :")
                GlobalLogs.logging.error(
                    "1. Required file not exists[DB容量使用状況_YYYYMMDD.xlsx] ,please place the required file")
                GlobalLogs.logging.error(
                    "2. Placed [DB容量使用状況_YYYYMMDD.xlsx] file is not current week's file, please place the current week's file")
                GlobalLogs.logging.error(
                    "3. Current week's output file already exists,please move the existing output file to bk & execute")

                sys.exit()

