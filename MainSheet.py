import json

import GlobalLogs
from CheckLastTime import CheckLastTime
from ClickMacro import ClickMacro
from CompareFourFile import CompareForuFile
import sys


def main():
    print("AMO_011 AUTOMATION-DB容量使用状況作成")
    GlobalLogs.logging.info("AMO_11 Automation - Start of Log")
    GlobalLogs.logging.info("AMO_011 AUTOMATION-DB容量使用状況作成")
    print("Enter Y to RUN AUTOMATION")
    print("Enter N to Exit AUTOMATION")
    choice = input("Please enter : ")
    if choice.lower() == 'y':
        GlobalLogs.logging.info("user input :  " + choice)

        with open("config.json", "r", encoding="utf-8") as config_file:
            config = json.load(config_file)
            print("Check when the file was last modified...  ")
            object = CheckLastTime(first_file=config["PADCTRA1"], second_file=config["PADCTRS2"],
                                   target_folder=config["targetFolderPADCTRA1"])

            object = CheckLastTime(first_file=config["PADDWRA1"], second_file=config["PADDWRS2"],
                                   target_folder=config["targetFolderPADDWRA1"])

            object = CompareForuFile(first_File=config["PADNVRA1"], second_file=config["PADNVRS2"],
                                     third_file=config["PADNVRA3"], four_file=config["PADNVRS4"],
                                     target_folder=config["targetFolderPADNVRA1"])

            object = CheckLastTime(first_file=config["PADPRRA1"], second_file=config["PADPRRS2"],
                                   target_folder=config["targetFolderPADPRRA1"])

            object = CompareForuFile(first_File=config["PADSCRA1"], second_file=config["PADSCRS2"],
                                     third_file=config["PADSCRA3"], four_file=config["PADSCRS4"],
                                     target_folder=config["targetFolderPADSCRA1"])

            object = CheckLastTime(first_file=config["PADUSRA1"], second_file=config["PADUSRS2"],
                                   target_folder=config["targetFolderPADUSRA1"])

            object = ClickMacro(file_path=config["file_path"], macro_name=config["macro_name"])
            GlobalLogs.logging.info("AMO_11 Automation - End of Log")

            print("Executed Succesfully")
            GlobalLogs.logging.info("Executed Succesfully")
    elif choice.lower() == 'n':
        GlobalLogs.logging.info("user input :  " + choice)
        sys.exit()
    else:
        print("\n\nEnter the correct input\n")
        GlobalLogs.logging.info("Enter the correct input  ")
        GlobalLogs.logging.info("mismatch user input :  " + choice)
        main()


main()
