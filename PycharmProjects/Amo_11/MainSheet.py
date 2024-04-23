import json
import GlobalLogs
from CheckLastTime import CheckLastTime
from ClickMacro import ClickMacro
from CompareFourFile import CompareForuFile
import sys
import pywintypes
from Demo import CharSheet
from ValidFile import Vaild_file
from open import  OpenExcel


def main():
    OpenExcel()

    try:
        print("AMO_011 AUTOMATION-DB容量使用状況作成")
        GlobalLogs.logging.info("AMO_11 Automation - Start of Log")
        GlobalLogs.logging.info("AMO_011 AUTOMATION-DB容量使用状況作成")
        print("Enter Y to RUN AUTOMATION")
        print("Enter N to Exit AUTOMATION")
        choice = input("Please enter : ")
        if choice.lower() == 'y':

            GlobalLogs.logging.info("user input :  " + choice)
            with open("config.json", "r", encoding="utf-8") as config_file:
                try:
                    config = json.load(config_file)

                except json.decoder.JSONDecodeError as e:
                    print("Error decoding JSON:")
                    GlobalLogs.logging.error("Error decoding JSON:", e)
                    sys.exit()

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
                Vaild_file(directory=config["graph_path"])
                object = ClickMacro(file_path=config["file_path"], macro_name=config["macro_name"],
                                    graph_path=config["graph_path"])
                object = CharSheet(directory=config["graph_path"])
                GlobalLogs.logging.info("AMO_11 Automation - End of Log")
                print("Executed Successfully")
                GlobalLogs.logging.info("Executed Successfully")
        elif choice.lower() == 'n':
            GlobalLogs.logging.info("user input :  " + choice)
            sys.exit()
        else:
            print("\n\nEnter the correct input\n")
            GlobalLogs.logging.info("Enter the correct input  ")
            GlobalLogs.logging.info("mismatch user input :  " + choice)
            main()

    except KeyError as e:
        print("KeyError:", e)
        GlobalLogs.logging.error("KeyError: %s", e)
        sys.exit()

    except FileNotFoundError:
        print("Error: Source file not found.")
        GlobalLogs.logging.exception("Source file not found")
        sys.exit()
    except PermissionError:
        print("Error: Permission denied.")
        GlobalLogs.logging.exception("Permission denied")
        sys.exit()

    except pywintypes.com_error as e:
        print("Error: Sheet not found")
        GlobalLogs.logging.error("Sheet not found")
        sys.exit()

    except Exception as e:
        print("An error occurred:", e)
        GlobalLogs.logging.error("An error occurred:")
        sys.exit()

main()
