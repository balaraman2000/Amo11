import json
from CheckLastTime import CheckLastTime
from CompareFourFile import CompareForuFile


def main():
    with open("config.json", "r", encoding="utf-8") as config_file:
        config = json.load(config_file)
        object1 = CheckLastTime(first_file=config["PADCTRA1"], second_file=config["PADCTRS2"],
                               target_folder=config["targetFolderPADCTRA1"])

        object2 = CheckLastTime(first_file=config["PADDWRA1"], second_file=config["PADDWRS2"],
                               target_folder=config["targetFolderPADDWRA1"])

        object3 = CompareForuFile(first_File=config["PADNVRA1"], second_file=config["PADNVRS2"],
                                 third_file=config["PADNVRA3"], four_file=config["PADNVRS4"],
                                 target_folder=config["targetFolderPADNVRA1"])

        object4 = CheckLastTime(first_file=config["PADPRRA1"], second_file=config["PADPRRS2"],
                               target_folder=config["targetFolderPADPRRA1"])

        object5 = CompareForuFile(first_File=config["PADSCRA1"], second_file=config["PADSCRS2"],
                                 third_file=config["PADSCRA3"], four_file=config["PADSCRS4"],
                                 target_folder=config["targetFolderPADSCRA1"])

        object6 = CheckLastTime(first_file=config["PADUSRA1"], second_file=config["PADUSRS2"],
                               target_folder=config["targetFolderPADUSRA1"])

        print("Executed Succesfully")


main()
