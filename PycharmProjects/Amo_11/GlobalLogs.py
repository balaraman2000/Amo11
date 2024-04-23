import logging
import sys
from datetime import datetime
from pathlib import Path
import json
import pywintypes

try:
    with open("config.json", "r", encoding="utf-8") as config_file:
        try:
            config = json.load(config_file)
        except json.decoder.JSONDecodeError as e:
            print("Error: decoding JSON")
            sys.exit()

    output_dir_path = config.get("logfile")
    if not output_dir_path:
        raise ValueError("Config file does not contain a valid  path")


    output_dir = Path(output_dir_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = output_dir / f'DB容量使用状況作成_log_{current_datetime}.txt'
    logging.basicConfig(filename=log_file_name, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

except FileNotFoundError:
    print("Error: Config file not found.")
    logging.error("Config file not found.")
    sys.exit()

except PermissionError:
    print("Error: Permission denied.")
    logging.error("Permission denied.")
    sys.exit()

except pywintypes.com_error as e:
    print("Error: Sheet not found:", e)
    logging.error("Sheet not found: %s", e)
    sys.exit()

except KeyError as e:
    print("Error: 'logfile' key not found in config file:", e)
    logging.error("KeyError: 'logfile' key not found in config file: %s", e)
    sys.exit()

except ValueError as e:
    print("Error:", e)
    logging.error("Error: %s", e)
    sys.exit()

except Exception as e:
    print("An error occurred:", e)
    logging.error("An error occurred: %s", e)
    sys.exit()
