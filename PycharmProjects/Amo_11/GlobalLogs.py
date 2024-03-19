import logging
from datetime import datetime
from logging.handlers import SysLogHandler
from pathlib import Path
import json

with open("config.json", "r", encoding="utf-8") as config_file:
    config = json.load(config_file)
output_dir = Path(config["logfile"])
output_dir.mkdir(parents=True, exist_ok=True)
current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file_name = output_dir / f'DB容量使用状況作成_log_{current_datetime}.txt'
logging.basicConfig(filename=log_file_name, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
