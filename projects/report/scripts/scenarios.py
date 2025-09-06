from pathlib import Path
from datetime import datetime
import util.file_util as file_util
import xlwings as xw
import pandas as pd
import util.logger_setup as logger_setup
import json
import time
import util.excel_util as xl
import util.csv_util as csv_util
import models.rea.inputs as rea_input_class
import util.analysis_util as analysis_util

#TODO 1) modify file copy so that it only copies the specifi input and rea files (pass those file names to function probably)
#TODO 2) Add flag/option in config file to copy files or directly read from inputs folder   
#TODO REfactor so that main can be called: pull CONSTANTS to constants file; incorporate load config into main?;      

if __name__ == "__main__":
    CONFIG_FILE = 'scenarios_config.json'
    REPO_DIR = file_util.find_repository_root()
    PROJECT_BASE_DIR = (REPO_DIR / 'projects' / 'report')
    CONFIG_DIR = PROJECT_BASE_DIR / 'config'
    CONFIG_PATH = CONFIG_DIR / CONFIG_FILE
    RESULTS_DIR = (PROJECT_BASE_DIR / 'results')
    analysis_util.run_rea_scenario_total(CONFIG_PATH)

