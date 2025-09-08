'''Define functions for manipulating csv files'''
import csv
import logging
from pathlib import Path

def create_output_csv(path: Path | str, headers: list | dict, logger: logging.Logger | None = None):
    '''Create csv and populate with specified headers'''

    if isinstance(headers, dict):
        headers = list(headers.keys()) 
        
    path = Path(path)

    if logger is None:
        logger = logging.getLogger(__name__)

    with open(path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        logger.info(f'Ouput file created')

def append_output_to_csv(path, row_data):
    '''Append data to csv. Make sure data matches'''
    with open(path, 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(list(row_data))
        