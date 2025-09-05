'''Define functions for manipulating csv files'''
import csv
import logging
from pathlib import Path

def create_output_csv(path: Path | str, headers: list, logger: logging.Logger | None = None):
    '''Create csv and populate with specified headers'''

    path = Path(path)

    if logger is None:
        logger = logging.getLogger(__name__)

    with open(path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        logger.info(f'Ouput file created')