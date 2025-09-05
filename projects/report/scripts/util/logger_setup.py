from pathlib import Path
import logging
from datetime import datetime

def setup_loggers(directory):
    '''Setup loggers for different logging info'''

    # Create timestamped folder for this run
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_dir = directory / Path(f'run_{timestamp}') / 'logs'
    log_dir.mkdir(parents=True, exist_ok=True)

    #main logger for general info
    main_logger = logging. getLogger('main')
    main_logger.setLevel(logging.INFO)
    main_handler = logging.FileHandler(log_dir / 'main_log.log')
    main_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    main_handler.setFormatter(main_formatter)
    main_logger.addHandler(main_handler)

    # Warning logger for warnings and errors
    warning_logger = logging.getLogger('warnings')
    warning_logger.setLevel(logging.WARNING)
    warning_handler = logging.FileHandler(log_dir / 'warnings.log')
    warning_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    warning_handler.setFormatter(warning_formatter)
    warning_logger.addHandler(warning_handler)
        
    # Detailed logger for detailed variable outputs
    detail_logger = logging.getLogger('details')
    detail_logger.setLevel(logging.INFO)
    detail_handler = logging.FileHandler(log_dir / 'detailed_outputs.log')
    detail_formatter = logging.Formatter('%(asctime)s - %(message)s')
    detail_handler.setFormatter(detail_formatter)
    detail_logger.addHandler(detail_handler)

    #setup console logger for logging to main log file AND console
    console_logger = logging.getLogger('console')
    console_logger.setLevel(logging.INFO)
    #file specific setup
    console_file_handler = logging.FileHandler(log_dir / 'main_log.log')
    console_file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_file_handler.setFormatter(console_file_formatter)
    console_logger.addHandler(console_file_handler)
    #terminal setup
    console_handler = logging.StreamHandler()
    console_formatter = logging.Formatter('%(message)s')
    console_handler.setFormatter(console_formatter)
    console_logger.addHandler(console_handler)

    main_logger.propagate = False
    warning_logger.propagate = False
    detail_logger.propagate = False
    console_logger.propagate = False

    return main_logger, warning_logger, detail_logger, console_logger