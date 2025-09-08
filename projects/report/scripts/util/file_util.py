''' Various functions for manipulating files for cleanup '''
import os
import shutil
import subprocess
import logging
from pathlib import Path
import sys

logging.basicConfig(level=logging.INFO)

def mount_drive(drive_letter: str, mount_point: str) -> None:
    '''
    Mounts a Windows drive into WSL
    '''

    logging.info(f'Mounting {drive_letter} to {mount_point}...')
    
    try:
        subprocess.run(['sudo', 'mount', '-t', 'drvfs', drive_letter, mount_point], check=True)
        logging.info('Drive mounted successfully')
    except subprocess.CalledProcessError as e:
        logging.info(f'Failed to mount drive {drive_letter}: {e}')
        raise


def copy_input_folder(source_folder: str | Path, destination_folder: str | Path) -> None:
    '''
    Copies ALL files in source folder to destination folder. 
    
    Make sure source folder contains only relevant inputs files.

    Must mount drive first if copying from Google Drive folder
    '''

    #ensure inputs are Path objects
    src = Path(source_folder)
    dst = Path(destination_folder)

    dst.mkdir(parents=True, exist_ok=True)

    logging.info(f'Copying files from {source_folder} to {destination_folder}')

    for filename in src.iterdir():
        if filename.is_file():
            destination_path = dst / filename.name
            logging.info(f'Destination Path: "{destination_path}"')
            try:
                shutil.copy(filename, destination_path)
                logging.info(f'Copied "{filename.name}"')
            except Exception as e:
                logging.error(f'Failed to copy {filename.name}: {e}')

def copy_input_from_config(source_folder, destination_folder, files) -> None:
    '''
    Copies files specified in source folder to destination folder. 
    
    Files and folders are specified by provided loaded config

    Must mount drive first if copying from Google Drive folder
    '''

    #ensure inputs are Path objects
    src = Path(source_folder)
    dst = Path(destination_folder)

    dst.mkdir(parents=True, exist_ok=True)

    logging.info(f'Copying files from {source_folder} to {destination_folder}')

    for filename in files.values():
        file_path = src / Path(filename)
        destination_path = dst / filename
        logging.info(f'Destination Path: "{destination_path}"')
        try:
            shutil.copy(file_path, destination_path)
            logging.info(f'Copied "{filename}"')
        except Exception as e:
            logging.error(f'Failed to copy {filename}: {e}')

def find_repository_root(marker = 'pyproject.toml'):
    folder = Path(__file__).resolve().parent
    while folder != folder.parent:  # stop at filesystem root
        if (folder / marker).exists():
            return folder
        folder = folder.parent
    raise FileNotFoundError(f"Could not find repo root containing {marker}")

def make_directory(path):
    '''Create specified directory if not exist'''
    path.mkdir(parents=True, exist_ok=True)

def get_script_name():
    import __main__ as main
    '''Return the name of current script'''
    # script_path = sys.argv[0]
    script_path = main.__file__
    script_name = Path(script_path).stem
    return(script_name)

