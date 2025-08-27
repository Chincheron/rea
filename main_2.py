from pathlib import Path
import file_util
import xlwings as xw
import openpyxl
import logging

def main():
    #Set Base directory to project root
    BASE_DIR = Path(__file__).resolve().parents[0]
    print(BASE_DIR)
    
    #mount google drive
    file_util.mount_drive('G:', '/mnt/g')

    #copy current REA version file 
    file_util.copy_input_files('/mnt/g/My Drive/Paper Prep/Leslie Matrix/Versions/Current Working Version', BASE_DIR)

    input_file = 'NEW CorrectedLeslieMatrix.2025.08.26.v1.2.xlsx'

    wb_rea = xw.Book(input_file)

    print(type(wb_rea))


    
if __name__ == "__main__":
    main()
