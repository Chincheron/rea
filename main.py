from pathlib import Path
import file_util
import xlwings as xw
import openpyxl
import logging

def main():
    #Set Base directory to project root
    BASE_DIR = Path(__file__).resolve().parents[0]
    print(BASE_DIR)
    
    #copy current REA version file 
    file_util.copy_input_files('G:\My Drive\Paper Prep\Leslie Matrix\Versions\Current Working Version', BASE_DIR)

    input_file = 'NEW CorrectedLeslieMatrix.2025.08.26.v1.3.1.xlsx

    #load workbook
    wb_rea = xw.Book(input_file)

    #load sheet with inputs and outputs
    io_sheet = wb_rea.sheets['Debit Inputs']

    #create variables for inputs for scenario
    max_age = 50
    base_year = 2016

    #set cells to scenario inputs
    io_sheet['M12'].value = max_age
    io_sheet['M11'].value = base_year

    #force excel to recalculate
    wb_rea.app.calculate()

    #output next - need 1)direct DMSY lsot 2) Indirect DMSY lost 3) Totl DMSY Lost 4) DMSY Restored and 5) annual reelase over 1o eyars
    #last one will be trickiest need to solve for it each time?

    #outputs to variables
    direct_loss_total = io_sheet['T3'].value
    print(direct_loss_total)

    #save results
    with open('results.csv', 'w') as f:
        f.write('input,output\n')
        f.write(f'{max_age}, {direct_loss_total}\n')

    #close excel instance
    wb_rea.app.quit()
    
if __name__ == "__main__":
    main()
