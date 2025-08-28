from pathlib import Path
import file_util
import xlwings as xw
import openpyxl
import pandas as pd
import logging

def main():
    #Set Base directory to project root
    BASE_DIR = Path(__file__).resolve().parents[0]
    print(BASE_DIR)
    
    #copy current REA version file 
    # file_util.copy_input_files('G:\My Drive\Paper Prep\Leslie Matrix\Versions\Current Working Version', BASE_DIR)

    rea_file = 'NEW CorrectedLeslieMatrix.2025.08.26.v1.3.1.xlsx'
    input_file = 'Example_input_file.csv'

    #load input file
    scenarios = pd.read_csv(input_file)
    logging.info(f'Loaded {len(scenarios)} from scenarios input file')

    for index, row in scenarios.iterrows(): #consider switching to iterating over tuples if performance becomes issue
        number_killed = row['number_killed']
        discount_factor = row['discount_factor']
        discount_year = row['discount_start_year']
        max_age = row['maximum_age']
        logging.info(f'Scenario {index +1} inputs:\n' 
                     f'Number Killed set to {number_killed}\n'
                     f'Discount factor set to {discount_factor}\n'
                     f'Base year set to {discount_year}\n'
                     f'Max age set to {max_age}'
                     )

    # #load workbook
    # wb_rea = xw.Book(rea_file)

    # #load sheet with inputs and outputs
    # io_sheet = wb_rea.sheets['Debit Inputs']

    # #create variables for inputs for scenario
    # max_age = 50
    # base_year = 2016

    # #set cells to scenario inputs
    # io_sheet['M12'].value = max_age
    # io_sheet['M11'].value = base_year

    # #force excel to recalculate
    # wb_rea.app.calculate()

    # #output next - need 1)direct DMSY lsot 2) Indirect DMSY lost 3) Totl DMSY Lost 4) DMSY Restored and 5) annual reelase over 1o eyars
    # #last one will be trickiest need to solve for it each time?

    # #outputs to variables
    # direct_loss_total = io_sheet['T3'].value
    # indirect_loss_total_exclude = io_sheet['T4'].value
    # loss_total = io_sheet['T5'].value
    # gains_total = io_sheet['T7'].value

    # logging.info(f'Outputs copied to variable\n' 
    #              f'  Direct loss: {direct_loss_total}\n'
    #              f'  Indirect loss: {indirect_loss_total_exclude}\n'
    #              f'  Total Loss: {loss_total}\n'
    #              f'  Gain: {gains_total}')

    # #save results
    # with open('results.csv', 'w') as f:
    #     f.write('Maximum Age, Base Year, Direct Loss, Indirect Loss, Total Loss, Total Gains\n')
    #     f.write(f'{max_age}, {base_year}, {direct_loss_total}, {indirect_loss_total_exclude}, {loss_total}, {gains_total} \n')

    # #close excel instance
    # wb_rea.app.quit()
    
if __name__ == "__main__":
    main()
