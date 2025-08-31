from pathlib import Path
from datetime import datetime
import file_util
import xlwings as xw
import openpyxl
import pandas as pd
import logging
import csv
#TODO: 1) calculate annual reintroduction
# 2) logging message to file

def setup_loggers():
    '''Setup loggers for different logging info'''

    # Create timestamped folder for this run
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_dir = Path('logs') / f'run_{timestamp}'
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

def main():
    main_logger, warning_logger, detail_logger, console_logger  = setup_loggers()

    # initial constants
    BASE_DIR = Path(__file__).resolve().parents[0] #set to root directory
    main_logger.info(f'Base directory set')
    TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')


    #set inputs
    rea_file = 'Mussel REA_v2.0.xlsx'
    input_file = 'Example_input_file.csv'
    copy_dir = 'G:\My Drive\Paper Prep\Leslie Matrix\Versions\Current Working Version'
    
    output_dir = Path('output') / f'run_{TIMESTAMP}'
    output_dir.mkdir(parents=True, exist_ok=True)
    
    input_dir = Path('inputs') / f'run_{TIMESTAMP}'
    input_dir.mkdir(parents=True, exist_ok=True)
    input_file = input_dir / input_file
    rea_file = input_dir / rea_file

    
    #copy current REA version file 
    file_util.copy_input_files(copy_dir, input_dir)
    main_logger.info(f'Input files copied from current working version folder')

    #load input file
    scenarios = pd.read_csv(input_file)
    main_logger.info(f'Loaded {len(scenarios)} from scenarios input file')

    #load workbook
    wb_rea = xw.Book(rea_file)
    main_logger.info(f'REA model workbook loaded ({rea_file})')

    excel_application = wb_rea.app.api

    #load sheet with inputs and outputs
    io_sheet = wb_rea.sheets['Debit Inputs']
    main_logger.info(f'Scenario input file loaded ({input_file})')

    fail_scenario_written = False

    #open output csv
    with open(output_dir / 'output.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Number Killed', 'Discount Factor', 'Base Year', 'Maximum Age',
                          'Direct Loss', 'Indirect Loss', 'Total Loss', 'Total Gains', 'Annual Reintroduction'])
        main_logger.info(f'Ouput file created')
           
        for index, row in scenarios.iterrows(): #consider switching to iterating over tuples if performance becomes issue
            scenario_number = index + 1
            number_killed = row['number_killed']
            discount_factor = row['discount_factor']
            base_year = row['discount_start_year']
            max_age = row['maximum_age']
            main_logger.info(f'Scenario {scenario_number}: Inputs loaded')
            detail_logger.info(f'Scenario {scenario_number}: Inputs:\n' 
                        f'Number Killed set to {number_killed}\n'
                        f'Discount factor set to {discount_factor}\n'
                        f'Base year set to {base_year}\n'
                        f'Max age set to {max_age}'
                        )

            #set cells to scenario inputs
            io_sheet['M8'].value = number_killed
            io_sheet['M16'].value = discount_factor
            io_sheet['M11'].value = base_year
            io_sheet['M12'].value = max_age
            main_logger.info(f'Scenario {scenario_number}: Excel cells set to scenario inputs')
            
            #use goal seek to determine number of annual reintroductions needed for gain to equal loss
            # Goal Seek: set Goal:Loss ratio (T8) to 1 by changing Annual Mussel Reintroduction (M21)
            set_cell = io_sheet.range('T8').api 
            to_value = 1                          
            by_changing_cell = io_sheet.range('M21').api 
            set_cell.GoalSeek(Goal=to_value, ChangingCell=by_changing_cell)
            main_logger.info(f'Scenario {scenario_number}: Required annual reintroduction calculated (for gain to equal loss)')

            #No such thing as partial mussel so round annual mussel reintroduction down to nearest whole number and set cell to value
            whole_annual_reintroduction = io_sheet['M21'].value
            detail_logger.info(f'Scenario {scenario_number}: Exact annual reintroduction: {whole_annual_reintroduction}')
            whole_annual_reintroduction = int(whole_annual_reintroduction)
            detail_logger.info(f'Scenario {scenario_number}: Rounded Annual reintroduction: {whole_annual_reintroduction}')
            io_sheet['M21'].value = whole_annual_reintroduction

            #force excel to recalculate
            wb_rea.app.calculate()
            main_logger.info(f'Scenario {scenario_number}: Excel workbook recalculated')

            #outputs to variables
            direct_loss_total = io_sheet['T3'].value
            indirect_loss_total_exclude = io_sheet['T4'].value
            loss_total = io_sheet['T5'].value
            gains_total = io_sheet['T7'].value
            annual_reintroduction = io_sheet['M21'].value
            main_logger.info(f'Scenario {scenario_number}: Excel outputs copied to variable')
            
            #append results
            writer.writerow([
                row['number_killed'],
                row['discount_factor'],
                row['discount_start_year'],
                row['maximum_age'],
                direct_loss_total,
                indirect_loss_total_exclude,
                loss_total,
                gains_total,
                annual_reintroduction
            ])
            main_logger.info(f'Scenario {scenario_number}: Excel outputs written to output file')
            detail_logger.info(f'Scenario {scenario_number}: Excel outputs written to output file:\n' 
                        f'  Direct loss: {direct_loss_total}\n'
                        f'  Indirect loss: {indirect_loss_total_exclude}\n'
                        f'  Total Loss: {loss_total}\n'
                        f'  Gain: {gains_total}\n'
                        f'  Annual Reintroduction: {annual_reintroduction}'
                        )


            #check QC tests
            qc_test = io_sheet['N24'].value
            main_logger.info(f'Checking whether Excel workbook QC tests pass')
            if qc_test == 'PASS':
                main_logger.info(f'Scenario {scenario_number}: QC test passed')
            else: 
                warning_logger.warning(f'Scenario {scenario_number}: QC test failed')
                
                if fail_scenario_written == False:
                    with open(output_dir / 'failed_scenario.csv', 'w', newline='') as fail_file:
                        fail_writer = csv.writer(fail_file)
                        fail_writer.writerow(['Scenario', 'Number Killed', 'Discount Factor', 'Base Year', 'Maximum Age',
                                'Direct Loss', 'Indirect Loss', 'Total Loss', 'Total Gains', 'Annual Reintroduction'])
                    fail_scenario_written = True
                    warning_logger.warning(f'Scenario {scenario_number}: Created failed scenario output file')
                with open('failed_scenario.csv', 'a', newline='') as fail_file:
                        fail_writer = csv.writer(fail_file)
                        fail_writer.writerow([
                            index + 1,
                            row['number_killed'],
                            row['discount_factor'],
                            row['discount_start_year'],
                            row['maximum_age'],
                            direct_loss_total,
                            indirect_loss_total_exclude,
                            loss_total,
                            gains_total,
                            annual_reintroduction
                        ])
                        warning_logger.warning(f'Scenario {scenario_number}: Failed scenario inputs/outputs written to failed scenario outputs file ')

            main_logger.info(f'Scenario {scenario_number}: Scenario completed')
            console_logger.info(f'{scenario_number}/{len(scenarios)} complete')

    #close excel instance
    wb_rea.app.quit()
    main_logger.info(f'Closed excel instance')
    main_logger.info(f'Script finished')

if __name__ == "__main__":
    main()
