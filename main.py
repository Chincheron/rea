from pathlib import Path
from datetime import datetime
import file_util
import xlwings as xw
import pandas as pd
import logger_setup
import json
import time
import excel_util as xl
import csv_util

#TODO 1) modify file copy so that it only copies the specifi input and rea files (pass those file names to function probably)
#TODO 2) Add flag/option in config file to copy files or directly read from inputs folder   

def load_config(config_path='test_config.json'):
    """Load configuration from JSON file"""
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config

def main():
    try:
        config = load_config()

        main_logger, warning_logger, detail_logger, console_logger  = logger_setup.setup_loggers()

        # initial constants
        BASE_DIR = Path(__file__).resolve().parents[0] #set to root directory
        main_logger.info(f'Base directory set')
        TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')
        START_TIME = time.perf_counter()

        #settings from config
        files = config['files']
        directories = config['directories']
        excel_config = config['excel']
        goal_seek_config = config['goal_seek']
        input_cells_config = excel_config['input_cells']
        output_cells_config = excel_config['output_cells']
        sheets_config = excel_config['sheet_name']

        rea_file = files['rea_file']
        scenario_file = files['input_file']
        copy_dir = directories['copy_source']
        
        output_dir = Path(f'run_{TIMESTAMP}') / Path(directories['output_folder'])
        output_dir.mkdir(parents=True, exist_ok=True)
        output_file = output_dir / Path('scenario_output.csv')
        
        input_dir = Path(f'run_{TIMESTAMP}') / Path(directories['input_folder'])
        input_dir.mkdir(parents=True, exist_ok=True)
        scenario_file = input_dir / scenario_file
        rea_file = input_dir / rea_file

        #copy current REA version file 
        file_util.copy_input_files(copy_dir, input_dir)
        main_logger.info(f'Input files copied from current working version folder')

        #load scenario input file
        scenarios = pd.read_csv(scenario_file)
        main_logger.info(f'Loaded {len(scenarios)} from scenarios input file into dataframe')

        #load REA model and I/O sheet
        wb_rea, app = xl.load_workbook(rea_file)
        main_logger.info(f'REA model workbook loaded ({rea_file})')
        io_sheet = xl.load_worksheet(wb_rea, sheets_config['input_sheet'], warning_logger)
        main_logger.info(f'REA input sheet loaded ({io_sheet})')

        # for loop runs through different scenarios and:
        # 1) Sets inputs
        # 2) Solves for number of annual reintroductions required for gains to equal losses
        # 3) Reads desired outputs and writes both inputs and outputs to an output csv for later processing
        # 4) QC check of REA QC tests
        for scenario_number, row in enumerate(scenarios.itertuples(index=False), start =1): 
            
            #Step #1: Set inputs
            #read specified input values from the scenarios input dataframe for each scenario
            inputs = {
                'number_killed' : row.number_killed,
                'discount_factor' : row.discount_factor,
                'base_year' : row.discount_start_year, 
                "max_age" : row.maximum_age
            }  
            main_logger.info(f'Scenario {scenario_number}: Inputs loaded')

            xl.set_excel_inputs(io_sheet, inputs, input_cells_config, scenario_number, main_logger)
            detail_logger.info(f'Scenario {scenario_number}: Inputs:\n' 
                f'Number Killed set to {row.number_killed}\n'
                f'Discount factor set to {row.discount_factor}\n'
                f'Base year set to {row.discount_start_year}\n'
                f'Max age set to {row.maximum_age}'
                )
            
            #Step #2: Solves for number of annual reintroductions required for gains to equal losses
            # Goal Seek: set Goal:Loss ratio to 1 by changing Annual Mussel Reintroduction 
            xl.run_goal_seek(io_sheet, input_cells_config['loss_ratio'], input_cells_config['annual_reintroduction'], goal_seek_config['target_value'])
            main_logger.info(f'Scenario {scenario_number}: Required annual reintroduction calculated (for gain to equal loss)')

            #No such thing as partial mussel so round annual mussel reintroduction down to nearest whole number and set cell to value
            annual_reintroduction_exact = round(io_sheet[input_cells_config['annual_reintroduction']].value, 2)
            detail_logger.info(f'Scenario {scenario_number}: Exact annual reintroduction: {annual_reintroduction_exact}')
            annual_reintroduction_rounded = int(annual_reintroduction_exact)
            detail_logger.info(f'Scenario {scenario_number}: Rounded Annual reintroduction: {annual_reintroduction_rounded}')
            io_sheet[input_cells_config['annual_reintroduction']].value =annual_reintroduction_exact

            #Step #3: Reads desired outputs and writes both inputs and outputs to an output csv for later processing
            #force excel to recalculate
            wb_rea.app.calculate()
            main_logger.info(f'Scenario {scenario_number}: Excel workbook recalculated')

            #read model outputs and append to csv file
            outputs = xl.read_excel_outputs(io_sheet, output_cells_config, 0, main_logger)
            csv_data = {'Scenario_number': scenario_number, **inputs, **outputs, 'Annual Reintroduction Rounded': annual_reintroduction_rounded, 'Annual Reintroduction Exact': annual_reintroduction_exact}
            if not output_file.exists():
                csv_util.create_output_csv(output_file, csv_data)
            csv_util.append_output_to_csv(output_file, list(csv_data.values()))
            
            #detail logging of scenario outputs
            log_lines = [f'Scenario {scenario_number}: Excel outputs written to output file:']
            #read and return all outputs from config file
            for key in output_cells_config.keys():
                if key in csv_data:
                    clean_key = key.replace('_', ' ').title()
                    log_lines.append(f'  {clean_key}: {csv_data[key]}')
            #inclue outputs created in processing (not in config)
            log_lines.append(f'  Annual Reintroduction Rounded: {annual_reintroduction_rounded}')
            log_lines.append(f'  Annual Reintroduction Exact: {annual_reintroduction_exact}')
            detail_logger.info('\n'.join(log_lines))

            # Step #4: QC check of REA QC tests
            #Excel sheet has multiple qc tests whose results are summarized in a single cell as either 'PASS' or 'FAIL'
            #Check this cell and write I/O to file if 'FAIL' for review
            xl.check_qc(io_sheet, input_cells_config['qc_test'], output_dir, csv_data, scenario_number, main_logger, warning_logger)
           
            main_logger.info(f'Scenario {scenario_number}: Scenario completed')
            console_logger.info(f'{scenario_number}/{len(scenarios)} complete')
            
    finally:
        #close excel instance
        wb_rea.close()
        app.quit() 
        main_logger.info(f'Closed excel instance')

    #Measure elasped runtime of script
    END_TIME = time.perf_counter()
    RUN_TIME = END_TIME - START_TIME
    RUN_MINUTES = RUN_TIME // 60
    RUN_SECONDS = round((RUN_TIME % 60), 1 )

    console_logger.info(f'Script finished. Total Runtime: {RUN_MINUTES} minutes and {RUN_SECONDS} seconds')

if __name__ == "__main__":
    main()
