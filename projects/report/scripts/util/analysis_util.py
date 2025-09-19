'''Functions for running full analyses'''
from pathlib import Path
from datetime import datetime
import util.file_util as file_util
import pandas as pd
import util.logger_setup as logger_setup
import time
import util.excel_util as xl
import util.csv_util as csv_util
import models.rea.inputs as rea_input_class
import util.config as config_utl
from util.constants import * 
import util.config as config_util
import util.data_util as data_util
import util.math_util as math_util

def run_rea_scenario_total(config_file: Path | str):
    '''Runs REA based on scenario input file and returns total outputs (i.e. single cell outputs) '''
    
    wb_rea = None
    app = None
    try:
        # initial constants
        CONFIG_FILE = config_file
        CONFIG_PATH = CONFIG_DIR / CONFIG_FILE
        TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')
        START_TIME = time.perf_counter()
        config = config_utl.load_config(CONFIG_PATH)

        #get script name of root script to use to create directory to hold results for each run of script (e.g., scenarios.py returns 'Scenarios')
        SCRIPT_NAME = file_util.get_script_name()
        script_run_results_directory = Path(RESULTS_DIR) / Path(f'{SCRIPT_NAME}_{TIMESTAMP}')
        #setup logger
        main_logger, warning_logger, detail_logger, console_logger  = logger_setup.setup_loggers(script_run_results_directory)

        #settings from config
        files = config['files']
        directories = config['directories']
        excel_config = config['excel']
        misc_config = config['misc']
        goal_seek_config = config['goal_seek']
        input_cells_config = excel_config['input_cells']
        output_cells_config = excel_config['output_cells_excluded']
        sheets_config = excel_config['sheet_name']

        decimal_precision_results = misc_config['result_decimal_precision']

        rea_file = files['rea_file']
        scenario_file = files['input_file']
        copy_dir = directories['copy_source']
        
        output_dir = RESULTS_DIR / Path(script_run_results_directory) / Path(directories['output_folder'])
        output_dir.mkdir(parents=True, exist_ok=True)
        output_file = output_dir / Path('scenario_output.csv')
        
        input_dir = RESULTS_DIR / Path(script_run_results_directory) / Path(directories['input_folder'])
        input_dir.mkdir(parents=True, exist_ok=True)
        scenario_file = input_dir / scenario_file
        rea_file = input_dir / rea_file


        #copy current REA version file 
        file_util.copy_input_from_config(copy_dir, input_dir, files)
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
            # Create scenario inputs class with defaults values
            scenario_inputs = rea_input_class.REAScenarioInputs.create_from_config(CONFIG_PATH, debug=True)
            # and override with input values from the scenarios input dataframe for each scenario
            scenario_inputs.update_from_row(row)
            #must convert to dict for easier reading into later functions
            scenario_inputs_dict = scenario_inputs.to_dict()
            main_logger.info(f'Scenario {scenario_number}: Inputs loaded')

            xl.set_excel_inputs(io_sheet, scenario_inputs_dict, input_cells_config, scenario_number, main_logger)
            #detail logging of scenario inputs
            log_lines = [f'Scenario {scenario_number}: Excel inputs set:']
            #read and return all inputs from config file and scenario inputs override
            for key in input_cells_config.keys():
                if key in scenario_inputs_dict:
                    clean_key = key.replace('_', ' ').title()
                    log_lines.append(f'  {clean_key} set to: {scenario_inputs_dict[key]}')
            detail_logger.info('\n'.join(log_lines))
            
            #Step #2: Solves for number of annual reintroductions required for gains to equal losses
            # Goal Seek: set Goal:Loss ratio to 1 by changing Annual Mussel Reintroduction 
            xl.run_goal_seek(io_sheet, input_cells_config['loss_ratio'], input_cells_config['annual_reintroduction'], goal_seek_config['target_value'])
            main_logger.info(f'Scenario {scenario_number}: Required annual reintroduction calculated (for gain to equal loss)')

            #No such thing as partial mussel so round annual mussel reintroduction down to nearest whole number and set cell to value
            annual_reintroduction_exact = math_util.round_outputs(io_sheet[input_cells_config['annual_reintroduction']].value, decimal_precision_results)
            detail_logger.info(f'Scenario {scenario_number}: Exact annual reintroduction: {annual_reintroduction_exact}')
            annual_reintroduction_rounded = math_util.round_annual_reintro(annual_reintroduction_exact)
            detail_logger.info(f'Scenario {scenario_number}: Rounded Annual reintroduction: {annual_reintroduction_rounded}')
            total_gain_exact = math_util.round_outputs(io_sheet[output_cells_config['total_gains']].value, decimal_precision_results) 
            io_sheet[input_cells_config['annual_reintroduction']].value =annual_reintroduction_rounded

            #Step #3: Reads desired outputs and writes both inputs and outputs to an output csv for later processing
            #force excel to recalculate
            wb_rea.app.calculate()
            main_logger.info(f'Scenario {scenario_number}: Excel workbook recalculated')

            #read model outputs and append to csv file
            outputs = xl.read_excel_outputs(io_sheet, output_cells_config, decimal_precision_results, main_logger)
            csv_data = {'Scenario_number': scenario_number, **scenario_inputs_dict, **outputs, 'total_gains_exact': total_gain_exact, 'Annual Reintroduction Rounded': annual_reintroduction_rounded, 'Annual Reintroduction Exact': annual_reintroduction_exact}
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
        if wb_rea: wb_rea.close()
        if app: app.quit() 
        main_logger.info(f'Closed excel instance')

    #Measure elasped runtime of script
    END_TIME = time.perf_counter()
    RUN_TIME = END_TIME - START_TIME
    RUN_MINUTES = int(RUN_TIME // 60)
    RUN_SECONDS = round((RUN_TIME % 60), 1 )

    console_logger.info(f'Script finished. Total Runtime: {RUN_MINUTES} minutes and {RUN_SECONDS} seconds')

def run_rea_scenario_yearly(config_file: Path | str):
    '''Runs REA based on scenario input file (excel) and returns yearly outputs (i.e. ranged cell outputs) '''
    
    wb_rea = None
    app = None
    try:
        # initial constants
        CONFIG_FILE = config_file
        CONFIG_PATH = CONFIG_DIR / CONFIG_FILE
        TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')
        START_TIME = time.perf_counter()
        config = config_utl.load_config(CONFIG_PATH)

        #get script name of root script to use to create directory to hold results for each run of script (e.g., scenarios.py returns 'Scenarios')
        SCRIPT_NAME = file_util.get_script_name()
        script_run_results_directory = Path(RESULTS_DIR) / Path(f'{SCRIPT_NAME}_{TIMESTAMP}')
        #setup logger
        main_logger, warning_logger, detail_logger, console_logger  = logger_setup.setup_loggers(script_run_results_directory)

        #settings from config
        files = config['files']
        directories = config['directories']
        excel_config = config['excel']
        misc_config = config['misc']
        goal_seek_config = config['goal_seek']
        input_cells_config = excel_config['input_cells']
        output_cells_config = excel_config['output_cells_excluded_yearly']
        config_folder = directories['config_folder']
        sheets_config = excel_config['sheet_name']

        decimal_precision_results = misc_config['result_decimal_precision']

        rea_file = files['rea_file']
        scenario_file = files['input_file']
        copy_dir = directories['copy_source']
        
        output_dir = RESULTS_DIR / Path(script_run_results_directory) / Path(directories['output_folder'])
        output_dir.mkdir(parents=True, exist_ok=True)
        output_file = output_dir / Path('scenario_output.xlsx')
        output_input_dir = output_dir / Path('scenario_inputs')
        output_input_dir.mkdir(parents=True, exist_ok=True)
        
        input_dir = RESULTS_DIR / Path(script_run_results_directory) / Path(directories['input_folder'])
        input_dir.mkdir(parents=True, exist_ok=True)
        scenario_file = input_dir / scenario_file
        rea_file = input_dir / rea_file
        
        config_folder = PROJECT_BASE_DIR / directories['config_folder']


        #copy current REA version file 
        file_util.copy_input_from_config(copy_dir, input_dir, files)
        main_logger.info(f'Input files copied from current working version folder')

        #load REA model and I/O sheet
        wb_rea, app = xl.load_workbook(rea_file)
        main_logger.info(f'REA model workbook loaded ({rea_file})')
        io_sheet = xl.load_worksheet(wb_rea, sheets_config['input_sheet'], warning_logger)
        main_logger.info(f'REA input sheet loaded ({io_sheet})')

        
        #for each config file in config folder, run specified scenarios
        for file in config_folder.iterdir():
            figure_config = config_util.load_config(file)
            figure_worksheet = figure_config['worksheet_name']
            desired_yearly_outputs = figure_config['desired_outputs']['output_cells_excluded_yearly']
            desired_yearly_outputs = {key:value for key, value in desired_yearly_outputs.items() if value == 'True'}    
            main_logger.info(f'Exhibit "{file.stem}": Desired outputs: \n'
                             f'{desired_yearly_outputs}')

            #update output_cells_config based on desired yearly outputs values for this sheet
            #must reinitiate both variables for each figure
            output_cells_config = excel_config['output_cells_excluded_yearly'].copy()
            keys_to_delete = [key for key in output_cells_config if key not in desired_yearly_outputs]
            detail_logger.info(f'Exhibit "{file.stem}": keys to delete \n'
                               f'{keys_to_delete}')
            for key in keys_to_delete:
                del output_cells_config[key]
            detail_logger.info(f'Exhibit "{file.stem}": Outputs to read:\n'
                               f'   {output_cells_config.keys()}\n'
                               f'   {output_cells_config.values()}')
                        
            #create empty outputs for each figure
            figure_outputs = {}
            # #will also nned to reinitiate the output cells config inside of loop?           

            #load scenario input file for figure
            try:
                scenarios = pd.read_excel(scenario_file, figure_worksheet)
                main_logger.info(f'Exhibit "{file.stem}": Loaded {len(scenarios)} scenarios from "{scenario_file.stem}" workbook and "{figure_worksheet}" worksheet into dataframe')
            except ValueError as e:
                warning_logger.warning(f'Exhibit "{file.stem}": Worksheet {figure_worksheet} not found')
                continue
            
            # for loop runs through different scenarios and:
            # 1) Sets inputs
            # 2) Solves for number of annual reintroductions required for gains to equal losses
            # 3) Reads desired outputs and writes both inputs and outputs to an output csv for later processing
            # 4) QC check of REA QC tests
            for scenario_number, row in enumerate(scenarios.itertuples(index=False), start =1): 
                scenario_name = scenarios.loc[(scenario_number-1),'scenario_name']

                #Step #1: Set inputs
                # Create scenario inputs class with defaults values
                scenario_inputs = rea_input_class.REAScenarioInputs.create_from_config(CONFIG_PATH, debug=True)
                # and override with input values from the scenarios input dataframe for each scenario
                scenario_inputs.update_from_row(row)
                # print(f'scenario inputs updated: {scenario_inputs}')
                #must convert to dict for easier reading into later functions
                scenario_inputs_dict = scenario_inputs.to_dict()
                main_logger.info(f'Exhibit "{file.stem}", scenario {scenario_number} ({scenario_name}): Inputs loaded')

                xl.set_excel_inputs(io_sheet, scenario_inputs_dict, input_cells_config, scenario_number, main_logger)
                #detail logging of scenario inputs
                log_lines = [f'Exhibit "{file.stem}", scenario {scenario_number} ({scenario_name}): Excel inputs set:']
                #read and return all inputs from config file and scenario inputs override
                for key in input_cells_config.keys():
                    if key in scenario_inputs_dict:
                        clean_key = key.replace('_', ' ').title()
                        log_lines.append(f'  {clean_key} set to: {scenario_inputs_dict[key]}')
                detail_logger.info('\n'.join(log_lines))
                
                #Step #2: Solves for number of annual reintroductions required for gains to equal losses
                # Goal Seek: set Goal:Loss ratio to 1 by changing Annual Mussel Reintroduction 
                xl.run_goal_seek(io_sheet, input_cells_config['loss_ratio'], input_cells_config['annual_reintroduction'], goal_seek_config['target_value'])
                main_logger.info(f'Exhibit "{file.stem}", scenario {scenario_number} ({scenario_name}): Required annual reintroduction calculated (for gain to equal loss)')

                #No such thing as partial mussel so round annual mussel reintroduction down to nearest whole number and set cell to value
                annual_reintroduction_exact = math_util.round_outputs(io_sheet[input_cells_config['annual_reintroduction']].value, decimal_precision_results)
                detail_logger.info(f'Exhibit "{file.stem}", scenario {scenario_number} ({scenario_name}): Exact annual reintroduction: {annual_reintroduction_exact}')
                annual_reintroduction_rounded = math_util.round_annual_reintro(annual_reintroduction_exact)
                detail_logger.info(f'Exhibit "{file.stem}", scenario {scenario_number} ({scenario_name}): Rounded Annual reintroduction: {annual_reintroduction_rounded}')
                io_sheet[input_cells_config['annual_reintroduction']].value =annual_reintroduction_rounded

                #Step #3: Reads desired outputs and writes both inputs and outputs to an output csv for later processing
                #force excel to recalculate
                wb_rea.app.calculate()
                main_logger.info(f'Exhibit "{file.stem}", scenario {scenario_number} ({scenario_name}): Excel workbook recalculated')

                #read model outputs and append to csv file
                outputs = xl.read_excel_outputs(io_sheet, output_cells_config, decimal_precision_results, scenarios, scenario_number, main_logger)
                detail_logger.info(f'Exhibit "{file.stem}", scenario {scenario_number} ({scenario_name}): Read outputs directly from excel: {outputs.keys()}')

                #append annual reintroduction results
                outputs[f'{scenario_name}: Annual Reintroduction Rounded'] = annual_reintroduction_rounded
                outputs[f'{scenario_name}: Annual Reintroduction Exact'] = annual_reintroduction_exact
               
                #append resutls of each scenario to figure_outputs (for exporting) 
                figure_outputs = data_util.append_to_dictionary(figure_outputs, outputs)
                

                output_input_file = output_input_dir / Path(f'{figure_worksheet}.csv')

                csv_data = {'Scenario_number': scenario_number, **scenario_inputs_dict}
                if not output_input_file.exists():
                    csv_util.create_output_csv(output_input_file, csv_data)
                csv_util.append_output_to_csv(output_input_file, list(csv_data.values()))

                console_logger.info(f'{scenario_name} complete')
            
            #export final figure results to excel file
            output_data = {**figure_outputs}
            xl.append_output_excel_file(output_file, output_data, figure_worksheet, console_logger, warning_logger)

            xl.text_wrap_headers(output_file)

            console_logger.info(f'{figure_worksheet} complete')
            
            

            
            # #detail logging of scenario outputs
            # log_lines = [f'Scenario {scenario_number}: Excel outputs written to output file:']
            # #read and return all outputs from config file
            # for key in output_cells_config.keys():
            #     if key in csv_data:
            #         clean_key = key.replace('_', ' ').title()
            #         log_lines.append(f'  {clean_key}: {csv_data[key]}')
            # #inclue outputs created in processing (not in config)
            # log_lines.append(f'  Annual Reintroduction Rounded: {annual_reintroduction_rounded}')
            # log_lines.append(f'  Annual Reintroduction Exact: {annual_reintroduction_exact}')
            # detail_logger.info('\n'.join(log_lines))

            # # Step #4: QC check of REA QC tests
            # #Excel sheet has multiple qc tests whose results are summarized in a single cell as either 'PASS' or 'FAIL'
            # #Check this cell and write I/O to file if 'FAIL' for review
            # xl.check_qc(io_sheet, input_cells_config['qc_test'], output_dir, csv_data, scenario_number, main_logger, warning_logger)
        
            # main_logger.info(f'Scenario {scenario_number}: Scenario completed')
            # console_logger.info(f'{scenario_number}/{len(scenarios)} complete')
            
    finally:
        #close excel instance
        if wb_rea: wb_rea.close()
        if app: app.quit() 
        main_logger.info(f'Closed excel instance')

    #Measure elasped runtime of script
    END_TIME = time.perf_counter()
    RUN_TIME = END_TIME - START_TIME
    RUN_MINUTES = int(RUN_TIME // 60)
    RUN_SECONDS = round((RUN_TIME % 60), 1 )

    console_logger.info(f'Script finished. Total Runtime: {RUN_MINUTES} minutes and {RUN_SECONDS} seconds')