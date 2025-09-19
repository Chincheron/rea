'''Define functions for manipulating excel files using xlwings'''
import xlwings as xw
from pathlib import Path
import logging
import util.csv_util as csv_util
import pandas as pd

def load_workbook(workbook: Path |str, visible: bool = False) -> tuple[xw.Book, xw.App]:
    '''Loads Excel file using xlwing in background (headless) mode'''

    workbook = Path(workbook)
    
    try:
        app = xw.App(visible = visible)
        app.display_alerts = visible
        app.screen_updating = visible
        wb = app.books.open(workbook)
        return wb, app
    except Exception as e:
        app.quit()
        raise e
    
def load_worksheet(wb: xw.Book, sheet_name: str, logger: logging.Logger | None = None) -> xw.Sheet:
    '''Loads specificed worksheet from xlwings workbook'''

    if logger is None:
        logger = logging.getLogger(__name__)

    try:
        return wb.sheets[sheet_name]
    except Exception as e:
        msg = f'Sheet {sheet_name} not found in workbook {wb}'
        logger.error(msg, exc_info=e)
        raise ValueError(msg) from e
    
def run_goal_seek(sheet, goal_cell, changing_cell, target_value):
    '''Run Excel's goal seek for specified paramters'''

    goal_cell = sheet.range(goal_cell).api
    changing_cell = sheet.range(changing_cell).api
    goal_cell.GoalSeek(Goal=target_value, ChangingCell=changing_cell)

def set_excel_inputs(sheet, input_values, input_cells, scenario_number, logger = None):
        '''Set input cells to specified values'''

        if logger is None:
            logger = logging.getLogger(__name__)

        for key, value in input_values.items():
            sheet[input_cells[key]].value = value

        logger.info(f'Scenario {scenario_number}: Excel cells set to scenario inputs')

def round_cells(value, decimals):
    '''
    Check whether cell is a number and round to specified decimals if so
    If cells are a range (1D or 2D) (i.e., a list), recursively calls itself until a single cell is returned
    Return value of cell if not a number
    '''
    #case of single cell with number
    if isinstance(value, (int, float)):
        return round(value, decimals)
    #case of 1D or 2D range of cells
    elif isinstance(value, list):
        return [round_cells(cell, decimals) for cell in value] #works through provided range until a single cell is provided
    else:
        return value

def read_excel_outputs(sheet, output_cells, decimals, scenarios = None, scenario_number = None,  logger = None):
    '''Read desired output cells from excel workbook
    Can optionally pass the current scenario row to append scenario name to headers (for yearly outputs primarily)
    '''
    print(output_cells.items())
    if scenarios is None:
        outputs = {}
        for key, cell in output_cells.items():
            #read value of each cell
            value = sheet[cell].value
            outputs[f'{key}'] = round_cells(value, decimals)
        return outputs
        
    else:
        scenario_name = scenarios.loc[(scenario_number-1),'scenario_name']
        outputs = {}
        outputs[scenario_name] = '' #ensures an empty column between yearly scenarios for readability

        for key, cell in output_cells.items():
            #read value of each cell
            value = sheet[cell].value
            outputs[f'{scenario_name}:{key}'] = round_cells(value, decimals)
        return outputs

def check_qc(sheet, qc_cell, output_dir, csv_data, scenario_number, main_logger = None, warning_logger = None):
    '''Checks value of QC cell in Excel and write input/outputs to file if not "PASS"'''
    
    #setup loggers
    if main_logger is None:
        main_logger = logging.getLogger(__name__)
    if warning_logger is None:
        warning_logger = logging.getLogger(__name__)

    #setup variables
    output_file = output_dir / 'failed_scenario.csv'
    qc_test = sheet[qc_cell].value

    main_logger.info(f'Checking whether Excel workbook QC tests pass')
    if qc_test == 'PASS':
        main_logger.info(f'Scenario {scenario_number}: QC test passed: {qc_test}')
    else: 
        main_logger.warning(f'Scenario {scenario_number}: QC test failed')
        warning_logger.warning(f'Scenario {scenario_number}: QC test failed')
        if not output_file.exists():
            csv_util.create_output_csv(output_file, csv_data, warning_logger)
        
        warning_logger.warning(f'Scenario {scenario_number}: Created failed scenario output file')
        csv_util.append_output_to_csv(output_file, csv_data.values())
        warning_logger.warning(f'Scenario {scenario_number}: Failed scenario inputs/outputs written to failed scenario outputs file ')

def create_output_excel_file(path: Path | str, headers: list | dict, sheet_name = None, main_logger: logging.Logger | None = None, warning_logger: logging.Logger | None = None):
    '''Create excel file and populate with specified headers'''

    #setup loggers
    if main_logger is None:
        main_logger = logging.getLogger(__name__)
    if warning_logger is None:
        warning_logger = logging.getLogger(__name__)    
    
    try:
            
        if isinstance(headers, dict):
            headers = list(headers.keys()) 
            
        path = Path(path) #/ 'teset_excel_create'

        wb = xw.Book()
        main_logger.info(f'Workbook {path} created')

        # new_sheet = wb.sheets.add(name= 'teset') 

        #write headers ()
        # new_sheet.range('A1').value = headers
        # new_sheet.range('a2').options(transpose=True).value = ('Hello', 'test')

        wb.save(path)
        # with open(path, 'w', newline='') as file:
        #     writer = csv.writer(file)
        #     writer.writerow(headers)
        #     logger.info(f'Ouput file created')
    finally:
        if wb: wb.close()
        
def append_output_excel_file(path: Path | str, output_dic: dict, sheet_name = None, main_logger: logging.Logger | None = None, warning_logger: logging.Logger | None = None):
    '''Append outputs to excel file and populate with specified headers'''
    
    #setup loggers
    if main_logger is None:
        main_logger = logging.getLogger(__name__)
    if warning_logger is None:
        warning_logger = logging.getLogger(__name__)

    path = Path(path)

    if not path.exists():
        main_logger.info(f'{path} does not exist. Creating...')
        create_output_excel_file(path, output_dic, sheet_name, main_logger, warning_logger)
    
    
    try:
            
        if isinstance(output_dic, dict):
            headers = list(output_dic.keys()) 
            
        path = Path(path) 

        wb = xw.Book(path)


        #create worksheet
        new_sheet = wb.sheets.add(name= sheet_name) 
        main_logger.info(f'Worksheet {sheet_name} created')

        #write headers ()
        new_sheet.range('A1').value = headers
        
        #tracking columns for loop below
        col_num = 1
        row_num = 2

        #write data
        for key in output_dic:
            col_output = output_dic[key]
            start_cell = new_sheet.cells(row_num, col_num)
            start_cell.options(transpose=True).value = col_output
            col_num += 1
            # new_sheet.range('A2').options(transpose=True).value = col_output
            # new_sheet.range('a2').options(transpose=True).value = (output_dic.values())

        wb.save(path)
        # with open(path, 'w', newline='') as file:
        #     writer = csv.writer(file)
        #     writer.writerow(headers)
        #     logger.info(f'Ouput file created')
    finally:
        if wb: wb.close()
    
def text_wrap_headers(path: Path | str):
    '''Formats all worksheets in provided workbook so that headers are text wrapped'''
    try:
        path = Path(path)
        wb = xw.Book(path)

        for sheet in wb.sheets:
            row = sheet.range('1:1').expand('right')
            # row.select()
            row.column_width = 10
            row.api.WrapText = True
        
        wb.save(path)
    finally:
        if wb: wb.close()

    # open
    # loop through worksheets
    #     text wrap top line