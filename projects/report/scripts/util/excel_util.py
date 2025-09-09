'''Define functions for manipulating excel files using xlwings'''
import xlwings as xw
from pathlib import Path
import logging
import util.csv_util as csv_util

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

def read_excel_outputs(sheet, output_cells, decimals, logger = None):
    '''Read desired output cells from excel workbook'''
    
    outputs = {}

    for key, cell in output_cells.items():
        #read value of each cell
        value = sheet[cell].value
        outputs[key] = round_cells(value, decimals)
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

