'''Define functions for manipulating excel files using xlwings'''
import xlwings as xw
from pathlib import Path
import logging

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


