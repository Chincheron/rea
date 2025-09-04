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
