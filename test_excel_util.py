import excel_util
from pathlib import Path

excel_path = Path('test_inputs') / 'Mussel REA_v2.0.1.xlsx'

wb, app = excel_util.load_workbook(excel_path)
print(type(wb))
print(type(app))

sheet = excel_util.load_worksheet(wb, 'Matrix Inputs')
print(type(sheet))
#excel_util.load_worksheet(wb, ['test'])

wb.close()
app.quit()