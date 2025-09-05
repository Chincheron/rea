import excel_util
from pathlib import Path

excel_path = Path('test_inputs') / 'Mussel REA_v2.0.1.xlsx'
test_dir = Path('test_module')
test_dir.mkdir(parents=True, exist_ok=True)

wb, app = excel_util.load_workbook(excel_path)
print(type(wb))
print(type(app))

sheet = excel_util.load_worksheet(wb, 'Matrix Inputs')
# print(type(sheet))
# #excel_util.load_worksheet(wb, ['test'])


# excel_util.set_excel_inputs(sheet, {"number_killed": 500}, {"number_killed": "B2"}, 1)

excel_util.check_qc(sheet, "B5", test_dir, ['header1', 'header2'], [1,2],1)


wb.close()
app.quit()
