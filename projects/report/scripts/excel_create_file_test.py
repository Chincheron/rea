import util.excel_util as excel_util
from util.constants import * 

dict = {'yearly_direct': (5,10,15), 'yearly_indirect': 10}
print(dict)
print(list(dict.keys()) )
print(list(dict.values()))

excel_util.create_output_excel_file(TEST_DIR, dict)