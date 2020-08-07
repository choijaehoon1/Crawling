import xlrd
from collections import OrderedDict
import json

excel_file_path = 'D:\\fintech\\financialAPI\\stockpricesDB.xlsx'

wb = xlrd.open_workbook(excel_file_path)
sh = wb.sheet_by_index(0)

data_list = []

for rownum in range(1, sh.nrows):
    data = OrderedDict()
    row_values = sh.row_values(rownum)
    data['rank'] = row_values[1]
    data['issue'] = row_values[2]
    data['search_ratio'] = row_values[3]
    data['present_price'] = row_values[4]
    data['diff'] = row_values[5]
    data['diff_ratio'] = row_values[6]
    data['volume'] = row_values[7]
    data['open_price'] = row_values[8]
    data['high_price'] = row_values[9]
    data['low_price'] = row_values[10]
    data['per'] = row_values[11]
    data['roe'] = row_values[12]
    data_list.append(data)

j = json.dumps(data_list, ensure_ascii=False)

with open('stockprices.json','w+') as f:
    f.write(j)


