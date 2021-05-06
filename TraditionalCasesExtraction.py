import xlrd
import pandas as pd
import os
import numpy as np

def isnumber(aString):
    try:
        float(aString)
        return True
    except:
        return False

data_traditional = xlrd.open_workbook("data/traditional.xlsx")

table = data_traditional.sheet_by_name('Sheet1')

nrows = table.nrows

total_case = []
read_case = []
total_sum = 0
read_sum = 0

data = np.zeros((62, 3))

for i in range(nrows):
    if isnumber(table.row_values(i)[0]):
        index = str(int(table.row_values(i)[0]))
        total_case.clear()
        read_case.clear()
        cnt = 0
        case = pd.DataFrame(columns=['Round', 'Link', 'Label', 'NoDuplicate'])
    elif table.row_values(i)[0] == "查询词":
        cnt += 1
    elif len(table.row_values(i)[1]) == 0:
        print(index, len(total_case), len(read_case), cnt)
        data[int(index)][0] = len(total_case)
        data[int(index)][1] = len(read_case)
        data[int(index)][2] = cnt
        total_sum += len(total_case)
        read_sum += len(read_case)
    else:
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')]
        if link not in total_case:
            total_case.append(link)
            NoDuplicate = 1
        else:
            NoDuplicate = 0
        if table.row_values(i)[2] == 1 or table.row_values(i)[2] == 2:
            label = int(table.row_values(i)[2])
            if link not in read_case:
                read_case.append(link)
        else:
            label = 0
        # print(cnt, link, label, NoDuplicate)

print(index, len(total_case), len(read_case), cnt)
data[int(index)][0] = len(total_case)
data[int(index)][1] = len(read_case)
data[int(index)][2] = cnt
total_sum += len(total_case)
read_sum += len(read_case)

print(total_sum, read_sum)
data = pd.DataFrame(data)

writer = pd.ExcelWriter('traditional.xlsx')
data.to_excel(writer, 'page_1', float_format='%.5f')
writer.save()

writer.close()