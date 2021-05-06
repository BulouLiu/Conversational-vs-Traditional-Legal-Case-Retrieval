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

data_conversational = xlrd.open_workbook("data/conversational.xlsx")

table = data_conversational.sheet_by_name('Sheet1')

nrows = table.nrows

total_case = []
read_case = []
return_case = []
total_sum = 0
read_sum = 0
return_sum = 0
state = 1

data = np.zeros((62,5))

for i in range(nrows):
    if isnumber(table.row_values(i)[0]):
        index = str(int(table.row_values(i)[0]))
        total_case.clear()
        read_case.clear()
        return_case.clear()
        cnt = 0
        reform = 0
        question = 0
        case = pd.DataFrame(columns=['Round', 'Link', 'Label', 'NoDuplicate'])
    elif table.row_values(i)[0] == "查询词":
        cnt += 1
        state = 1
    elif table.row_values(i)[0] == "返回案件":
        cnt = 0
        reform += 1
        state = 2
    elif len(table.row_values(i)[1]) == 0:
        print(index, len(total_case), len(read_case))
        data[int(index)][0] = len(total_case)
        data[int(index)][1] = len(read_case)
        data[int(index)][2] = reform
        data[int(index)][3] = len(return_case)
        data[int(index)][4] = question - 1 - reform
        total_sum += len(total_case)
        read_sum += len(read_case)
        return_sum += len(return_case)
    elif table.row_values(i)[0] == "专家" or table.row_values(i)[0] == "用户":
        if table.row_values(i)[0] == "专家":
            question += 1
    elif state == 1:
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
        if table.row_values(i)[2] == 2:
            if link not in return_case:
                return_case.append(link)
        # print(cnt, link, label, NoDuplicate)

print(index, len(total_case), len(read_case))
data[int(index)][0] = len(total_case)
data[int(index)][1] = len(read_case)
data[int(index)][2] = reform
data[int(index)][3] = len(return_case)
data[int(index)][4] = question - 1 - reform
total_sum += len(total_case)
read_sum += len(read_case)
return_sum += len(return_case)

data = pd.DataFrame(data)

writer = pd.ExcelWriter('conversational.xlsx')
data.to_excel(writer, 'page_1', float_format='%.5f')
writer.save()

writer.close()
print(total_sum, read_sum)