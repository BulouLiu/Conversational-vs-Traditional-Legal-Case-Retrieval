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

def make_hyperlink(value):
    return '=HYPERLINK("%s", "%s")' % (value, value)

def make_locallink(value):
    local_link = "./cases/{}.txt"
    return '=HYPERLINK("%s", "%s")' % (local_link.format(value), value)

source_link = []
target_link = []

f = open("data/ptaloutlines.txt", "r")
for line in f:
    source_link.append(line.split()[0])
    target_link.append(line.split()[1])

s2t = dict(zip(source_link, target_link))

read_case_link = []
total_case_link = []
read_file = []
total_file = []

for i in range(62):
    total_case_link.append([])
    read_case_link.append([])
    total_file.append([])
    read_file.append([])

if not os.path.exists('./Annotation'):
    os.makedirs('./Annotation')

data = xlrd.open_workbook("data/traditional.xlsx")

table = data.sheet_by_name('Sheet1')

nrows = table.nrows

for i in range(nrows):
    if isnumber(table.row_values(i)[0]):
        index = str(int(table.row_values(i)[0]))
        cnt = 0
    elif table.row_values(i)[0] == "查询词":
        cnt += 1
    elif len(table.row_values(i)[1]) == 0:
        empty = ''
    else:
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')].strip()
        if link in source_link:
            link = s2t[link]
        if link not in total_case_link[int(index)]:
            total_case_link[int(index)].append(link)
            total_file[int(index)].append(link.split("/")[-1])
        if table.row_values(i)[2] == 1 or table.row_values(i)[2] == 2:
            if link not in read_case_link[int(index)]:
                read_case_link[int(index)].append(link)
                read_file[int(index)].append(link.split("/")[-1])

data_conversational = xlrd.open_workbook("data/conversational.xlsx")

table = data_conversational.sheet_by_name('Sheet1')

nrows = table.nrows

total_sum = 0
read_sum = 0
data = np.zeros((62, 2))

for i in range(nrows):
    if isnumber(table.row_values(i)[0]):
        index = str(int(table.row_values(i)[0]))
    elif table.row_values(i)[0] == "查询词":
        cnt += 1
        state = 1
    elif table.row_values(i)[0] == "返回案件":
        cnt = 0
        state = 2
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')].strip()
        if link in source_link:
            link = s2t[link]
        if link not in read_case_link[int(index)]:
            print("error")
            print(link)
    elif len(table.row_values(i)[1]) == 0:
        total_sum += len(total_case_link[int(index)])
        read_sum += len(read_case_link[int(index)])
        data[int(index)][0] = len(total_case_link[int(index)])
        data[int(index)][1] = len(read_case_link[int(index)])
        if not os.path.exists('./Annotation/'+index):
            os.makedirs('./Annotation/'+index)
        if not os.path.exists('./Annotation/'+index+'/cases'):
            os.makedirs('./Annotation/'+index+'/cases')
        total_case = pd.DataFrame({"Link": total_case_link[int(index)], "File": total_file[int(index)]})
        read_case = pd.DataFrame({"File": read_file[int(index)], "Label": ""})
        total_case["Link"] = total_case["Link"].apply(lambda x: make_hyperlink(x))
        writer = pd.ExcelWriter('./Annotation/'+index+'/total.xlsx')
        total_case.to_excel(writer, 'page_1', float_format='%.5f')
        writer.save()
        writer.close()
        writer = pd.ExcelWriter('./Annotation/' + index + '/read.xlsx')
        read_case.to_excel(writer, 'page_1', float_format='%.5f')
        writer.save()
        writer.close()

    elif table.row_values(i)[0] == "专家" or table.row_values(i)[0] == "用户":
        name = table.row_values(i)[0]
    elif state == 1:
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')].strip()
        if link in source_link:
            link = s2t[link]
        if link not in total_case_link[int(index)]:
            total_case_link[int(index)].append(link)
            total_file[int(index)].append(link.split("/")[-1])
        if table.row_values(i)[2] == 1 or table.row_values(i)[2] == 2:
            if link not in read_case_link[int(index)]:
                read_case_link[int(index)].append(link)
                read_file[int(index)].append(link.split("/")[-1])
    elif state == 2:
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')].strip()
        if link in source_link:
            link = s2t[link]
        if link not in read_case_link[int(index)]:
            print("error")
            print(link)

total_sum += len(total_case_link[int(index)])
read_sum += len(read_case_link[int(index)])
if not os.path.exists('./Annotation/'+index):
    os.makedirs('./Annotation/'+index)
if not os.path.exists('./Annotation/' + index + '/cases'):
    os.makedirs('./Annotation/' + index + '/cases')
total_case = pd.DataFrame({"Link": total_case_link[int(index)], "File": total_file[int(index)]})
read_case = pd.DataFrame({"File": read_file[int(index)], "Label": ""})
total_case["Link"] = total_case["Link"].apply(lambda x: make_hyperlink(x))
writer = pd.ExcelWriter('./Annotation/'+index+'/total.xlsx')
total_case.to_excel(writer, 'page_1', float_format='%.5f')
writer.save()
writer.close()
writer = pd.ExcelWriter('./Annotation/' + index + '/read.xlsx')
read_case.to_excel(writer, 'page_1', float_format='%.5f')
writer.save()
writer.close()

for i in range(1, 62):
    for item in read_case_link[i]:
        if item not in total_case_link[i]:
            print(i ,item)

# total = []

# for i in range(62):
#     if os.path.exists('./Annotation/'+str(i)):
#         for j in range(len(total_case_link[i])):
#             if total_case_link[i][j] not in total:
#                 if len(total_case_link[i][j]) == 0:
#                     print(i, total_case_link[i][j])
#                 elif total_case_link[i][j][0]!='h':
#                     print(i, total_case_link[i][j])
#                 total.append(total_case_link[i][j])
#
# all_case = pd.DataFrame({"Link": total})
# writer = pd.ExcelWriter('candidates.xlsx')
# all_case.to_excel(writer, 'page_1', float_format='%.5f')
# writer.save()
# writer.close()

data[int(index)][0] = len(total_case_link[int(index)])
data[int(index)][1] = len(read_case_link[int(index)])
data = pd.DataFrame(data)
writer = pd.ExcelWriter('case_number.xlsx')
data.to_excel(writer, 'page_1', float_format='%.5f')
writer.save()
writer.close()

print(total_sum, read_sum)