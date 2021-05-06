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

source_link = []
target_link = []

f = open("data/ptaloutlines.txt", "r")
for line in f:
    source_link.append(line.split()[0])
    target_link.append(line.split()[1])

s2t = dict(zip(source_link, target_link))

read_case_link = [[]]
total_case_link = [[]]
read_file = [[]]
total_file = [[]]
read_label = [[]]

data = xlrd.open_workbook("data/traditional.xlsx")

table = data.sheet_by_name('Sheet1')

nrows = table.nrows

if not os.path.exists('./Cases'):
    os.makedirs('./Cases')

if not os.path.exists('./Cases/traditional'):
    os.makedirs('./Cases/traditional')

if not os.path.exists('./Cases/conversational'):
    os.makedirs('./Cases/conversational')

count = np.zeros((62, 3))

for i in range(nrows):
    if isnumber(table.row_values(i)[0]):
        index = str(int(table.row_values(i)[0]))
        if not os.path.exists('./Cases/traditional/'+index):
            os.makedirs('./Cases/traditional/'+index)

        cnt = 0
    elif table.row_values(i)[0] == "查询词":
        cnt += 1
        total_case_link.append([])
        total_file.append([])
        read_case_link.append([])
        read_file.append([])
        read_label.append([])
    elif len(table.row_values(i)[1]) == 0:
        writer = pd.ExcelWriter('./Cases/traditional/' + index + '/total.xlsx')
        for i in range(cnt+1):
            total_case = pd.DataFrame({"File": total_file[i]})
            total_case.to_excel(writer, 'page_' + str(i), float_format='%.5f')
        writer.save()
        writer.close()
        writer = pd.ExcelWriter('./Cases/traditional/' + index + '/read.xlsx')
        for i in range(cnt+1):
            read_case = pd.DataFrame({"File": read_file[i], "Label": read_label[i]})
            read_case.to_excel(writer, 'page_' + str(i), float_format='%.5f')
        writer.save()
        writer.close()
        count[int(index)][0] = cnt
        read_case_link = [[]]
        total_case_link = [[]]
        read_file = [[]]
        total_file = [[]]
        read_label = [[]]
    else:
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')].strip()
        if link in source_link:
            link = s2t[link]
        if link not in total_case_link[0]:
            total_case_link[0].append(link)
            total_file[0].append(link.split("/")[-1])
        if table.row_values(i)[2] == 1 or table.row_values(i)[2] == 2:
            if link not in read_case_link[0]:
                read_case_link[0].append(link)
                read_file[0].append(link.split("/")[-1])
                read_label[0].append(table.row_values(i)[2])
        if link not in total_case_link[cnt]:
            total_case_link[cnt].append(link)
            total_file[cnt].append(link.split("/")[-1])
        if table.row_values(i)[2] == 1 or table.row_values(i)[2] == 2:
            if link not in read_case_link[cnt]:
                read_case_link[cnt].append(link)
                read_file[cnt].append(link.split("/")[-1])
                read_label[cnt].append(table.row_values(i)[2])
writer = pd.ExcelWriter('./Cases/traditional/' + index + '/total.xlsx')
for i in range(cnt+1):
    total_case = pd.DataFrame({"File": total_file[i]})
    total_case.to_excel(writer, 'page_' + str(i), float_format='%.5f')
writer.save()
writer.close()
writer = pd.ExcelWriter('./Cases/traditional/' + index + '/read.xlsx')
for i in range(cnt+1):
    read_case = pd.DataFrame({"File": read_file[i], "Label": read_label[i]})
    read_case.to_excel(writer, 'page_' + str(i), float_format='%.5f')
writer.save()
writer.close()
count[int(index)][0] = cnt
read_case_link = [[]]
total_case_link = [[]]
read_file = [[]]
total_file = [[]]
read_label = [[]]

data_conversational = xlrd.open_workbook("data/conversational.xlsx")

table = data_conversational.sheet_by_name('Sheet1')

nrows = table.nrows

for i in range(nrows):
    if isnumber(table.row_values(i)[0]):
        index = str(int(table.row_values(i)[0]))
        if not os.path.exists('./Cases/conversational/'+index):
            os.makedirs('./Cases/conversational/'+index)
        cnt = 0
        reform = 0
        question = 0
    elif table.row_values(i)[0] == "查询词":
        cnt += 1
        state = 1
        total_case_link.append([])
        total_file.append([])
    elif table.row_values(i)[0] == "返回案件":
        state = 2
        reform += 1
        read_case_link.append([])
        read_file.append([])
        read_label.append([])
        for k in range(len(read_label)):
            for j in range(len(read_label[k])):
                read_label[k][j] = 1
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')].strip()
        if link in source_link:
            link = s2t[link]
        if link not in read_case_link[0]:
            read_case_link[0].append(link)
            read_file[0].append(link.split("/")[-1])
            read_label[0].append(2)
        if link not in read_case_link[reform]:
            read_case_link[reform].append(link)
            read_file[reform].append(link.split("/")[-1])
            read_label[reform].append(2)
    elif len(table.row_values(i)[1]) == 0:
        writer = pd.ExcelWriter('./Cases/conversational/' + index + '/total.xlsx')
        for i in range(cnt + 1):
            total_case = pd.DataFrame({"File": total_file[i]})
            total_case.to_excel(writer, 'page_' + str(i), float_format='%.5f')
        writer.save()
        writer.close()
        writer = pd.ExcelWriter('./Cases/conversational/' + index + '/read.xlsx')
        for i in range(reform + 1):
            read_case = pd.DataFrame({"File": read_file[i], "Label": read_label[i]})
            read_case.to_excel(writer, 'page_' + str(i), float_format='%.5f')
        writer.save()
        writer.close()
        count[int(index)][1] = cnt
        count[int(index)][2] = reform
        read_case_link = [[]]
        total_case_link = [[]]
        read_file = [[]]
        total_file = [[]]
        read_label = [[]]
    elif table.row_values(i)[0] == "专家" or table.row_values(i)[0] == "用户":
        if table.row_values(i)[0] == "专家":
            question += 1
    elif state == 1:
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')].strip()
        if link in source_link:
            link = s2t[link]
        if link not in total_case_link[0]:
            total_case_link[0].append(link)
            total_file[0].append(link.split("/")[-1])
        if link not in total_case_link[cnt]:
            total_case_link[cnt].append(link)
            total_file[cnt].append(link.split("/")[-1])
    elif state == 2:
        link = table.row_values(i)[1][0:table.row_values(i)[1].find('?')].strip()
        if link in source_link:
            link = s2t[link]
        if link not in read_case_link[0]:
            read_case_link[0].append(link)
            read_file[0].append(link.split("/")[-1])
            read_label[0].append(2)
        if link not in read_case_link[reform]:
            read_case_link[reform].append(link)
            read_file[reform].append(link.split("/")[-1])
            read_label[reform].append(2)

writer = pd.ExcelWriter('./Cases/conversational/' + index + '/total.xlsx')
for i in range(cnt + 1):
    total_case = pd.DataFrame({"File": total_file[i]})
    total_case.to_excel(writer, 'page_' + str(i), float_format='%.5f')
writer.save()
writer.close()
writer = pd.ExcelWriter('./Cases/conversational/' + index + '/read.xlsx')
for i in range(reform + 1):
    read_case = pd.DataFrame({"File": read_file[i], "Label": read_label[i]})
    read_case.to_excel(writer, 'page_' + str(i), float_format='%.5f')
writer.save()
writer.close()
count[int(index)][1] = cnt
count[int(index)][2] = reform

count= pd.DataFrame(count)

writer = pd.ExcelWriter('count.xlsx')
count.to_excel(writer, 'page_1', float_format='%.5f')
writer.save()

writer.close()