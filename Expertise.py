import xlrd
import pandas as pd
import os
import numpy as np
import scipy.stats as stats

data = xlrd.open_workbook("data/Multi-level.xlsx")

table = data.sheet_by_name('Sheet1')

cols = table.ncols

def low(target, condition):
    target = np.array(target)
    condition = np.array(condition)
    tmp = condition.copy()
    tmp[condition<3] = 1
    tmp[condition>=3] = 0
    # print(np.sum(tmp), np.sum(target*tmp)/np.sum(tmp))
    condition_l = []
    for i in range(len(tmp)):
        if tmp[i] == 1:
            condition_l.append(target[i])
    return condition_l

def medium(target, condition):
    target = np.array(target)
    condition = np.array(condition)
    tmp = condition.copy()
    tmp[condition==3] = 1
    tmp[condition!=3] = 0
    # print(np.sum(tmp), np.sum(target*tmp)/np.sum(tmp))
    condition_l = []
    for i in range(len(tmp)):
        if tmp[i] == 1:
            condition_l.append(target[i])
    return condition_l

def high(target, condition):
    target = np.array(target)
    condition = np.array(condition)
    tmp = condition.copy()
    tmp[condition>3] = 1
    tmp[condition<=3] = 0
    # print(np.sum(tmp), np.sum(target*tmp)/np.sum(tmp))
    condition_l = []
    for i in range(len(tmp)):
        if tmp[i] == 1:
            condition_l.append(target[i])
    return condition_l

names = ["Cases", "Queries", "Time","case time", "Self-reported", "Satisfaction", "subject-success", "object-success",
         "all-P", "read-P", "last-P", "all-RR", "read-RR", "last-RR", "P", "R", "F1"]
print("General")
for i in range(17):
    print(names[i])
    print(np.mean(table.col_values(2*i+4)), np.mean(table.col_values(2*i+5)))
    print(stats.mannwhitneyu(table.col_values(2*i+4),table.col_values(2*i+5),alternative='two-sided')[1])



# if not os.path.exists("./Expertise"):
#     os.makedirs("./Expertise")
#
print("Expertise")
for i in range(17):
    print(names[i])
    A1 = low(table.col_values(2*i+4), table.col_values(0))
    B1 = medium(table.col_values(2*i+4), table.col_values(0))
    C1 = high(table.col_values(2*i+4), table.col_values(0))
    A2 = low(table.col_values(2*i+5), table.col_values(1))
    B2 = medium(table.col_values(2*i+5), table.col_values(1))
    C2 = high(table.col_values(2*i+5), table.col_values(1))
    # print(stats.mannwhitneyu(A1, A2, alternative='two-sided')[1])
    # print(stats.mannwhitneyu(B1, B2, alternative='two-sided')[1])
    # print(stats.mannwhitneyu(C1, C2, alternative='two-sided')[1])
    print(np.mean(np.array(A1+B1)), np.mean(np.array(A2+B2)))
    print(np.mean(np.array(C1)), np.mean(np.array(C2)))
    print(stats.mannwhitneyu(A1+B1, A2+B2, alternative='two-sided')[1])
    print(stats.mannwhitneyu(C1, C2, alternative='two-sided')[1])
#     Method_list = []
#     Expertise_list = []
#     target_list = []
#     for num in A1+B1:
#         Method_list.append("traditional")
#         Expertise_list.append("Out-domain")
#         target_list.append(num)
#     for num in A2+B2:
#         Method_list.append("conversational")
#         Expertise_list.append("Out-domain")
#         target_list.append(num)
#     for num in C1:
#         Method_list.append("traditional")
#         Expertise_list.append("In-domain")
#         target_list.append(num)
#     for num in C2:
#         Method_list.append("conversational")
#         Expertise_list.append("In-domain")
#         target_list.append(num)
#     data = pd.DataFrame({"Method": Method_list, "Expertise": Expertise_list, names[i]: target_list})
#     data.to_csv("./Expertise/" + names[i] + ".csv", index=False)


# print("Cases")
# A1 = low(table.col_values(4), table.col_values(0))
# B1 = medium(table.col_values(4), table.col_values(0))
# C1 = high(table.col_values(4), table.col_values(0))
# A2 = low(table.col_values(5), table.col_values(1))
# B2 = medium(table.col_values(5), table.col_values(1))
# C2 = high(table.col_values(5), table.col_values(1))
#
# print(stats.mannwhitneyu(A1,C1,alternative='two-sided')[1])
# print(stats.kruskal(A1, B1, C1)[1])
#
# print("Traditional Expertise Query")
# A = low(table.col_values(6), table.col_values(0))
# B = medium(table.col_values(6), table.col_values(0))
# C = high(table.col_values(6), table.col_values(0))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Traditional Expertise Time")
# A = low(table.col_values(8), table.col_values(0))
# B = medium(table.col_values(8), table.col_values(0))
# C = high(table.col_values(8), table.col_values(0))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Traditional Expertise Self-reported")
# A = low(table.col_values(10), table.col_values(0))
# B = medium(table.col_values(10), table.col_values(0))
# C = high(table.col_values(10), table.col_values(0))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Traditional Expertise Satisfaction")
# A = low(table.col_values(12), table.col_values(0))
# B = medium(table.col_values(12), table.col_values(0))
# C = high(table.col_values(12), table.col_values(0))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Traditional Expertise Success")
# A = low(table.col_values(14), table.col_values(0))
# B = medium(table.col_values(14), table.col_values(0))
# C = high(table.col_values(14), table.col_values(0))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Conversational Expertise Cases")
# A = low(table.col_values(5), table.col_values(1))
# B = medium(table.col_values(5), table.col_values(1))
# C = high(table.col_values(5), table.col_values(1))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Conversational Expertise Query")
# A = low(table.col_values(7), table.col_values(1))
# B = medium(table.col_values(7), table.col_values(1))
# C = high(table.col_values(7), table.col_values(1))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Conversational Expertise Time")
# A = low(table.col_values(9), table.col_values(1))
# B = medium(table.col_values(9), table.col_values(1))
# C = high(table.col_values(9), table.col_values(1))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Conversational Expertise Self-reported")
# A = low(table.col_values(11), table.col_values(1))
# B = medium(table.col_values(11), table.col_values(1))
# C = high(table.col_values(11), table.col_values(1))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Conversational Expertise Satisfaction")
# A = low(table.col_values(13), table.col_values(1))
# B = medium(table.col_values(13), table.col_values(1))
# C = high(table.col_values(13), table.col_values(1))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
#
# print("Conversational Expertise Success")
# A = low(table.col_values(15), table.col_values(1))
# B = medium(table.col_values(15), table.col_values(1))
# C = high(table.col_values(15), table.col_values(1))
# # print(stats.mannwhitneyu(A+B,C,alternative='two-sided'))
# # print(stats.mannwhitneyu(A,B+C,alternative='two-sided'))
# print(stats.mannwhitneyu(A,C,alternative='two-sided'))
# print(stats.kruskal(A, B, C))
