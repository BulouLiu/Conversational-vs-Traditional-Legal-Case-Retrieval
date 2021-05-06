import xlrd
import pandas as pd
import os
import numpy as np

data = xlrd.open_workbook("data/Multi-level.xlsx")

table = data.sheet_by_name('Sheet1')

cols = table.ncols

def out_easy(target, expertise, difficulty, expertise_list, difficulty_list, target_list):
    target = np.array(target)
    expertise = np.array(expertise)
    difficulty = np.array(difficulty)

    tmp_exp = expertise.copy()
    tmp_exp[expertise<2.5] = 1
    tmp_exp[expertise>2.5] = 0

    tmp_diff = difficulty.copy()
    tmp_diff[difficulty<2.5] = 1
    tmp_diff[difficulty>2.5] = 0

    tmp = tmp_exp*tmp_diff

    for i in range(len(tmp)):
        if tmp[i] == 1:
            expertise_list.append("Out-domain")
            difficulty_list.append("Easy")
            target_list.append(target[i])

    print(np.sum(tmp), np.sum(target*tmp)/np.sum(tmp))

    return expertise_list, difficulty_list, target_list

def out_hard(target, expertise, difficulty, expertise_list, difficulty_list, target_list):
    target = np.array(target)
    expertise = np.array(expertise)
    difficulty = np.array(difficulty)

    tmp_exp = expertise.copy()
    tmp_exp[expertise < 2.5] = 1
    tmp_exp[expertise > 2.5] = 0

    tmp_diff = difficulty.copy()
    tmp_diff[difficulty > 2.5] = 1
    tmp_diff[difficulty < 2.5] = 0

    tmp = tmp_exp * tmp_diff

    for i in range(len(tmp)):
        if tmp[i] == 1:
            expertise_list.append("Out-domain")
            difficulty_list.append("Hard")
            target_list.append(target[i])

    print(np.sum(tmp), np.sum(target * tmp) / np.sum(tmp))

    return expertise_list, difficulty_list, target_list

def in_easy(target, expertise, difficulty, expertise_list, difficulty_list, target_list):
    target = np.array(target)
    expertise = np.array(expertise)
    difficulty = np.array(difficulty)

    tmp_exp = expertise.copy()
    tmp_exp[expertise > 2.5] = 1
    tmp_exp[expertise < 2.5] = 0

    tmp_diff = difficulty.copy()
    tmp_diff[difficulty < 2.5] = 1
    tmp_diff[difficulty > 2.5] = 0

    tmp = tmp_exp * tmp_diff

    for i in range(len(tmp)):
        if tmp[i] == 1:
            expertise_list.append("In-domain")
            difficulty_list.append("Easy")
            target_list.append(target[i])

    print(np.sum(tmp), np.sum(target * tmp) / np.sum(tmp))

    return expertise_list, difficulty_list, target_list

def in_hard(target, expertise, difficulty, expertise_list, difficulty_list, target_list):
    target = np.array(target)
    expertise = np.array(expertise)
    difficulty = np.array(difficulty)

    tmp_exp = expertise.copy()
    tmp_exp[expertise > 2.5] = 1
    tmp_exp[expertise < 2.5] = 0

    tmp_diff = difficulty.copy()
    tmp_diff[difficulty > 2.5] = 1
    tmp_diff[difficulty < 2.5] = 0

    tmp = tmp_exp * tmp_diff

    for i in range(len(tmp)):
        if tmp[i] == 1:
            expertise_list.append("In-domain")
            difficulty_list.append("Hard")
            target_list.append(target[i])

    print(np.sum(tmp), np.sum(target * tmp) / np.sum(tmp))

    return expertise_list, difficulty_list, target_list

def compute_traditional(target, target_name):
    print(target_name)
    expertise_list = []
    difficulty_list = []
    target_list = []
    expertise_list, difficulty_list, target_list = out_easy(target, table.col_values(0), table.col_values(2), expertise_list, difficulty_list, target_list)
    expertise_list, difficulty_list, target_list = out_hard(target, table.col_values(0), table.col_values(2), expertise_list, difficulty_list, target_list)
    expertise_list, difficulty_list, target_list = in_easy(target, table.col_values(0), table.col_values(2), expertise_list, difficulty_list, target_list)
    expertise_list, difficulty_list, target_list = in_hard(target, table.col_values(0), table.col_values(2), expertise_list, difficulty_list, target_list)

    file_name = "./TwoFactorFiles/traditional_" + target_name + ".csv"
    data = pd.DataFrame({"Expertise": expertise_list, "Difficulty": difficulty_list, target_name: target_list})
    data.to_csv(file_name, index=False)

def compute_conversational(target, target_name):
    print(target_name)
    expertise_list = []
    difficulty_list = []
    target_list = []
    expertise_list, difficulty_list, target_list = out_easy(target, table.col_values(1), table.col_values(3), expertise_list, difficulty_list, target_list)
    expertise_list, difficulty_list, target_list = out_hard(target, table.col_values(1), table.col_values(3), expertise_list, difficulty_list, target_list)
    expertise_list, difficulty_list, target_list = in_easy(target, table.col_values(1), table.col_values(3), expertise_list, difficulty_list, target_list)
    expertise_list, difficulty_list, target_list = in_hard(target, table.col_values(1), table.col_values(3), expertise_list, difficulty_list, target_list)

    file_name = "./TwoFactorFiles/conversational_" + target_name + ".csv"
    data = pd.DataFrame({"Expertise": expertise_list, "Difficulty": difficulty_list, target_name: target_list})
    data.to_csv(file_name, index=False)

if not os.path.exists("./TwoFactorFiles"):
    os.makedirs("./TwoFactorFiles")

compute_traditional(table.col_values(4), "Cases")
compute_conversational(table.col_values(5), "Cases")

compute_traditional(table.col_values(6), "Queries")
compute_conversational(table.col_values(7), "Queries")

compute_traditional(table.col_values(8), "Time")
compute_conversational(table.col_values(9), "Time")

compute_traditional(table.col_values(10), "Self-reported")
compute_conversational(table.col_values(11), "Self-reported")

compute_traditional(table.col_values(12), "Satisfaction")
compute_conversational(table.col_values(13), "Satisfaction")

compute_traditional(table.col_values(14), "Success")
compute_conversational(table.col_values(15), "Success")