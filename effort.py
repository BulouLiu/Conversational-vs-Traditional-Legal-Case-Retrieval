import xlrd
import pandas as pd
import os
import numpy as np

data = xlrd.open_workbook("data/Multi-level.xlsx")

table = data.sheet_by_name('Sheet1')

cols = table.ncols

def low(target, condition):
    target = np.array(target)
    condition = np.array(condition)
    tmp = condition.copy()
    tmp[condition<3] = 1
    tmp[condition>=3] = 0
    print(np.sum(tmp), np.sum(target*tmp)/np.sum(tmp))

def medium(target, condition):
    target = np.array(target)
    condition = np.array(condition)
    tmp = condition.copy()
    tmp[condition==3] = 1
    tmp[condition!=3] = 0
    print(np.sum(tmp), np.sum(target*tmp)/np.sum(tmp))

def high(target, condition):
    target = np.array(target)
    condition = np.array(condition)
    tmp = condition.copy()
    tmp[condition>3] = 1
    tmp[condition<=3] = 0
    print(np.sum(tmp), np.sum(target*tmp)/np.sum(tmp))

print("Traditional Expertise Cases")
low(table.col_values(4), table.col_values(0))
medium(table.col_values(4), table.col_values(0))
high(table.col_values(4), table.col_values(0))

print("Traditional Expertise Query")
low(table.col_values(6), table.col_values(0))
medium(table.col_values(6), table.col_values(0))
high(table.col_values(6), table.col_values(0))

print("Traditional Expertise Time")
low(table.col_values(8), table.col_values(0))
medium(table.col_values(8), table.col_values(0))
high(table.col_values(8), table.col_values(0))

print("Traditional Expertise Self-reported")
low(table.col_values(10), table.col_values(0))
medium(table.col_values(10), table.col_values(0))
high(table.col_values(10), table.col_values(0))

print("Traditional Expertise Satisfaction")
low(table.col_values(12), table.col_values(0))
medium(table.col_values(12), table.col_values(0))
high(table.col_values(12), table.col_values(0))

print("Traditional Expertise Success")
low(table.col_values(14), table.col_values(0))
medium(table.col_values(14), table.col_values(0))
high(table.col_values(14), table.col_values(0))

print("Traditional Difficulty Cases")
low(table.col_values(4), table.col_values(2))
medium(table.col_values(4), table.col_values(2))
high(table.col_values(4), table.col_values(2))

print("Traditional Difficulty Query")
low(table.col_values(6), table.col_values(2))
medium(table.col_values(6), table.col_values(2))
high(table.col_values(6), table.col_values(2))

print("Traditional Difficulty Time")
low(table.col_values(8), table.col_values(2))
medium(table.col_values(8), table.col_values(2))
high(table.col_values(8), table.col_values(2))

print("Traditional Difficulty Self-reported")
low(table.col_values(10), table.col_values(2))
medium(table.col_values(10), table.col_values(2))
high(table.col_values(10), table.col_values(2))

print("Traditional Difficulty Satisfaction")
low(table.col_values(12), table.col_values(2))
medium(table.col_values(12), table.col_values(2))
high(table.col_values(12), table.col_values(2))

print("Traditional Difficulty Success")
low(table.col_values(14), table.col_values(2))
medium(table.col_values(14), table.col_values(2))
high(table.col_values(14), table.col_values(2))

print("Conversational Expertise Cases")
low(table.col_values(5), table.col_values(1))
medium(table.col_values(5), table.col_values(1))
high(table.col_values(5), table.col_values(1))

print("Conversational Expertise Query")
low(table.col_values(7), table.col_values(1))
medium(table.col_values(7), table.col_values(1))
high(table.col_values(7), table.col_values(1))

print("Conversational Expertise Time")
low(table.col_values(9), table.col_values(1))
medium(table.col_values(9), table.col_values(1))
high(table.col_values(9), table.col_values(1))

print("Conversational Expertise Self-reported")
low(table.col_values(11), table.col_values(1))
medium(table.col_values(11), table.col_values(1))
high(table.col_values(11), table.col_values(1))

print("Conversational Expertise Satisfaction")
low(table.col_values(13), table.col_values(1))
medium(table.col_values(13), table.col_values(1))
high(table.col_values(13), table.col_values(1))

print("Conversational Expertise Success")
low(table.col_values(15), table.col_values(1))
medium(table.col_values(15), table.col_values(1))
high(table.col_values(15), table.col_values(1))

print("Conversational Difficulty Cases")
low(table.col_values(5), table.col_values(3))
medium(table.col_values(5), table.col_values(3))
high(table.col_values(5), table.col_values(3))

print("Conversational Difficulty Query")
low(table.col_values(7), table.col_values(3))
medium(table.col_values(7), table.col_values(3))
high(table.col_values(7), table.col_values(3))

print("Conversational Difficulty Time")
low(table.col_values(9), table.col_values(3))
medium(table.col_values(9), table.col_values(3))
high(table.col_values(9), table.col_values(3))

print("Conversational Difficulty Self-reported")
low(table.col_values(11), table.col_values(3))
medium(table.col_values(11), table.col_values(3))
high(table.col_values(11), table.col_values(3))

print("Conversational Difficulty Satisfaction")
low(table.col_values(13), table.col_values(3))
medium(table.col_values(13), table.col_values(3))
high(table.col_values(13), table.col_values(3))

print("Conversational Difficulty Success")
low(table.col_values(15), table.col_values(3))
medium(table.col_values(15), table.col_values(3))
high(table.col_values(15), table.col_values(3))