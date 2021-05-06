import pandas as pd
import os
import numpy as np
import sklearn.metrics
import random



def fleiss_kappa(table):
    table = 1.0 * np.asarray(table)  # avoid integer division
    n_sub, n_cat = table.shape
    n_total = table.sum()
    n_rater = table.sum(1)
    n_rat = n_rater.max()
    # assume fully ranked
    # print n_total, n_sub, n_rat
    assert n_total == n_sub * n_rat

    # marginal frequency  of categories
    p_cat = table.sum(0) / n_total

    table2 = table * table
    p_rat = (table2.sum(1) - n_rat) / (n_rat * (n_rat - 1.))
    p_mean = p_rat.mean()

    p_mean_exp = (p_cat * p_cat).sum()

    kappa = (p_mean - p_mean_exp) / (1 - p_mean_exp)
    return kappa


def return_list(annotation_id):
    annotation_l = []
    for i in range(62):
        if i == 7 or i == 26:
            continue
        if os.path.exists("./Annotation"+annotation_id+"/Civil/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Civil/" + str(i) + "/read.xlsx")
            annotation_l += data["Label"].values.tolist()
        elif os.path.exists("./Annotation"+annotation_id+"/Criminal/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Criminal/" + str(i) + "/read.xlsx")
            annotation_l += data["Label"].values.tolist()
        elif os.path.exists("./Annotation"+annotation_id+"/Commercial/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Commercial/" + str(i) + "/read.xlsx")
            annotation_l += data["Label"].values.tolist()
        else:
            continue
    return annotation_l

def return_2list(annotation_id):
    annotation_traditional_l = []
    annotation_conversational_l = []
    for i in range(62):
        if i == 7 or i == 26:
            continue
        if os.path.exists("./Annotation"+annotation_id+"/Civil/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Civil/" + str(i) + "/read.xlsx")
        elif os.path.exists("./Annotation"+annotation_id+"/Criminal/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Criminal/" + str(i) + "/read.xlsx")
        elif os.path.exists("./Annotation"+annotation_id+"/Commercial/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Commercial/" + str(i) + "/read.xlsx")
        else:
            continue
        label = dict(zip(data["File"].values.tolist(), data["Label"].values.tolist()))
        traditional_data = pd.read_excel("./Cases/traditional/" + str(i) + "/read.xlsx")
        conversational_data = pd.read_excel("./Cases/conversational/" + str(i) + "/read.xlsx")

        for item in traditional_data["File"].values.tolist():
            annotation_traditional_l.append(label[item])
        for item in label.keys():
            if item not in traditional_data["File"].values.tolist():
                annotation_conversational_l.append(label[item])


    return annotation_traditional_l, annotation_conversational_l

def total():
    traditional_length = 0
    conversational_length = 0
    annotation_id = "2"
    for i in range(62):
        if i == 7 or i == 26:
            continue
        if os.path.exists("./Annotation"+annotation_id+"/Civil/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Civil/" + str(i) + "/read.xlsx")
        elif os.path.exists("./Annotation"+annotation_id+"/Criminal/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Criminal/" + str(i) + "/read.xlsx")
        elif os.path.exists("./Annotation"+annotation_id+"/Commercial/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Commercial/" + str(i) + "/read.xlsx")
        else:
            continue
        traditional_length += len(pd.read_excel("./Cases/traditional/" + str(i) + "/total.xlsx")["File"].values.tolist())
        conversational_length += len(
            pd.read_excel("./Cases/conversational/" + str(i) + "/total.xlsx")["File"].values.tolist())
    print(traditional_length/55)
    print(conversational_length/55)
    print((traditional_length+conversational_length)/55)

def apply_kappa(res1, res2, res3, length):
    Data = np.zeros((length, 2))
    relevant = 0
    for i in range(length):
        if res1[i] == 2:
            Data[i][1] += 1
        else:
            Data[i][0] += 1
        if res2[i] == 2:
            Data[i][1] += 1
        else:
            Data[i][0] += 1
        if res2[i] == 2:
            Data[i][1] += 1
        else:
            Data[i][0] += 1
        if Data[i][1] >= 2:
            relevant += 1
    print(length/55)
    print(relevant/55)

    print(fleiss_kappa(Data))

def merge_result(res1, res2, res3, length, annotation_id):
    Data = np.zeros((length, 2))
    final_result = []
    for i in range(length):
        if res1[i] == 2:
            Data[i][1] += 1
        else:
            Data[i][0] += 1
        if res2[i] == 2:
            Data[i][1] += 1
        else:
            Data[i][0] += 1
        if res2[i] == 2:
            Data[i][1] += 1
        else:
            Data[i][0] += 1
        if Data[i][1] >= 2:
            final_result.append(2)
        else:
            final_result.append(1)

    cnt = 0
    for i in range(62):
        if i == 7 or i == 26:
            continue
        if os.path.exists("./Annotation"+annotation_id+"/Civil/" + str(i)):
            data = pd.read_excel("./Annotation"+annotation_id+"/Civil/" + str(i) + "/read.xlsx")
            file = data["File"].values.tolist()
            label = []
            for j in range(len(file)):
                label.append(final_result[cnt])
                cnt += 1
            writer = pd.ExcelWriter("./Annotation"+annotation_id+"/Civil/" + str(i) + "/read.xlsx")
            read_case = pd.DataFrame({"File": file, "Label": label})
            read_case.to_excel(writer, 'page_1', float_format='%.5f')
            writer.save()
            writer.close()
        elif os.path.exists("./Annotation"+annotation_id+"/Criminal/" + str(i)):

            data = pd.read_excel("./Annotation"+annotation_id+"/Criminal/" + str(i) + "/read.xlsx")
            file = data["File"].values.tolist()
            label = []
            for j in range(len(file)):
                label.append(final_result[cnt])
                cnt += 1
            writer = pd.ExcelWriter("./Annotation" + annotation_id + "/Criminal/" + str(i) + "/read.xlsx")
            read_case = pd.DataFrame({"File": file, "Label": label})
            read_case.to_excel(writer, 'page_1', float_format='%.5f')
            writer.save()
            writer.close()
        elif os.path.exists("./Annotation"+annotation_id+"/Commercial/" + str(i)):

            data = pd.read_excel("./Annotation"+annotation_id+"/Commercial/" + str(i) + "/read.xlsx")
            file = data["File"].values.tolist()
            label = []
            for j in range(len(file)):
                label.append(final_result[cnt])
                cnt += 1
            writer = pd.ExcelWriter("./Annotation" + annotation_id + "/Commercial/" + str(i) + "/read.xlsx")
            read_case = pd.DataFrame({"File": file, "Label": label})
            read_case.to_excel(writer, 'page_1', float_format='%.5f')
            writer.save()
            writer.close()
        else:
            continue
# total()
#
t1, c1 = return_2list("2")
t2, c2 = return_2list("3")
t3, c3 = return_2list("4")
r1= return_list("2")
r2= return_list("3")
r3= return_list("4")
#
#
apply_kappa(np.array(t1), np.array(t2), np.array(t3), len(t1))
apply_kappa(np.array(c1), np.array(c2), np.array(c3), len(c1))
apply_kappa(np.array(r1), np.array(r2), np.array(r3), len(r1))

merge_result(np.array(r1), np.array(r2), np.array(r3), len(r1), "1")



