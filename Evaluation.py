import pandas as pd
import os
import numpy as np
import scipy.stats as stats


count = pd.read_excel('count.xlsx', "page_1")
count = np.array(count)

def Precision(score):
    if len(score) > 0:
        return score.count(2)/len(score)
    return 0

def RR(score):
    if len(score) > 0:
        for i in range(len(score)):
            if score[i] == 2:
                return 1/(i+1)
    return 0

def list_Precision(score_list):
    print(len(score_list)-1)
    res = 0
    for i, score in enumerate(score_list):
        if i == 0:
            continue
        else:
            res += Precision(score)
    return res / len(score_list)

def list_RR(score_list):
    res = 0
    for i, score in enumerate(score_list):
        if i == 0:
            continue
        else:
            res += RR(score)
    return res / len(score_list)

def lastK_Precision(score_list, k):
    res = 0
    for i in range(len(score_list)):
        if i >= k:
            break
        res += Precision(score_list[len(score_list)-1-i])
    return res / min(len(score_list), k)

def lastK_RR(score_list, k):
    res = 0
    for i in range(len(score_list)):
        if i >= k:
            break
        res += RR(score_list[len(score_list) - 1 - i])
    return res / min(len(score_list), k)

def success(self_score, real_score):
    for i in range(len(self_score)):
        if self_score[i] == 2 and real_score[i] == 2:
            return 1
    return 0

def P_R_F1(self_score, real_score):
    TP = 0
    FN = 0
    FP = 0
    for i in range(len(self_score)):
        if self_score[i] == 2 and real_score[i] == 2:
            TP += 1
        elif self_score[i] == 2 and real_score[i] == 1:
            FP += 1
        elif self_score[i] == 1 and real_score[i] == 2:
            FN += 1
    if TP+FP == 0:
        P = 0
    else:
        P = TP / (TP+FP)
    if TP+FN == 0:
        R = 0
    else:
        R = TP / (TP+FN)

    if P+R == 0:
        F1 = 0
    else:
        F1 = 2*P*R/(P+R)
    return P, R, F1

def np2list(data):
    empty = [0, 15, 26, 30, 50, 54]
    res = []
    for i in range(len(data)):
        if i not in empty:
            res.append(data[i])
    return res

annotation_id = "1"
lastK = 1

MP = np.zeros((62, 6))
MRR = np.zeros((62, 6))
SUC = np.zeros((62, 2))
F = np.zeros((62, 6))

for i in range(62):
    if os.path.exists("./Annotation"+annotation_id+"/Civil/" + str(i)):
        data = pd.read_excel("./Annotation"+annotation_id+"/Civil/" + str(i) + "/read.xlsx")
    elif os.path.exists("./Annotation"+annotation_id+"/Criminal/" + str(i)):
        data = pd.read_excel("./Annotation"+annotation_id+"/Criminal/" + str(i) + "/read.xlsx")
    elif os.path.exists("./Annotation"+annotation_id+"/Commercial/" + str(i)):
        data = pd.read_excel("./Annotation"+annotation_id+"/Commercial/" + str(i) + "/read.xlsx")
    else:
        continue

    print(i, end=' ')
    label = dict(zip(data["File"].values.tolist(), data["Label"].values.tolist()))

    traditional_sheets = []
    conversational_total_sheets = []
    conversational_read_sheets = []

    for j in range(count[i][1]+1):
        traditional_sheets.append("page_"+str(j))
    for j in range(count[i][2]+1):
        conversational_total_sheets.append("page_"+str(j))
    for j in range(count[i][3]+1):
        conversational_read_sheets.append("page_"+str(j))

    traditional_total_list = []
    traditional_read_score = []
    conversational_total_list = []
    conversational_read_score = []

    for j, sheet in enumerate(traditional_sheets):
        traditional_query_id = pd.read_excel('./Cases/traditional/' + str(i) + '/total.xlsx', sheet)["File"].values.tolist()
        traditional_total_list.append([])
        for id in traditional_query_id:
            if id in label.keys():
                traditional_total_list[j].append(label[id])
            else:
                traditional_total_list[j].append(0)

    traditional_session_id = pd.read_excel('./Cases/traditional/' + str(i) + '/read.xlsx', "page_0")["File"].values.tolist()
    traditional_read_self_score = pd.read_excel('./Cases/traditional/' + str(i) + '/read.xlsx', "page_0")["Label"].values.tolist()
    for id in traditional_session_id:
        if id in label.keys():
            traditional_read_score.append(label[id])
        else:
            traditional_read_score.append(0)

    for j, sheet in enumerate(conversational_total_sheets):
        conversational_query_id = pd.read_excel('./Cases/conversational/' + str(i) + '/total.xlsx', sheet)[
            "File"].values.tolist()
        conversational_total_list.append([])
        for id in conversational_query_id:
            if id in label.keys():
                conversational_total_list[j].append(label[id])
            else:
                conversational_total_list[j].append(0)

    conversational_session_id = pd.read_excel('./Cases/conversational/' + str(i) + '/read.xlsx', "page_0")[
        "File"].values.tolist()
    conversational_read_self_score = pd.read_excel('./Cases/conversational/' + str(i) + '/read.xlsx', "page_0")[
        "Label"].values.tolist()
    for id in conversational_session_id:
        if id in label.keys():
            conversational_read_score.append(label[id])
        else:
            conversational_read_score.append(0)

    # MP[i][0] = list_Precision(traditional_total_list)
    MP[i][1] = list_Precision(conversational_total_list)
    MP[i][2] = Precision(traditional_read_score)
    MP[i][3] = Precision(conversational_read_score)
    MP[i][4] = lastK_Precision(traditional_total_list, lastK)
    MP[i][5] = lastK_Precision(conversational_total_list, lastK)

    MRR[i][0] = list_RR(traditional_total_list)
    MRR[i][1] = list_RR(conversational_total_list)
    MRR[i][2] = RR(traditional_read_score)
    MRR[i][3] = RR(conversational_read_score)
    MRR[i][4] = lastK_RR(traditional_total_list, lastK)
    MRR[i][5] = lastK_RR(conversational_total_list, lastK)

    SUC[i][0] = success(traditional_read_self_score, traditional_read_score)
    SUC[i][1] = success(conversational_read_self_score, conversational_read_score)

    F[i][0], F[i][2], F[i][4] = P_R_F1(traditional_read_self_score, traditional_read_score)
    F[i][1], F[i][3], F[i][5] = P_R_F1(conversational_read_self_score, conversational_read_score)

# print(stats.mannwhitneyu(np2list(MP[:, 0]), np2list(MP[:, 1]), alternative='two-sided')[1])
# print(stats.mannwhitneyu(np2list(MP[:, 2]), np2list(MP[:, 3]), alternative='two-sided')[1])
# print(stats.mannwhitneyu(np2list(MP[:, 4]), np2list(MP[:, 5]), alternative='two-sided')[1])
# print(stats.mannwhitneyu(np2list(MRR[:, 0]), np2list(MRR[:, 1]), alternative='two-sided')[1])
# print(stats.mannwhitneyu(np2list(MRR[:, 2]), np2list(MRR[:, 3]), alternative='two-sided')[1])
# print(stats.mannwhitneyu(np2list(MRR[:, 4]), np2list(MRR[:, 5]), alternative='two-sided')[1])
#
# print(np.sum(MP, axis=0)/56)
# print(np.sum(MRR, axis=0)/56)
# print(np.sum(SUC, axis=0))
# print(np.sum(F, axis=0)/56)
#
# MP = pd.DataFrame(MP)
# MRR = pd.DataFrame(MRR)
# SUC = pd.DataFrame(SUC)
# F = pd.DataFrame(F)
#
# writer = pd.ExcelWriter('Precision.xlsx')
# MP.to_excel(writer, 'page_1', float_format='%.5f')
# writer.save()
# writer.close()
#
# writer = pd.ExcelWriter('RR.xlsx')
# MRR.to_excel(writer, 'page_1', float_format='%.5f')
# writer.save()
# writer.close()
#
# writer = pd.ExcelWriter('success.xlsx')
# SUC.to_excel(writer, 'page_1', float_format='%.5f')
# writer.save()
# writer.close()
#
# writer = pd.ExcelWriter('F1.xlsx')
# F.to_excel(writer, 'page_1', float_format='%.5f')
# writer.save()
# writer.close()
    # print(i)
    # print(traditional_read_score)
    # print(traditional_read_self_score)