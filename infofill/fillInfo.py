import pandas as pd
import random
import math

def getHigh(mid=95):
    high = mid + 3
    if mid > 97:
        high = 100
        mid = 97
    ss = random.randint(mid, high)
    return ss


def getLow(mid=95):
    ss = random.randint(mid-4, mid)
    return ss


def equals(A, B):
    if A == B:
        return True
    return False


def calculate(tt):
    # ['线性', '树', '图', '查找', '排序']
    sum = tt[0] * 0.3 + tt[1] * 0.3 + tt[2] * 0.14 + tt[3] * 0.12 + tt[4] * 0.14
    # sum = sum * 10
    return round(sum)
    # round(sum, 1)

table = pd.read_excel('fillInfo.xlsx')

print(table.shape)
row, col = table.shape
count = 1
while count < row:
    flag = True
    while flag:
        # print(table.iloc[count, 6])
        ordinary = round(table.iloc[count, 6])
        # print(ordinary)
        listt = [getLow(ordinary), getLow(ordinary), getLow(ordinary), getLow(ordinary), getHigh(ordinary)]
        sum = calculate(listt)
        # ordinary = int(float(table.iloc[count, 6]) * 10)
        if equals(sum, ordinary):
            table.loc[count:count+1, ('线性', '树', '图', '查找', '排序', 'Test')]= listt+[sum]
            print(sum, ordinary)
            count = count + 1
            flag = False
            break

table.to_excel('fillInfo2.xlsx')

