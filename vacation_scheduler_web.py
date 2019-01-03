#-*- coding:utf-8 -*-

import os, sys
from random import *
from openpyxl import Workbook
from openpyxl.styles import Alignment

def find_time_with_least_student(ml):
    target = [0,0]
    num = 1000
    for i in range(len(ml)):
        for j in range(len(ml[i])):
            if(ml[i][j] == [-1]):
                continue
            if(num > len(ml[i][j])):
                num = len(ml[i][j])
                target = [i, j]
    return target

def is_non_allocate(ml):
    for i in range(len(ml)):
        for j in range(len(ml[i])):
            if(ml[i][j] != [-1]):
                return True
    return False

def find_inno_student(students, counts):
    sames = [students[0]]

    for i in range(1, len(students)):
        if(counts[students[i]] < counts[sames[0]]):
            target = students[i]
            sames = [students[i]]
        elif(counts[students[i]] == counts[sames[0]]):
            sames.append(students[i])

    shuffle(sames)

    return sames[0]

file_name = sys.argv[1]
shift_num = 4
# shift_num = int(input("Enter shift num(the number of time for one day, it must be same for all day) : "))

f = open("./input/" + file_name, 'r', encoding='UTF8')

#################################################
f.readline() #useless first line
#################################################
raw_dates = f.readline() # data with date
raw_dates_list = raw_dates.split(",")

real_dates = []
for i in range(len(raw_dates_list)):
    if raw_dates_list[i] != '':
        real_dates.append(raw_dates_list[i])
day = len(real_dates)
#################################################
raw_times = f.readline() # data with time
raw_times_list = raw_times.split(",")

real_times = []
for i in range(shift_num):
    real_times.append(raw_times_list[i+1])
#################################################

# Extract data from csv file
namelist = []
timedict = dict()
student_num = 0

for line in f:
    a = line.split(",")

    name = a.pop(0)
    matrix = []

    day = int(len(a)/4)
    for i in range(int(len(a)/shift_num)):
        day_ox = []
        for j in range(shift_num):
            if 'X' in a[shift_num*i+j].upper():
                day_ox.append(0)
            else:
                day_ox.append(1)
        matrix.append(day_ox)

    timedict[student_num] = matrix

    namelist.append(name)
    student_num += 1
#################################################
# make master_list that has student list who can do the time
master_list = []
for i in range(day):
    newone = []
    for j in range(shift_num):
        newone.append([])
    master_list.append(newone)

for index in range(student_num):
    target = timedict[index]
    for j in range(len(target)):
        for k in range(len(target[j])):
            if(target[j][k] == 1):
                master_list[j][k].append(index)
                
#################################################
# copy master_list and make alloc
cpy_ml = []
alloc = []
for i in range(len(master_list)):
    cpy_ml.append([])
    alloc.append([])
    for j in range(len(master_list[i])):
        cpy_ml[i].append(master_list[i][j][:])
        alloc[i].append([])

#################################################
# Process the matching
counts = [0] * student_num
while is_non_allocate(cpy_ml):
    target = find_time_with_least_student(cpy_ml)
    student_list = cpy_ml[target[0]][target[1]]
    cpy_ml[target[0]][target[1]] = [-1]

    if(len(student_list) == 0):
        winner = -1
    else:
        winner = find_inno_student(student_list, counts)
        counts[winner]+=1
    alloc[target[0]][target[1]] = winner
print("Process doned")

#################################################

wb = Workbook()
sheet1 = wb.active
sheet1.title = "상근시간표"

sheet1["C2"] = real_dates[0]
sheet1["D2"] = real_dates[1]
sheet1["E2"] = real_dates[2]
sheet1["F2"] = real_dates[3]
sheet1["G2"] = real_dates[4]

sheet1["B3"] = "1 ~ 3"
sheet1["B4"] = "3 ~ 5"
sheet1["B5"] = "5 ~ 7"
sheet1["B6"] = "7 ~ 9"

start_alpha = "C"
for i in range(len(alloc)):
    alpha = chr(ord(start_alpha) + i)
    row_num = 3
    for j in range(len(alloc[i])):
        target = ""
        if (alloc[i][j] == -1):
            target = target + "X" + ", "
        else:
            target = target + namelist[alloc[i][j]] + ", "
        sheet1[alpha + str(row_num)] = target[:-2]
        row_num += 1


ws = sheet1
dims = {}
max_v = 0
for row in ws.rows:
    for cell in row:
        if cell.value:
            # dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value)) * 3))
            if(max_v < len(str(cell.value))):
                max_v = len(str(cell.value))
for row in ws.rows:
    for cell in row:
        if cell.value:
            dims[cell.column] = max_v * 1.5

for col, value in dims.items():
    ws.column_dimensions[col].width = value

dir_path = os.path.dirname(os.path.realpath(__file__))
new_filename = dir_path + "/result/" + file_name.split(".")[0] + "_result.xlsx"

wb.save(filename=new_filename)

