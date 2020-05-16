import xlrd
import re
import copy
import pyperclip

# name of the standard students list
standard_data = xlrd.open_workbook('181.xls')
# excel file of today's attendance
class_data = xlrd.open_workbook('students.xlsx')
standard_data_table = standard_data.sheet_by_index(0)
class_data_table = class_data.sheet_by_index(0)
standard_student_namelist = standard_data_table.col_values(4)
class_namelist = class_data_table.col_values(3)
# delete some useless data distraction
del standard_student_namelist[0:6]
del class_namelist[0:5]

standard_name = []
class_name = []
for i in range(len(standard_student_namelist)):  # students in the standard namelist
    element1 = standard_student_namelist[i]
    # only apply to Chinese
    name1 = re.findall(r'[\u4e00-\u9fa5]', element1)
    if len(name1) == 0:
        continue
    # print(name2)
    # deal with some special cases
    if name1[0] == '海':
        del name1[0]
    name1 = ''.join(name1)
    if name1 == '刘新宇':
        cell_LXY = standard_data_table.row(i + 6)[6].value
        class_LXY = cell_LXY[-1]
        name1 = class_LXY + '-' + name1
    standard_name.append(name1)
# print(standard_name)
for j in range(len(class_namelist)):  # students joining the class
    element2 = class_namelist[j]
    # only apply to Chinese
    name2 = re.findall(r'[\u4e00-\u9fa5]', element2)
    if len(name2) == 0:
        continue
    # print(name2)
    # deal with some special cases
    if name2[0] == '海':
        del name2[0]
    if name2[-1] == '组':
        del name2[-1]
    if name2[0] == '号':
        del name2[0]
    name2 = ''.join(name2)
    if name2 == '刘新宇':
        class_LXY = element2[0]
        name2 = class_LXY + '-' + name2
    class_name.append(name2)

Notcome_list = copy.deepcopy(standard_name)
# [class_name1.append(i) for i in class_name if not i in class_name1]
class_name1 = list(set(class_name))
# print(class_name1)
# print(standard_name)
print(len(Notcome_list))
print(len(class_name1))
for k in range(len(standard_name)):
    for w in range(len(class_name1)):
        if standard_name[k] == class_name1[w]:
            # print(standard_name[w])
            Notcome_list.remove(standard_name[k])
            break
        else:
            continue

print(Notcome_list)

# copy the name to the clipboard 
name = ' '.join(Notcome_list)
print(name)
pyperclip.copy(name)

# print(class_namelist)
