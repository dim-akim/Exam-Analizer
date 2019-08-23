"""
Модуль для тестирования
"""
"""
from openpyxl import load_workbook as lw


file = lw('Результаты 2019/Пример.xlsx')
page = file.active

for i in range (1,page.max_column + 1):
    print (page.cell(row=11, column=i)).value
"""

from main import *


filename = 'Результаты 2019/Пример.xlsx'
file = ResultFile(filename)
# print(file.name)
# print(file.sheet)
count = 0
for i in file.students.keys():
    print(i)
    for j in file.students[i]:
        count += 1
        print(j)
print(count)
