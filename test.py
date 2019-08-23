"""
Модуль для тестирования
"""

from main import *
from openpyxl import load_workbook as lw


def test1():
    file = lw('Результаты 2019/Пример.xlsx')
    page = file.active

    for i in range(1, page.max_column + 1):
        print(page.cell(row=11,  column=i).value)


def test2():
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


def test3():
    filename1 = lw('Результаты 2019/Пример_11.xlsx')
    filename2 = lw('Результаты 2019/Пример_аппел.xlsx')
    filename3 = lw('Результаты 2019/Пример.xlsx')
    page1 = filename1.active
    page2 = filename2.active
    page3 = filename3.active
    cell = 'B10'
    print(page1[cell].value)
    print(page2[cell].value)
    print(page3[cell].value)


