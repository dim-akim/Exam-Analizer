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
    filenames = [
        'Пример_11.xlsx',
        'Пример_аппел.xlsx',
        'Пример.xlsx'
    ]
    cell = 'B10'
    for name in filenames:
        book = lw(name)
        page = book.active
        print(page[cell].value)


# test1()
# test2()
test3()
