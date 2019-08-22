"""
Модуль для тестирования
"""

from openpyxl import load_workbook as lw


file = lw('Результаты 2019/Пример.xlsx')
page = file.active

for i in range (1,page.max_column + 1):
    print (page.cell(row=11, column=i)).value

