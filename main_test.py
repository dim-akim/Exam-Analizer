"""
Программа анализирует файлы с результатами экзаменов ОГЭ и ЕГЭ
Сводные данные будут отражены на новой странице в файле
Версия 0.5
"""


from openpyxl import load_workbook as lw
from datetime import datetime
import os

from settings import *


class ResultFile(lw):
    """
    Файл с результатами экзаменов
    """
    def __init__(self, file):
        """
        Открывает файл Excel с дополнительными атрибутами:
        sheet - активная страница для работы
        columns - кортеж с номерами столбцов, отвечающих за соответствующие поля
        begin_row - номер строки с первым участником
        subject_cell - ячейка с названием предмета
        :param file: путь к файлу Excel. Тип данных - str
        """
        # открываем файл excel
        lw.__init__(self, file)
        # выбираем активную страницу для работы
        self.sheet = self.active
        # проверяем версию файла - для 9 или 11 класса
        page_version = self.set_page_version()
        # выбираем нужное из словаря с параметрами
        self.columns = versions[page_version]['columns']  # номера столбцов
        self.begin_row = versions[page_version]['begin_row']  # первая строка с данными ученика
        self.subject_cell = versions[page_version]['subject_cell']  # название предмета

    def set_page_version(self):
        """
        Возвращает версию файла - для 9 или 11 класса в виде ключа
        для словаря versions
        :return: 'nine' или 'eleven' в зависимости от версии файла
        """
        # Функция возвращает версию файла - для 9 или 11 класса
        return 'nine' if self.sheet[cell_check_version] is not '' else 'eleven'

    def get_student(self, row):
        """
        Набирает в список данные одного ученика
        :param row: номер строки, на которой находятся данные ученика
        :return: список в порядке Класс, Ф, И, О, Краткий ответ, Развернутый ответ, Первичный балл, Оценка
        """
        student = []
        for i in self.columns:
            student.append(row=row, col=i).value
        return student
