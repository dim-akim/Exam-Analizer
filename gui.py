"""
Модуль создает GUI для облегчения работы с программой-анализатором.
Основан на модуле tkinter
"""

from tkinter import *
import os


directory = 'Результаты 2019'
os.chdir(directory)
list_of_files = [i for i in os.listdir() if '013273' in i]
