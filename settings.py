"""
Параметры файлов с результатами,
настройки, которые можно изменить с помощью GUI (в разработке)
"""


'''Указатели на номера ячеек и столбцов в файлах'''
cell_check_version = 'A5'  # ячейка, по которой проверяем версию файла с результатами. Если пусто - 11 классы

versions = {  # словарь с вариантами столбцов для 9 и 11 классов
    'nine': {  # файл с результатами 9 классов
        'columns': (  # номера столбцов
            4,  # столбец D - класс
            12,  # столбец L - фамилия
            13,  # столбец M - имя
            14,  # столбец N - отчество
            16,  # столбец P - задания с кратким ответом
            18,  # столбец R - задания с развернутым ответом
            19,  # столбец S - первичный балл
            21  # столбец U - оценка
        ),
        'begin_row': 8,  # строка с первым учеником
        'subject_cell': 'A5'  # ячейка с названием предмета
    },
    'eleven': {  # файл с результатами 10-11 классов
        'columns': (  # номера столбцов
            7,  # столбец G - класс
            13,  # столбец M - фамилия
            14,  # столбец N - имя
            15,  # столбец O - отчество
            17,  # столбец Q - задания с кратким ответом
            19,  # столбец R - задания с развернутым ответом
            21,  # столбец U - первичный балл
            23  # столбец W - оценка в баллах
        ),
        'begin_row': 9,  # строка с первым учеником
        'subject_cell': 'A6'  # ячейка с названием предмета
    }
}