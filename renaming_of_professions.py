import tkinter as tk
from tkinter import filedialog, Entry, Label, Button
from openpyxl import load_workbook
import sqlite3
import os

def renaming_of_professions():
    """Переименение профессий"""

    print("Переименование профессий") # Переименование профессий

    filename = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')]) # Выбор файла
    print(f"Выбран файл: {filename}")

    workbook = load_workbook(filename=filename)  # Загружаем выбранный файл Excel
    sheet = workbook.active

    min_row = 1 # Начальная строка
    max_row = 19 # Конечная строка
    column = 5 # Начальная колонка (счет начинается с 1)

    # Считываем данные из колонки и заменяем их при необходимости
    for row in range(min_row, max_row + 1):
        cell = sheet.cell(row=row, column=column)  # Получаем ячейку
        table_column_1 = str(cell.value)  # Преобразуем значение в строку

        print(table_column_1)  # Выводим значение

        if table_column_1 == "нач. сектора (расчётного)":
            # print('Профессия "нач. сектора (расчётного)" не верная, верная "Начальник сектора расчётного"')
            # Заменяем неправильную профессию на правильную
            cell.value = "Начальник сектора расчётного"
            # print(f'Профессия исправлена на "Начальник сектора расчётного" в строке {row}')
        elif table_column_1 == "нач.уч-ка":
            # print('Профессия "нач.уч-ка" не верная, верная "Начальник участка"')
            # Заменяем неправильную профессию на правильную
            cell.value = "Начальник участка"
            # print(f'Профессия исправлена на "Начальник участка" в строке {row}')

    try:
        # Сохраняем изменения в том же файле
        workbook.save(filename)
        print(f"Изменения сохранены в файле: {filename}")
    except PermissionError:
        print(f"Файл {filename} открыт в другой программе")


if __name__ == '__main__':
    renaming_of_professions()