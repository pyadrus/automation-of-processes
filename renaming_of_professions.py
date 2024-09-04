from tkinter import filedialog

from openpyxl import load_workbook


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
            cell.value = "Начальник сектора расчётного"# Заменяем неправильную профессию на правильную
        elif table_column_1 == "нач.уч-ка":
            cell.value = "Начальник участка"# Заменяем неправильную профессию на правильную
        elif table_column_1 == "гл.бухгалтер":
            cell.value = "Главный бухгалтер"# Заменяем неправильную профессию на правильную
        elif table_column_1 == "вед. экономист":
            cell.value = "Ведущий экономист"# Заменяем неправильную профессию на правильную
        elif table_column_1 == "нач.отдела":
            cell.value = "Начальник отдела"# Заменяем неправильную профессию на правильную
        elif table_column_1 == "нач. сектора":
            cell.value = "Начальник сектора"# Заменяем неправильную профессию на правильную
        elif table_column_1 == "зам.гл.бух":
            cell.value = "Заместитель главного бухгалтера"# Заменяем неправильную профессию на правильную
        elif table_column_1 == "зам. директора по экономике":
            cell.value = "Заместитель директора по экономике"# Заменяем неправильную профессию на правильную
        elif table_column_1 == "бухгалтер 1 кат.(расчётного сектора)":
            cell.value = "Бухгалтер 1 категории расчётного сектора"# Заменяем неправильную профессию на правильную

    try:
        # Сохраняем изменения в том же файле
        workbook.save(filename)
        print(f"Изменения сохранены в файле: {filename}")
    except PermissionError:
        print(f"Файл {filename} открыт в другой программе")


if __name__ == '__main__':
    renaming_of_professions()
