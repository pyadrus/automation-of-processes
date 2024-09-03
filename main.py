import tkinter as tk
from tkinter import filedialog

# Создаем главное окно
root = tk.Tk()
root.title("Выбор файла") # Заголовок окна

root.geometry("400x400") # Размер окна

# Создаем текст (метка)
label = tk.Label(root, text="Нажмите на кнопку ниже, чтобы выбрать файл:")
label.pack(pady=20)

# Функция, выполняемая при нажатии на кнопку
def select_file():
    # Открытие диалогового окна для выбора файла
    file_path = filedialog.askopenfilename()

    # Если файл выбран, то выводим его путь
    if file_path:
        print(f"Выбран файл: {file_path}")


# Создаем кнопку
button = tk.Button(root, text="Выбрать файл", command=select_file)
button.pack(pady=20)

# Запуск главного цикла окна
root.mainloop()
