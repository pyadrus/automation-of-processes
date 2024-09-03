import tkinter as tk

# Создаем главное окно
root = tk.Tk()
root.title("Пример окна с кнопкой")

# Функция, выполняемая при нажатии на кнопку
def on_button_click():
    print("Кнопка нажата!")

# Создаем кнопку
button = tk.Button(root, text="Нажми меня", command=on_button_click)
button.pack(pady=20)

# Запуск главного цикла окна
root.mainloop()
