import tkinter as tk
from logging import root
from tkinter import Canvas, Scrollbar, Frame

from gui import root

guide = {
    '1':	'оплата по основной тарифной ставке',
'2':	'оплата по должностным окладам',
'18':	'оплата простоев',
'26':	'отпускные из фонда з/платы',
'31':	'компенсация уволенным',
'34':	'компенсация работающим',
'36':	'доплата за условия труда',
'38':	'вых.пособие при реорганиз.,сокращ.штата',
'103':	'доплата за работу в празд.дни',
'104':	'доплата за работу в выходные дни',
'109':	'доплата за расшир.зоны обслуж',
'112':	'доплата водителям за ненорм.раб.день',
'114':	'доплата за классность',
'122':	'доплата за совмещение',
'124':	'надбавка за высокие достижения в труде',
'130':	'доплата до мин.оплаты труда',
'134':	'оплата за работу в выходные дни',
'135':	'оплата за работу в праздничные дни',
'143':	'стипендия полн.кавалерам "шахтер.слава"',
'164':	'доплата за работу в военное время',
'206':	'доплата за работу в ночное время',
'207':	'доплата за работу в вечернее время',
'410':	'премии к праздничным дням и юбилеям',
'426':	'вознагражден.за выслугу лет',
'510':	'материальная помощь облагаемая (фз)',
'516':	'единоразовая мат.помощь(не облагаемая)',
'603':	'б/л по общему заболеванию(шахта)'
}

def withdrawal_of_all_surcharges():
    """Вывод всех доплат"""

    # Создаем главное окно
    # root.title("Все доплаты")

    root.title("Выбор файла")  # Заголовок окна

    root.geometry("400x400")  # Размер окна

    # Очистка содержимого окна и замена главного окна на новое окно
    for widget in root.winfo_children():
        widget.pack_forget()

    root.geometry('400x400')  # Устанавливаем размер окна

    # Создаем фрейм для прокрутки
    main_frame = Frame(root)
    main_frame.pack(fill="both", expand=True)

    # Добавляем холст для прокрутки
    canvas = Canvas(main_frame)
    canvas.pack(side="left", fill="both", expand=True)

    # Добавляем полосу прокрутки
    scrollbar = Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")

    # Связываем полосу прокрутки с холстом
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Внутренний фрейм для размещения текста
    content_frame = Frame(canvas)
    canvas.create_window((0, 0), window=content_frame, anchor="nw")

    # Добавляем информацию из словаря в content_frame
    for code, description in guide.items():
        label_text = f"Код: {code}, Описание: {description}"
        tk.Label(content_frame, text=label_text, font=("Arial", 10)).pack(anchor='w', padx=10, pady=2)


if __name__ == '__main__':
    withdrawal_of_all_surcharges()