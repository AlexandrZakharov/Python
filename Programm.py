from tkinter import *
from tkinter import ttk, filedialog, messagebox

from OpenKompas import *

lb = []  # список значений находящихся в listbox
dmin_value = None
dmax_value = None

# 21 - размер вала; b9 - класс допуска; 20.84 - Dmax; 20.788 - Dmin
values = {'21': {'b9': (20.84, 20.788), 'c11': (20.89, 20.76), 'h8': (21, 20.967), 'c8': (20.89, 20.857), 'e8': (20.96, 20.927), 'h10': (21, 20.916)},
     '23': {'h9': (23, 22.948), 'c10': (22.89, 22.806), 'c11': (22.89, 22.76), 'b10': (22.84, 22.756), 'a8': (22.7, 22.67), 'b11': (22.84, 22.71)},
     '24': {'x8': (24.087, 24.054), 'z8': (24.106, 24.073), 'u8': (24.074, 24.041), 'z7': (24.094, 24.073), 'n6': (24.028, 24.015), 'm6': (24.021, 24.008)}}


def selected_cb1(event):  # При выборе значения в cb1

    global combo1, dmin_value, dmax_value
    cb2.set('')  # Очищает значения классов допусков находящихся в комбобоксе cb2
    combo1 = event.widget.get()  # Значение, выбранное в комбобоксе cb1
    size_limits = []
    for size_limit in values.get(combo1).keys():  # size_limit - класс допуска (b9, c11, h8, ...)

        size_limits.append(size_limit)
        cb2['values'] = size_limits

    txt4['text'] = 'Dmin:'
    txt5['text'] = 'Dmax:'
    dmin_value = None  # Очищает значения Dmax
    dmax_value = None


def selected_cb2(event):  # При выборе значения в cb2
    global combo2, dmin_value, dmax_value, z1, z1_value, dmax, dmin

    z1 = {'IT6': 0.002,
          'IT7': 0.003,
          'IT8': 0.005,
          'IT9': 0.009,
          'IT10': 0.009,
          'IT11': 0.019}

    combo2 = event.widget.get()  # Значение, выбранное в комбобоксе cb2

    # Получаем значения Dmax и Dmin
    dmax = values.get(combo1).get(combo2)[0]
    dmin = values.get(combo1).get(combo2)[1]

    z1_value = float(z1.get('IT' + combo2[1:]))  # Получаем значение z1

    # Вычисляем значения Р-ПР и Р-НЕ
    dmax_value = dmax - z1_value
    dmin_value = dmin

    txt5['text'] = 'Dmax: %s' % dmax

    return combo2


def clicked():
    global dmin_value, dmax_value

    if dmin_value == None and dmax_value == None:
        messagebox.showerror('Ошибка', 'Заполните предыдущие поля!')
        return
    listbox.delete(0, END)
    listbox.insert(END, 'Годные валы:')
    listbox.insert(END, '')
    for i in lb:
        if str(dmin_value) < i[6:] < str(dmax_value):
            listbox.insert(END, i)
    listbox['fg'] = 'green'
    if listbox.size() == 2:
        messagebox.showerror('Сообщение', 'Ни один из валов не является годным!')


def LoadFile(ev):
    fn = filedialog.Open(root, filetypes=[('*.txt files', '.txt')]).show()
    if fn == '':
        return
    index = 0
    global lb

    file = open(fn, 'r')
    b = len(file.readlines())
    file.close()
    file = open(fn, 'r')
    while index != b:
        lb.append(file.readline())
        index += 1

    for i in lb:
        listbox.insert(END, i)
    lb = [line.rstrip() for line in lb]
    print(lb)


def Kompas():

    p = OpenKompas(dmax, dmin, z1_value)

    p.kompas()

    # Тонкие линии
    p.line(-p.width, 0, p.width * 3, 0)
    p.line(-p.width / 2, -p.width, p.width + p.width / 10, -p.width)
    p.line(p.width / 2 + p.width, -p.z1, p.width * 4, -p.z1)
    p.line(p.width / 2 + p.width, -p.width, p.width * 3.5, -p.width)

    # Прямоугольники
    p.rectangle(0, 0, -p.width)  # Поле допуска
    p.rectangle(p.width * 2, - p.z1/2, - p.z1)  # Верхнее отклонение
    p.rectangle(p.width * 2, -p.width + p.z1/2, -p.z1)  # Нижнее отклонение

    # Стрелки с надписями
    p.arrow(- (p.width / 2) - 6, -2 * p.width, - (p.width / 2) - 6, 0, 'Dmax')
    p.arrow(- p.width / 4, -2 * p.width, - p.width / 4, - p.width, 'Dmin')
    p.arrow(p.width * 3.9, -2 * p.width, p.width * 3.9, -p.z1, 'Р-ПР')
    p.arrow(p.width * 3.4, -2 * p.width, p.width * 3.4, -p.width, 'Р-НЕ')

    # Текст
    p.text(p.width//2.65, -p.width//1.78, 'Td')
    p.text(p.width + p.width/4, -p.z1 - p.width//12, 'ПР')
    p.text(p.width + p.width/4, -p.width - p.width//12, 'НЕ')

    # Штриховка
    p.hatching(p.width, 0, 0, -p.width, 45, p.width//2.8)
    p.hatching(3 * p.width, - p.z1/2, 2 * p.width, -p.z1 - p.z1/2, -45, 0)
    p.hatching(3 * p.width, -p.width + p.z1/2, 2 * p.width, -p.width + p.z1/2 - p.z1, -45, 0)

    # Размер z1
    p.arrow(p.width + 3*p.width/4, -p.width/3, p.width + 3*p.width/4, -p.z1, '')
    p.arrow(p.width + 3*p.width/4, p.width/2, p.width + 3*p.width/4, 0, 'z1')
    p.line(p.width + 3*p.width/4, 0, p.width + 3*p.width/4, -p.z1)


root = Tk()  # Создание окна программы

root.title('Главное окно')
root.resizable(False, False)  # Запрещает пользователю изменять размеры окна

# Располагаем окно программы по центру экрана
# root.winfo_screenwidth() - вычисляет размеры монитора
w = root.winfo_screenwidth()//2 - 143
h = root.winfo_screenheight()//2 - 225

root.geometry("286x450+%d+%d" % (w, h))

# Надписи
txt1 = Label(root, text='Размер вала:')
txt1.place(x=20, y=30)

txt2 = Label(root, text='Класс допуска:')
txt2.place(x=20, y=70)

txt3 = Label(root, text='Список валов:')
txt3.place(x=20, y=170)

txt4 = Label(root, text='Dmin:')
txt4.place(x=196, y=94)

txt5 = Label(root, text='Dmax:')
txt5.place(x=118, y=94)

# Комбобоксы
cb1 = ttk.Combobox(root, values=['21', '23', '24'], state='readonly')
cb1.place(x=120, y=30)
cb1.bind('<<ComboboxSelected>>', selected_cb1)
cb2 = ttk.Combobox(root, state='readonly')
cb2.set('')
cb2.place(x=120, y=70)
cb2.bind('<<ComboboxSelected>>', selected_cb2)

# Листбокс
listbox = Listbox(root, height=12, width=40, selectmode=EXTENDED)
listbox.place(x=20, y=190)

# Скролбар
scroll = Scrollbar(command=listbox.yview)
scroll.place(in_=listbox, x=224, relheight=1.0)
listbox.config(yscrollcommand=scroll.set)

# Кнопки
b = Button(root, text='Рассчитать',
           width='33',
           pady="8",
           activebackground='#c1c1c1',
           command=clicked
           )
b.place(x=21, y=394)

b2 = Button(root, text='Загрузить')
b2.bind("<ButtonRelease-1>", LoadFile)
b2.place(x=198, y=164)

b3 = Button(root, text='ОТОБРАЗИТЬ СХЕМУ', width='19', command=Kompas)
b3.place(x=120, y=120)

root.mainloop()  # Отображение окна программы