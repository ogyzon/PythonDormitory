'''Жоров Евгений Александрович
гр. 10701123
@vanasokolov844@gmail.com
Курсовой проект по дисциплине "Языки программирования"
Минск 2024'''

'''Учет дежурств в общежитии'''

'''Добавить год. При возможности реализовать штуку с пересылкой данных к примеру в ворд файл (типо кнопка отослать коменданту).
В будущем добавлю сортировки.'''

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading

#Класс для Label
class MyLabel(tk.Label):

    def __init__(self, master, text = None, x = 0, y = 0, font_size=12, font_weight="normal", justify="left", bg="#f0f0f0", image = None, **kwargs):

        font = ("Montserrat", font_size, font_weight)

        super().__init__(master, text = text, font = font, justify = justify, bg = bg, image = image, **kwargs)
        self.place(x=x, y=y)

#Класс для кнопки
class MyButton(tk.Button):

    def __init__(self, master, text, width, height, fg = "#000000", bg = "#f0f0f0", image = None, compound = None, font_size = 9, font_weight = "normal", command = None, x = 0, y = 0, **kwargs):

        font = ("Montserrat", font_size, font_weight)

        super().__init__(master, text=text, font = font, width = width, height = height, fg = fg, bg = bg, image = image, compound = compound, command = command, **kwargs)
        self.place(x = x, y = y)







'''class Inactivity():
    def __init__(self):'''

#Класс приложения
class MyApp(tk.Tk):

    def __init__(self):

        super().__init__()

        self.title("Учет дежурств")
        self.geometry('800x800')
        self.geometry('+500+100')
        self.resizable(width = False, height = False)

        #Словарь для хранения записанных дат
        self.dictZapisanye = {}

        # Для перевода из слова в номер месяца
        self.DictForMonthes = {"Январь": '01', "Февраль": "02", "Март": "03", "Апрель": "04",
                               "Май": "05", "Июнь": "06", "Июль": "07", "Август": "08",
                               "Сентябрь": "09", "Октябрь": "10", "Ноябрь": "11", "Декабрь": "12"}

        self.stringDay = ""
        self.stringMonthNum = ""
        self.FullConcatenate = ""

        #Нужно
        self.concatenateString = ""
        self.monthUser = ""


        #Фрейм самого первого Splash окна
        self.frame1 = tk.Frame(self, width = 800, height = 800)
        self.frame1.place(x = 0, y = 0)

        #Фрейм второго (основного окна)
        self.frameMain = tk.Frame(self, width = 800, height = 800)

        #Фрейм для окна об авторе
        self.frameAuthor = tk.Frame(self, width = 800, height = 800)

        self.frameAbtProgramm = tk.Frame(self, width = 800, height = 800)

        self.canvasAbtProgramm = tk.Canvas(self.frameAbtProgramm, width = 300, height = 800, bg = '#9681F0')
        self.canvasAbtProgramm.place(x = 0, y = 0)

        self.canvas1 = tk.Canvas(self.frameMain, width=750, height=600, bg='#eedcfc', borderwidth = 2, relief = 'solid')
        self.canvas1.place(x=0, y=50)

        self.textPlace = tk.Text(self.frameMain, state = 'normal', font = ("Montserrat", 8, 'bold'))
        self.textPlace.place(x=20, y=210, width=600, height=400)
        self.textPlace.insert(tk.END, f"Фамилия\t\t  Имя\t\tБлок\t\tTелефон\t\t   Дата\t\tВремя\n\n")
        self.textPlace.config(state = 'disabled')
      
        #Лейблы с текстом Splash окна
        self.label1 = MyLabel(self.frame1, text = 'Белорусский национальный технический университет', x = 195, y = 20)

        self.label2 = MyLabel(self.frame1, text = 'Факультет информационных технологий и робототехники', x = 175, y = 50)


        self.label3 = MyLabel(self.frame1, text = 'Кафедра программного обеспечения информационных систем и технологий', x = 120, y = 80)

        self.label4 = MyLabel(self.frame1, text='Курсовой проект',font_size=16, font_weight = 'bold', x = 300, y = 160)

        self.label5 = MyLabel(self.frame1, text='По дисциплине "Языки программирования"', font_size=16, font_weight='bold', x = 160, y = 200)

        self.label6 = MyLabel(self.frame1, text='Выполнил: студент группы 10701123',font_size = 12, font_weight = 'bold', x = 340, y = 360)

        self.label7 = MyLabel(self.frame1, text='Жоров Евгений Александрович', font_size = 12, font_weight = 'bold', x = 340, y = 390)

        self.label8 = MyLabel(self.frame1, text = 'Преподаватель: к.ф.-м.н., доц.', font_size=12, font_weight='bold', x = 340, y = 450)

        self.label9 = MyLabel(self.frame1, text = 'Сидорик Валерий Владимирович', font_size=12, font_weight='bold', x = 340, y = 480)

        self.label10 = MyLabel(self.frame1, text = 'Минск 2024', font_size=12, font_weight='bold', x = 350, y = 630)

        self.label11 = MyLabel(self.frame1, text = 'Учет дежурств в общежитии',font_size=16, font_weight='bold', x = 235, y = 240)

        #Лейблы с текстом окна об авторе
        self.label12 = MyLabel(self.frameAuthor, text = 'Автор', font_size=14, font_weight='bold', x = 350, y = 470)

        self.label13 = MyLabel(self.frameAuthor, text = 'Студент группы 10701123', font_size=12, font_weight='bold', x = 280, y = 500)

        self.label14 = MyLabel(self.frameAuthor, text = 'Жоров Евгений Александрович', font_size=12, font_weight='bold', x = 260, y = 530)

        self.label15 = MyLabel(self.frameAuthor, text = 'vanasokolov844@gmail.com', font_size=12, font_weight='bold', x = 270, y = 560)

        #Лейблы для текста окна о программе
        self.label16 = MyLabel(self.frameAbtProgramm, text='Учет дежурств в общежитии', font_size=18, font_weight='bold', x = 380, y = 40)

        self.label17 = MyLabel(self.frameAbtProgramm, text = 'Программа позволяет:', font_size=13, font_weight='bold', x = 440, y = 90)

        self.label18 = MyLabel(self.frameAbtProgramm, text = '1. Записывать ФИО и номер телефона дежурного\n'
                                                         '2. Сохранять записи в файл\n'
                                                         '3. Просматривать результат в главном окне\n'
                                                         '4. Удалять данные о дежурстве\n'
                                                         '5. Записывать и считывать данные с Excel', font_size=11, justify = 'left', bg = '#eddcfc', x = 350, y = 130)

        self.label19 = MyLabel(self.frameAbtProgramm, text = 'Версия: 1.0.0.2024', font_size=10, x = 350, y = 710)

        #Лейблы с текстом главного окна
        self.label20 = MyLabel(self.frameMain, text = "Данные о жильцах", font_size=11, font_weight='bold', bg = '#eedcfc', x = 300, y = 15)

        self.label21 = MyLabel(self.frameMain, text="Фамилия", font_size=10, bg='#eedcfc', x = 20, y = 70)

        self.label22 = MyLabel(self.frameMain, text="Имя", font_size=10, bg='#eedcfc', x = 150, y = 70)

        self.label23 = MyLabel(self.frameMain, text="Блок", font_size=10, bg='#eedcfc', x = 280, y = 70)

        self.label24 = MyLabel(self.frameMain, text="Номер телефона", font_size=10, bg='#eedcfc', x = 400, y = 140)

        self.label24 = MyLabel(self.frameMain, text="Время", font_size=10, bg='#eedcfc', x = 280, y = 140)

        self.label25 = MyLabel(self.frameMain, text="Месяц",font_size=10, bg='#eedcfc', x = 20, y = 140)

        self.label26 = MyLabel(self.frameMain, text="День", font_size=10, bg='#eedcfc', x = 150, y = 140)

        self.label27 = MyLabel(self.frameMain, text="Комната", font_size=10, bg='#eedcfc', x = 400, y = 70)

        #Entry поля на главном окне
        #Ввод фамилии
        self.entry1 = tk.Entry(self.frameMain, width = 15)
        self.entry1.place(x = 20, y = 100)

        #Ввод имени
        self.entry2 = tk.Entry(self.frameMain, width=15)
        self.entry2.place(x=150, y=100)

        #Ввод блока
        self.entry3 = tk.Entry(self.frameMain, width=10)
        self.entry3.place(x=280, y=100)

        #Ввод номера тф
        self.entry4 = tk.Entry(self.frameMain, width=18)
        self.entry4.place(x=400, y=170)


        self.dictForMonthDays = {"Январь": 31,
                                 "Февраль": 28,  # Простой год
                                 "Март": 31,
                                 "Апрель": 30,
                                 "Май": 31,
                                 "Июнь": 30,
                                 "Июль": 31,
                                 "Август": 31,
                                 "Сентябрь": 30,
                                 "Октябрь": 31,
                                 "Ноябрь": 30,
                                 "Декабрь": 31}

        #Ввод месяца
        self.comboboxMonth = ttk.Combobox(self.frameMain, values=['Январь', 'Февраль', 'Март',
                                                                  'Апрель', 'Май', 'Июнь',
                                                                  'Июль', 'Август', 'Сентябрь',
                                                                  'Октябрь', 'Ноябрь', 'Декабрь'], width=12, state="readonly")
        self.comboboxMonth.place(x=20, y=170)
        self.comboboxMonth.bind("<<ComboboxSelected>>", self.UpdateDays)

        # Ввод дня
        self.comboboxDay = ttk.Combobox(self.frameMain, width=12, state="readonly")
        self.comboboxDay.place(x=150, y=170)
        self.comboboxDay.bind("<<ComboboxSelected>>", self.getInfoComboboxDay)

        self.comboboxTimeValues = ['8.00 - 10.30','10.30 - 13.00','13.00 - 15.30','15.30 - 18.00','18.00 - 20.00']

        #Ввод времени
        self.comboboxTime = ttk.Combobox(self.frameMain, values = self.comboboxTimeValues, width = 11, state="readonly")
        self.comboboxTime.place(x = 280, y = 170)

        #Ввод комнаты
        self.comboboxRoomInBlock = ttk.Combobox(self.frameMain, values = ['А', 'Б'], width = 11, state = 'readonly')
        self.comboboxRoomInBlock.place(x = 400, y = 100)

        #---------------------------------------------------------------------------------------КАРТИНКИ
        self.imageDormitory = tk.PhotoImage(file = 'images//iconDormitory.png')
        self.labelForDormitory = MyLabel(self.frame1, image = self.imageDormitory, x = 160, y = 370)
        self.labelForDormitory.image = self.imageDormitory
        #self.labelForDormitory.place(x = 160, y = 370)

        self.imageMe = tk.PhotoImage(file = 'images//Me.png')
        self.resizedImageMe = self.imageMe.subsample(3)
        self.labelForMe = MyLabel(self.frameAuthor, image = self.resizedImageMe, x = 230, y = 30)
        self.labelForMe.image = self.resizedImageMe

        self.imageForAbtProgramm1 = tk.PhotoImage(file = 'images//dormitory2.png')
        self.labelForAbtProgramm1 = MyLabel(self.frameAbtProgramm, image = self.imageForAbtProgramm1, bg = "#9681F0", x = 30, y = 230)
        self.labelForAbtProgramm1.image = self.imageForAbtProgramm1

        self.imageForAbtProgramm2 = tk.PhotoImage(file='images//room.png')
        self.labelForAbtProgramm2 = MyLabel(self.frameAbtProgramm, image= self.imageForAbtProgramm2, bg = '#9681F0', x=150, y=235)
        self.labelForAbtProgramm2.image = self.imageForAbtProgramm2

        self.imageForAbtProgramm3 = tk.PhotoImage(file='images//file.png')
        self.labelForAbtProgramm3 = MyLabel(self.frameAbtProgramm, image= self.imageForAbtProgramm3, bg = '#9681F0', x=30, y=350)
        self.labelForAbtProgramm3.image = self.imageForAbtProgramm3

        self.imageForAbtProgramm4 = tk.PhotoImage(file='images//excel.png')
        self.labelForAbtProgramm4 = MyLabel(self.frameAbtProgramm, image= self.imageForAbtProgramm4, bg = '#9681F0', x=150, y=350)
        self.labelForAbtProgramm4.image = self.imageForAbtProgramm4

        self.imageForAbtProgramm5 = tk.PhotoImage(file='images//human2.png')
        self.labelForAbtProgramm5 = MyLabel(self.frameAbtProgramm, image= self.imageForAbtProgramm5, bg='#9681F0', x=30, y=460)
        self.labelForAbtProgramm5.image = self.imageForAbtProgramm5

        self.imageForAbtProgramm6 = tk.PhotoImage(file='images//telephone.png')
        self.labelForAbtProgramm6 = MyLabel(self.frameAbtProgramm, image= self.imageForAbtProgramm6, bg='#9681F0', x=150, y=460)
        self.labelForAbtProgramm6.image = self.imageForAbtProgramm6

        #Картинка для кнопки выйти
        self.imageExit = tk.PhotoImage(file = 'images//exitMain.png')

        #Картинка для кнопки назад
        self.imageBack = tk.PhotoImage(file = 'images//backMain.png')

        #Картинка для кнопки об авторе
        self.imageHuman = tk.PhotoImage(file = 'images//human.png')

        #Картинка для кнопки о программе
        self.imageAbtProgramm = tk.PhotoImage(file = 'images//about.png')

        #Картинка для кнопки на главную
        self.imageToTheMain = tk.PhotoImage(file = 'images//home.png')

        #Картинка для кнопки записаться
        self.imageWrite = tk.PhotoImage(file = 'images//write.png')

        #Картинка для кнопки очистки
        self.imageClear = tk.PhotoImage(file = 'images//clear.png')

        #Картинка для кнопки удалить данные
        self.imageDelete = tk.PhotoImage(file = 'images//delete.png')

        #Картинка для кнопки выбора сортировки
        self.imageSort = tk.PhotoImage(file = 'images//sort.png')
        #---------------------------------------------------------------------------------КАРТИНКИ

        #---------------------------------------------------------------------------------КНОПКИ

        #Кнопка далее
        self.buttonStart = MyButton(self.frame1, text = 'Далее', fg = 'white', width = 30, height = 3,
                                    font_size = 9, font_weight='bold', bg = '#8251FE', command = self.openSecondWindow, x = 150, y = 700)

        #Кнопка выход
        if self.frameMain:
            self.buttonExit = MyButton(self.frameMain, text = 'Выход', width = 100, height = 30, image = self.imageExit,
                                   compound = tk.RIGHT, font_size=9, font_weight='bold', command = self.exitApp, x = 525, y = 700)

        if self.frame1:
            self.buttonExit = MyButton(self.frame1, text = 'Выход', width = 200, height = 45, image = self.imageExit,
                                       compound = tk.RIGHT, font_size=9, font_weight='bold', command = self.exitApp, x = 410, y = 700)

        #Кнопка назад УНИВЕРСАЛЬНАЯ
        if self.frameMain:
            self.buttonBack = MyButton(self.frameMain, text = 'Назад', width = 100, height = 30, image = self.imageBack,
                                       compound = tk.RIGHT, font_size=9, font_weight='bold', command = self.backFromMainto1st, x = 645, y = 700)

        if self.frameAbtProgramm:
            self.buttonBack = MyButton(self.frameAbtProgramm, text='Назад', width=100, height=30, image=self.imageBack,
                                       compound=tk.RIGHT,font_size=9, font_weight='bold', command=self.backFromAbtProgrammToMain, x = 645, y = 700)

        #Кнопка об авторе (добавить потом это окно)
        self.buttonAuthor = MyButton(self.frameMain, text = 'Об авторе', width = 100, height = 30, image = self.imageHuman,
                                     compound = tk.RIGHT,font_size=9, font_weight='bold', command = self.openAuthorFromMain, x = 200, y = 700)

        #Кнопка о программе
        self.buttonAbtProgramm = MyButton(self.frameMain, text = 'О программе', width = 120, height = 30, image = self.imageAbtProgramm,
                                          compound = tk.RIGHT,font_size=9, font_weight='bold', command = self.openAbtProgramm, x = 60, y = 700)

        #Кнопка на главную
        self.buttonToTheMain = MyButton(self.frameAuthor, text = 'На главную', width = 250, height = 30, image = self.imageToTheMain,
                                        compound = tk.RIGHT,font_size=9, font_weight='bold', command = self.backFromAuthorToMain, x = 265, y = 700)
        #Кнопка записаться
        self.buttonZapisat = MyButton(self.frameMain, text = 'Записаться', width = 80, height = 60, image = self.imageWrite,
                                      compound = tk.TOP,font_size=9, font_weight='bold', command = self.GetResults, x = 645, y = 100)

        #Кнопка очистить ввод
        self.buttonClear = MyButton(self.frameMain, text = "Очистить\nввод", width = 80, height = 60, image = self.imageClear,
                                    compound = tk.TOP,font_size=9, font_weight='bold', command = self.clearEntry, x = 645, y = 210)

        #Кнопка удалить данные
        self.buttonDeleteData = MyButton(self.frameMain, text = 'Удалить\nданные', width = 80, height = 60, image = self.imageDelete,
                                         compound = tk.TOP, font_size = 9, font_weight='bold', command = self.deleteAllData, x = 645, y = 320)

        # Кнопка выбрать сортировку
        self.buttonChooseSort = MyButton(self.frameMain, text='Выбрать\nсортировку', width=80, height=60, image=self.imageSort,
                                         compound=tk.TOP, font_size=9, font_weight='bold',command = self.chooseSort, x=645, y=425)

    #-----------------------------------------------------------------------------------------------------------------КНОПКИ

    #-----------------------------------------------------------------------------------------------------------------ФУНКЦИИ
    #Закрытие окна
    def exitApp(self):
        self.destroy()

    #Открытие второго окна (забиндить и добавить эл-ты второго окна в фрейм 2)
    def openSecondWindow(self):
        self.frame1.pack_forget()
        self.frameMain.pack()

    #Возвращение со второго окна на первое(с основного в Splash)
    def backFromMainto1st(self):
        self.frameMain.pack_forget()
        self.frame1.pack()

    #Открытие окна об авторе с главного окна
    def openAuthorFromMain(self):
        self.frameMain.pack_forget()
        self.frameAuthor.pack()

    #Возвращение с окна об авторе на главную
    def backFromAuthorToMain(self):
        self.frameAuthor.pack_forget()
        self.frameMain.pack()

    def openAbtProgramm(self):
        self.frameMain.pack_forget()
        self.frameAbtProgramm.pack()

    def backFromAbtProgrammToMain(self):
        self.frameAbtProgramm.pack_forget()
        self.frameMain.pack()

    #Очистка всех полей при нажатии кнопки Очистить ввод
    def clearEntry(self):

        self.comboboxMonth.config(state = "normal")
        self.comboboxDay.config(state = "normal")
        self.comboboxTime.config(state = "normal")
        self.comboboxRoomInBlock.config(state = "normal")

        self.entry1.delete(0, tk.END)
        self.entry2.delete(0, tk.END)
        self.entry3.delete(0, tk.END)
        self.entry4.delete(0, tk.END)
        self.comboboxMonth.delete(0, tk.END)
        self.comboboxDay.delete(0, tk.END)
        self.comboboxTime.delete(0, tk.END)
        self.comboboxRoomInBlock.delete(0, tk.END)

        self.comboboxMonth.config(state="readonly")
        self.comboboxDay.config(state="readonly")
        self.comboboxTime.config(state="readonly")
        self.comboboxRoomInBlock.config(state="readonly")

    def GetResults(self):

        # Получение данных от пользователя
        self.lastName = self.entry1.get()
        self.firstName = self.entry2.get()
        self.Block = self.entry3.get()
        self.TelNumber = self.entry4.get()
        self.Month = self.comboboxMonth.get()
        self.Day = self.comboboxDay.get()
        self.Time = self.comboboxTime.get()
        self.Room = self.comboboxRoomInBlock.get()

        # Проверка на корректность ввода фамилии
        for symbol in self.lastName:
            if symbol.isdigit() or symbol in "_-+=!@#$%*^()&?~/.,":
                tk.messagebox.showerror("Некорректный ввод", "Фамилия введена некорректно!")
                self.entry1.delete(0, tk.END)
                return

        # Проверка на корректность ввода имени
        for symbol in self.firstName:
            if symbol.isdigit() or symbol in "_-+=!@#$%*^()&?~/., ":
                tk.messagebox.showerror("Некорректный ввод", "Имя введено некорректно!")
                self.entry2.delete(0, tk.END)
                return

        # Проверка на корректность ввода блока
        for symbol in self.Block:
            if not symbol.isdigit():
                tk.messagebox.showerror("Некорректный ввод", "Блок введен некорректно!")
                self.entry3.delete(0, tk.END)
                return

        # Проверка на корректность ввода номера телефона
        for symbol in self.TelNumber:
            if not symbol.isdigit() and symbol not in "+":
                tk.messagebox.showerror("Некорректный ввод", "Номер введен некорректно!")
                self.entry4.delete(0, tk.END)
                return

        self.textPlace.config(state='normal')

        # Проверка все ли поля введены
        if not all([self.lastName, self.firstName, self.Block, self.TelNumber, self.Month, self.Day, self.Time,
                    self.Room]):
            tk.messagebox.showerror("Ошибка", "Введите все данные!")
            return

        # Цикл для вывода даты и записи занятых дат в словарь
        for nameOfMonth, MonthNumber in self.DictForMonthes.items():
            if self.Month == nameOfMonth:
                if int(self.Day) < 10:
                    self.concatenateString = "0" + self.Day + "." + MonthNumber
                else:
                    self.concatenateString = self.Day + "." + MonthNumber

                # Запись занятой даты в словарь
                self.recordDuty(self.concatenateString, self.Time)  # Вызов функции записи данных

        self.textPlace.insert(tk.END,
                              f"{(str(self.lastName)).capitalize()}\t\t  {(str(self.firstName)).capitalize()}\t\t"
                              f"{str(self.Block) + str(self.Room)}\t\t{str(self.TelNumber)}\t\t   {str(self.concatenateString)}\t\t{str(self.Time)}\n")

        self.clearEntry()

        self.textPlace.config(state='disabled')

    #Функция для изменения чисел при выборе месяца
    def UpdateDays(self, event):

        self.comboboxDay.set("")
        self.monthUser = self.comboboxMonth.get()
        self.days = []

        for month, date in self.dictForMonthDays.items():
            if self.monthUser == month:
                for i in range(1, date + 1):
                    self.days.append(str(i))

        self.comboboxDay['values'] = self.days
        self.comboboxTime['values'] = self.comboboxTimeValues  # Сброс значений времени

    #Combobox Day Selected
    def getInfoComboboxDay(self, event):

        stringForDay = self.comboboxDay.get()

        for nameOfMonth, MonthNumber in self.DictForMonthes.items():

            if self.monthUser == nameOfMonth:

                if int(stringForDay) < 10:
                    self.FullConcatenate = "0" + stringForDay + "." + MonthNumber

                else:
                    self.FullConcatenate = stringForDay + "." + MonthNumber

        if self.FullConcatenate in self.dictZapisanye:

            reserved_times = self.dictZapisanye[self.FullConcatenate]
            available_times = [time for time in self.comboboxTimeValues if time not in reserved_times]
            self.comboboxTime['values'] = available_times

        else:
            self.comboboxTime['values'] = self.comboboxTimeValues

    #Запись данных в словарь
    def recordDuty(self, date, time):
        if date in self.dictZapisanye:
            if time not in self.dictZapisanye[date]:
                self.dictZapisanye[date].append(time)
        else:
            self.dictZapisanye[date] = [time]

    #Удаление всех данных
    def deleteAllData(self):

        self.textPlace.config(state='normal')

        result = tk.messagebox.askyesno('Удаление данных', 'Все данные о дежурстве будут безвозвратно удалены.\nПродолжить?')

        if result:
            self.textPlace.delete('2.0',tk.END)

        self.textPlace.config(state='disabled')

    #Открывает окно выбора сортировки
    def chooseSort(self):

        SortWindow(self.frameMain, self)

    #---------------------------------------------------------------------------------------------------------------ФУНКЦИИ


# Класс для создания окна выбора сортировки
class SortWindow():

    def __init__(self, frame, parent):

        self.parent = parent

        self.sortWindow = tk.Toplevel(frame)

        self.sortWindow.title("Выбор сортировки")
        self.sortWindow.geometry("500x500")
        self.sortWindow.geometry('+700+300')

        self.canvasSort1 = tk.Canvas(self.sortWindow, width=150, height=500, bg="#9681F0")
        self.canvasSort1.place(x=0, y=0)

        self.canvasSort2 = tk.Canvas(self.sortWindow, width=350, height=3, bg="#bcbbbd")
        self.canvasSort2.place(x=153, y=55)

        self.canvasSort3 = tk.Canvas(self.sortWindow, width=350, height=3, bg="#bcbbbd")
        self.canvasSort3.place(x=153, y=355)

        # ------------------------------------------------------------------КАРТИНКИ

        self.imageSort5 = tk.PhotoImage(file='images\\sort5.png')
        self.imageSort4 = tk.PhotoImage(file='images\\sort4.png')
        self.imageSort6 = tk.PhotoImage(file='images\\btnSort.png')
        self.imageBack2 = tk.PhotoImage(file='images\\back2.png')

        # ------------------------------------------------------------------КАРТИНКИ

        # ------------------------------------------------------------------ЛЕЙБЛЫ

        self.labelSort5 = MyLabel(self.sortWindow, image=self.imageSort5, bg="#9681F0", x=30, y=220)
        self.labelSort5.image = self.imageSort5

        self.labelSort4 = MyLabel(self.sortWindow, image=self.imageSort4, bg="#9681F0", x=30, y=100)
        self.labelSort4.image = self.imageSort4

        self.label28 = MyLabel(self.sortWindow, text="Выберите способ сортировки", x=210, y=20, font_weight='bold')

        self.label29 = MyLabel(self.sortWindow, text="(пример сортировки)", x=260, y=320, font_size=8)

        # ------------------------------------------------------------------ЛЕЙБЛЫ

        # ------------------------------------------------------------------КОМБОБОКСЫ

        self.comboboxSort = ttk.Combobox(self.sortWindow, values=['В алфавитном порядке', 'По возрастанию блоков',
                                                                  'По убыванию блоков', 'По дате'], width=30,
                                         font=('Montserrat', 10, 'bold'))
        self.comboboxSort.place(x=200, y=80)
        self.comboboxSort.bind("<<ComboboxSelected>>", self.uploadExamples)

        # ------------------------------------------------------------------КОМБОБОКСЫ

        # ------------------------------------------------------------------КНОПКИ

        self.buttonSort = MyButton(self.sortWindow, text='Отсортировать', width=15, height=2,
                                   font_size=9, font_weight='bold', x=200, y=400, command=self.funcOfSort)

        self.buttonBack = MyButton(self.sortWindow, text='Назад', width=100, height=30, image=self.imageBack2,
                                   compound=tk.RIGHT, font_size=9, font_weight='bold', command=self.closeWindow,
                                   x=350, y=400)

        # ------------------------------------------------------------------КНОПКИ

        self.textExampleSort = tk.Text(self.sortWindow, width=50, height=13, font=("Montserrat", 8, 'bold'))
        self.textExampleSort.place(x=170, y=130)

        self.textExampleSort.config(state='disabled')

    # Обновление примеров для сортировки
    def uploadExamples(self, event):

        self.textExampleSort.config(state='normal')

        self.textExampleSort.delete("0.0", tk.END)

        self.textExampleSort.insert("end", "\n\n")

        self.selectedSort = self.comboboxSort.get()

        self.checkIsEntry()

        if (self.selectedSort == "В алфавитном порядке"):

            self.textExampleSort.insert('0.0', "До")

            self.textExampleSort.insert("end", "- Жоров\n- Ловчиновский\n- Бататкин\n\n")
            self.textExampleSort.insert("end", "После\n\n")
            self.textExampleSort.insert("end", "- Бататкин\n- Жоров\n- Ловчиновский")

            self.textExampleSort.config(state='disabled')

        elif (self.selectedSort == "По возрастанию блоков"):

            self.textExampleSort.insert('0.0', "До")

            self.textExampleSort.insert("end", "- Жоров 513А\n- Бататкин 610Б\n- Ловчиновский 235A\n\n")
            self.textExampleSort.insert("end", "После\n\n")
            self.textExampleSort.insert("end", "- Ловчиновский 235А\n- Жоров 513А\n- Бататкин 610Б")

            self.textExampleSort.config(state='disabled')

        elif (self.selectedSort == "По убыванию блоков"):

            self.textExampleSort.insert('0.0', "До")

            self.textExampleSort.insert("end", "- Жоров 513А\n- Бататкин 610Б\n- Ловчиновский 235A\n\n")
            self.textExampleSort.insert("end", "После\n\n")
            self.textExampleSort.insert("end", "- Бататкин 610Б\n- Жоров 513А\n- Ловчиновский 235А")

            self.textExampleSort.config(state='disabled')

        elif (self.selectedSort == "По дате"):

            self.textExampleSort.insert('0.0', "До")

            self.textExampleSort.insert("end",
                                        "- Жоров 01.04.2025\n- Бататкин 10.01.2025\n- Ловчиновский 24.05.2025\n\n")
            self.textExampleSort.insert("end", "После\n\n")
            self.textExampleSort.insert("end", "- Бататкин 10.01.2025\n- Жоров 01.04.2025\n- Ловчиновский 24.05.2025")

            self.textExampleSort.config(state='disabled')

    # Закрытие окна
    def closeWindow(self):
        self.sortWindow.destroy()

    # Проверка не пустой ли комбобокс
    def checkIsEntry(self):

        if (self.comboboxSort.get()):

            self.comboboxSort.config(state='readonly')

            return True

        else:
            return False

    # Сортировка по возрастанию блоков
    #def sortBlocksUp(self):

    #Функция которая проверяет не одна ли строка в тексте (для сортировки нужно минимум 2)
    def checkIsNoOneLine(self):
        listsd = self.parent.textPlace.get()

        

    #Функция сортировки
    def funcOfSort(self):

        resChecking = self.checkIsEntry()

        if not resChecking:
            # parent для того, чтобы при вызове ошибки sortWindow не закрывалось
            messagebox.showerror("Ошибка", "Выберите способ сортировки!", parent=self.sortWindow)

        # Здесь сама сортировка через else (просто напиши else и там вызывай функции сортировки)

app = MyApp()
app.mainloop()
