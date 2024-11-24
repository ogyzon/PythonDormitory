'''Жоров Евгений Александрович
гр. 10701123
@vanasokolov844@gmail.com
Курсовой проект по дисциплине "Языки программирования"
Минск 2024'''

'''Учет дежурств в общежитии'''

'''Добавить год. При возможности реализовать штуку с пересылкой данных к примеру в ворд файл (типо кнопка отослать коменданту).'''

'''Таймер только на стартовое окно, так же блоки заполнения должны быть ФИО + НОМЕР; ЧИСЛО + МЕСЯЦ + ГОД (сначала число); БЛОК + КОМНАТА
 Добавить сортировку от Я до А. 5 декабря показываем готовое приложение. Данные о жильцах надо как то ренеймнуть.'''

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl import load_workbook
from docx import Document
from openpyxl.styles import Font
import time

# Класс для Label
class MyLabel(tk.Label):

    def __init__(self, master, text=None, x=0, y=0, font_size=12, font_weight="normal", justify="left", bg="#f0f0f0",
                 image=None, **kwargs):
        font = ("Montserrat", font_size, font_weight)

        super().__init__(master, text=text, font=font, justify=justify, bg=bg, image=image, **kwargs)
        self.place(x=x, y=y)


# Класс для кнопки
class MyButton(tk.Button):

    def __init__(self, master, text, width, height, fg="#000000", bg="#f0f0f0", image=None, compound=None, font_size=9,
                 font_weight="normal", command=None, x=0, y=0, **kwargs):
        font = ("Montserrat", font_size, font_weight)

        super().__init__(master, text=text, font=font, width=width, height=height, fg=fg, bg=bg, image=image,
                         compound=compound, command=command, **kwargs)
        self.place(x=x, y=y)

#Класс для отслеживания неактивности в течение 1 мин
class Inactivity():

    def __init__(self, frame):

        self.frame = frame
        self.lastActivityTime = time.time()
        self.setupBindings() #Обработчик событий
        self.checkInactivity() #Проверка на неактивость

    def resetTime(self, event):
        self.lastActivityTime = time.time()

    def setupBindings(self):
        self.frame.bind_all('<Any-KeyPress>', self.resetTime)
        self.frame.bind_all('<Any-Button>', self.resetTime)

    def checkInactivity(self):

        currentTime = time.time()
        timeDuration = currentTime - self.lastActivityTime


        if timeDuration >= 60:
            self.showWarning()

        # Проверка на неактивность каждые 1000 миллисекунд (1 секунда)
        self.frame.after(1000, self.checkInactivity)

    def showWarning(self):
        messagebox.showwarning("Бездействие", "Вы афк 60 секунд")

# Класс приложения
class MyApp(tk.Tk):

    def __init__(self):

        super().__init__()

        self.title("Учет дежурств")
        self.geometry('800x850')
        self.geometry('+500+100')
        self.resizable(width=False, height=False)

        self.inactivity = Inactivity(self)

        # Словарь для хранения записанных дат
        self.dictZapisanye = {}

        # Для перевода из слова в номер месяца
        self.DictForMonthes = {"Январь": '01', "Февраль": "02", "Март": "03", "Апрель": "04",
                               "Май": "05", "Июнь": "06", "Июль": "07", "Август": "08",
                               "Сентябрь": "09", "Октябрь": "10", "Ноябрь": "11", "Декабрь": "12"}

        self.stringDay = ""
        self.stringMonthNum = ""
        self.FullConcatenate = ""

        # Нужно
        self.concatenateString = ""
        self.monthUser = ""

        self.fullStringSort = ""

        # Для сортировки и Excel
        self.recordBlocks = []
        self.recordRooms = []
        self.recordSecondNames = []
        self.recordFirstNames = []
        self.recordDays = []
        self.recordMonthes = []
        self.recordTelNumber = []
        self.recordTimes = []
        self.recordFullStrings = []
        self.recordDates = []


        # Фрейм самого первого Splash окна
        self.frame1 = tk.Frame(self, width=800, height=800)
        self.frame1.place(x=0, y=0)

        # Фрейм второго (основного окна)
        self.frameMain = tk.Frame(self, width=800, height=800)

        # Фрейм для окна об авторе
        self.frameAuthor = tk.Frame(self, width=800, height=800)

        self.frameAbtProgramm = tk.Frame(self, width=800, height=800)

        self.canvasAbtProgramm = tk.Canvas(self.frameAbtProgramm, width=300, height=800, bg='#9681F0')
        self.canvasAbtProgramm.place(x=0, y=0)

        self.canvas1 = tk.Canvas(self.frameMain, width=750, height=670, bg='#eedcfc', borderwidth=2, relief='solid')
        self.canvas1.place(x=0, y=50)

        self.textPlace = tk.Text(self.frameMain, state='normal', font=("Montserrat", 8, 'bold'))
        self.textPlace.place(x=20, y=210, width=600, height=400)
        self.textPlace.insert(tk.END, f"Фамилия\t\t  Имя\t\tБлок\t\tTелефон\t\t   Дата\t\tВремя\n\n")
        self.textPlace.config(state='disabled')
        self.textPlace.bind('<ButtonRelease-1>', self.onTextClick)

        # Лейблы с текстом Splash окна
        self.label1 = MyLabel(self.frame1, text='Белорусский национальный технический университет', x=195, y=20)

        self.label2 = MyLabel(self.frame1, text='Факультет информационных технологий и робототехники', x=175, y=50)

        self.label3 = MyLabel(self.frame1, text='Кафедра программного обеспечения информационных систем и технологий',
                              x=120, y=80)

        self.label4 = MyLabel(self.frame1, text='Курсовой проект', font_size=16, font_weight='bold', x=300, y=160)

        self.label5 = MyLabel(self.frame1, text='По дисциплине "Языки программирования"', font_size=16,
                              font_weight='bold', x=160, y=200)

        self.label6 = MyLabel(self.frame1, text='Выполнил: студент группы 10701123', font_size=12, font_weight='bold',
                              x=340, y=360)

        self.label7 = MyLabel(self.frame1, text='Жоров Евгений Александрович', font_size=12, font_weight='bold', x=340,
                              y=390)

        self.label8 = MyLabel(self.frame1, text='Преподаватель: к.ф.-м.н., доц.', font_size=12, font_weight='bold',
                              x=340, y=450)

        self.label9 = MyLabel(self.frame1, text='Сидорик Валерий Владимирович', font_size=12, font_weight='bold', x=340,
                              y=480)

        self.label10 = MyLabel(self.frame1, text='Минск 2024', font_size=12, font_weight='bold', x=350, y=630)

        self.label11 = MyLabel(self.frame1, text='Учет дежурств в общежитии', font_size=16, font_weight='bold', x=235,
                               y=240)

        # Лейблы с текстом окна об авторе
        self.label12 = MyLabel(self.frameAuthor, text='Автор', font_size=14, font_weight='bold', x=350, y=470)

        self.label13 = MyLabel(self.frameAuthor, text='Студент группы 10701123', font_size=12, font_weight='bold',
                               x=280, y=500)

        self.label14 = MyLabel(self.frameAuthor, text='Жоров Евгений Александрович', font_size=12, font_weight='bold',
                               x=260, y=530)

        self.label15 = MyLabel(self.frameAuthor, text='vanasokolov844@gmail.com', font_size=12, font_weight='bold',
                               x=270, y=560)

        # Лейблы для текста окна о программе
        self.label16 = MyLabel(self.frameAbtProgramm, text='Учет дежурств в общежитии', font_size=18,
                               font_weight='bold', x=380, y=40)

        self.label17 = MyLabel(self.frameAbtProgramm, text='Программа позволяет:', font_size=13, font_weight='bold',
                               x=440, y=90)

        self.label18 = MyLabel(self.frameAbtProgramm, text='1. Записывать ФИО и номер телефона дежурного\n'
                                                           '2. Сохранять записи в файл\n'
                                                           '3. Просматривать результат в главном окне\n'
                                                           '4. Удалять данные о дежурстве\n'
                                                           '5. Записывать и считывать данные с Excel', font_size=11,
                               justify='left', bg='#eddcfc', x=350, y=130)

        self.label19 = MyLabel(self.frameAbtProgramm, text='Версия: 1.0.0.2024', font_size=10, x=350, y=710)

        # Лейблы с текстом главного окна
        self.label20 = MyLabel(self.frameMain, text="Данные о жильцах", font_size=11, font_weight='bold', bg='#eedcfc',
                               x=300, y=15)

        self.label21 = MyLabel(self.frameMain, text="Фамилия", font_size=10, bg='#eedcfc', x=20, y=70)

        self.label22 = MyLabel(self.frameMain, text="Имя", font_size=10, bg='#eedcfc', x=150, y=70)

        self.label23 = MyLabel(self.frameMain, text="Блок", font_size=10, bg='#eedcfc', x = 450, y = 70)

        self.label24 = MyLabel(self.frameMain, text="Номер телефона", font_size=10, bg='#eedcfc', x=280, y=70)

        self.label24 = MyLabel(self.frameMain, text="Время", font_size=10, bg='#eedcfc', x=280, y=140)

        self.label25 = MyLabel(self.frameMain, text="Месяц", font_size=10, bg='#eedcfc', x=150, y=140)

        self.label26 = MyLabel(self.frameMain, text="День", font_size=10, bg='#eedcfc', x=20, y=140)

        self.label27 = MyLabel(self.frameMain, text="Комната", font_size=10, bg='#eedcfc',  x = 540, y = 70)

        self.label28 = MyLabel(self.frameMain, text="Год", font_size=10, bg='#eedcfc', x=400, y=140)

        # Entry поля на главном окне
        # Ввод фамилии
        self.entry1 = tk.Entry(self.frameMain, width=15)
        self.entry1.place(x=20, y=100)

        # Ввод имени
        self.entry2 = tk.Entry(self.frameMain, width=15)
        self.entry2.place(x=150, y=100)

        # Ввод блока
        self.entry3 = tk.Entry(self.frameMain, width=10)
        self.entry3.place(x = 450, y = 100)

        # Ввод номера тф
        self.entry4 = tk.Entry(self.frameMain, width=18)
        self.entry4.place(x=280, y=100)

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

        # Ввод месяца
        self.comboboxMonth = ttk.Combobox(self.frameMain, values=['Январь', 'Февраль', 'Март',
                                                                  'Апрель', 'Май', 'Июнь',
                                                                  'Июль', 'Август', 'Сентябрь',
                                                                  'Октябрь', 'Ноябрь', 'Декабрь'], width=12,
                                          state="readonly")
        self.comboboxMonth.place(x=150, y=170)
        self.comboboxMonth.bind("<<ComboboxSelected>>", self.UpdateDays)

        # Ввод дня
        self.comboboxDay = ttk.Combobox(self.frameMain, width=12, state="readonly")
        self.comboboxDay.place(x=20, y=170)

        self.comboboxDay.bind("<<ComboboxSelected>>", self.getInfoComboboxDay)

        self.comboboxTimeValues = ['8.00 - 10.30', '10.30 - 13.00', '13.00 - 15.30', '15.30 - 18.00', '18.00 - 20.00']

        # Ввод времени
        self.comboboxTime = ttk.Combobox(self.frameMain, values=self.comboboxTimeValues, width=11, state="readonly")
        self.comboboxTime.place(x=280, y=170)

        self.UpdateDays(None)

        # Ввод комнаты
        self.comboboxRoomInBlock = ttk.Combobox(self.frameMain, values=['А', 'Б'], width=11, state='readonly')
        self.comboboxRoomInBlock.place(x = 540, y = 100)

        #Ввод года
        self.comboboxYear = ttk.Combobox(self.frameMain, values = ['2025'], width=11, state='readonly')
        self.comboboxYear.place(x = 400, y = 170)

        # ---------------------------------------------------------------------------------------КАРТИНКИ
        self.imageDormitory = tk.PhotoImage(file='images//iconDormitory.png')
        self.labelForDormitory = MyLabel(self.frame1, image=self.imageDormitory, x=160, y=370)
        self.labelForDormitory.image = self.imageDormitory
        # self.labelForDormitory.place(x = 160, y = 370)

        self.imageMe = tk.PhotoImage(file='images//Me.png')
        self.resizedImageMe = self.imageMe.subsample(3)
        self.labelForMe = MyLabel(self.frameAuthor, image=self.resizedImageMe, x=230, y=30)
        self.labelForMe.image = self.resizedImageMe

        self.imageForAbtProgramm1 = tk.PhotoImage(file='images//dormitory2.png')
        self.labelForAbtProgramm1 = MyLabel(self.frameAbtProgramm, image=self.imageForAbtProgramm1, bg="#9681F0", x=30,
                                            y=230)
        self.labelForAbtProgramm1.image = self.imageForAbtProgramm1

        self.imageForAbtProgramm2 = tk.PhotoImage(file='images//room.png')
        self.labelForAbtProgramm2 = MyLabel(self.frameAbtProgramm, image=self.imageForAbtProgramm2, bg='#9681F0', x=150,
                                            y=235)
        self.labelForAbtProgramm2.image = self.imageForAbtProgramm2

        self.imageForAbtProgramm3 = tk.PhotoImage(file='images//file.png')
        self.labelForAbtProgramm3 = MyLabel(self.frameAbtProgramm, image=self.imageForAbtProgramm3, bg='#9681F0', x=30,
                                            y=350)
        self.labelForAbtProgramm3.image = self.imageForAbtProgramm3

        self.imageForAbtProgramm4 = tk.PhotoImage(file='images//excel.png')
        self.labelForAbtProgramm4 = MyLabel(self.frameAbtProgramm, image=self.imageForAbtProgramm4, bg='#9681F0', x=150,
                                            y=350)
        self.labelForAbtProgramm4.image = self.imageForAbtProgramm4

        self.imageForAbtProgramm5 = tk.PhotoImage(file='images//human2.png')
        self.labelForAbtProgramm5 = MyLabel(self.frameAbtProgramm, image=self.imageForAbtProgramm5, bg='#9681F0', x=30,
                                            y=460)
        self.labelForAbtProgramm5.image = self.imageForAbtProgramm5

        self.imageForAbtProgramm6 = tk.PhotoImage(file='images//telephone.png')
        self.labelForAbtProgramm6 = MyLabel(self.frameAbtProgramm, image=self.imageForAbtProgramm6, bg='#9681F0', x=150,
                                            y=460)
        self.labelForAbtProgramm6.image = self.imageForAbtProgramm6

        # Картинка для кнопки выйти
        self.imageExit = tk.PhotoImage(file='images//exitMain.png')

        # Картинка для кнопки назад
        self.imageBack = tk.PhotoImage(file='images//backMain.png')

        # Картинка для кнопки об авторе
        self.imageHuman = tk.PhotoImage(file='images//human.png')

        # Картинка для кнопки о программе
        self.imageAbtProgramm = tk.PhotoImage(file='images//about.png')

        # Картинка для кнопки на главную
        self.imageToTheMain = tk.PhotoImage(file='images//home.png')

        # Картинка для кнопки записаться
        self.imageWrite = tk.PhotoImage(file='images//write.png')

        # Картинка для кнопки очистки
        self.imageClear = tk.PhotoImage(file='images//clear.png')

        # Картинка для кнопки удалить данные
        self.imageDelete = tk.PhotoImage(file='images//delete.png')

        # Картинка для кнопки выбора сортировки
        self.imageSort = tk.PhotoImage(file='images//sort.png')

        # Картинка для кнопки записи в Excel
        self.imageExcelBtn = tk.PhotoImage(file = 'images//excel2Btn.png')

        # Картинка для кнопки записи в Excel
        self.imageClearExcelBtn = tk.PhotoImage(file='images//clearFile.png')

        #Картинка для кнопки отменить изменения
        self.imageCancel = tk.PhotoImage(file = 'images//cancel.png')

        #Картинка для кнопки сохранить изменения
        self.imageSaveChanges = tk.PhotoImage(file = 'images//saveChanges.png')

        # Картинка для кнопки отправить коменданту
        self.imageWord = tk.PhotoImage(file='images//word.png')
        # ---------------------------------------------------------------------------------КАРТИНКИ

        # ---------------------------------------------------------------------------------КНОПКИ

        # Кнопка далее
        self.buttonStart = MyButton(self.frame1, text='Далее', fg='white', width=30, height=3,
                                    font_size=9, font_weight='bold', bg='#8251FE', command=self.openSecondWindow, x=150,
                                    y=700)

        # Кнопка выход
        if self.frameMain:
            self.buttonExit = MyButton(self.frameMain, text='Выход', width=100, height=30, image=self.imageExit,
                                       compound=tk.RIGHT, font_size=9, font_weight='bold', command=self.exitApp, x=525,
                                       y=760)

        if self.frame1:
            self.buttonExit = MyButton(self.frame1, text='Выход', width=200, height=45, image=self.imageExit,
                                       compound=tk.RIGHT, font_size=9, font_weight='bold', command=self.exitApp, x=410,
                                       y=700)

        # Кнопка назад УНИВЕРСАЛЬНАЯ
        if self.frameMain:
            self.buttonBack = MyButton(self.frameMain, text='Назад', width=100, height=30, image=self.imageBack,
                                       compound=tk.RIGHT, font_size=9, font_weight='bold',
                                       command=self.backFromMainto1st, x=645, y=760)

        if self.frameAbtProgramm:
            self.buttonBack = MyButton(self.frameAbtProgramm, text='Назад', width=100, height=30, image=self.imageBack,
                                       compound=tk.RIGHT, font_size=9, font_weight='bold',
                                       command=self.backFromAbtProgrammToMain, x=645, y=700)

        # Кнопка об авторе (добавить потом это окно)
        self.buttonAuthor = MyButton(self.frameMain, text='Об авторе', width=100, height=30, image=self.imageHuman,
                                     compound=tk.RIGHT, font_size=9, font_weight='bold',
                                     command=self.openAuthorFromMain, x=200, y=760)

        # Кнопка о программе
        self.buttonAbtProgramm = MyButton(self.frameMain, text='О программе', width=120, height=30,
                                          image=self.imageAbtProgramm,
                                          compound=tk.RIGHT, font_size=9, font_weight='bold',
                                          command=self.openAbtProgramm, x=60, y=760)

        # Кнопка на главную
        self.buttonToTheMain = MyButton(self.frameAuthor, text='На главную', width=250, height=30,
                                        image=self.imageToTheMain,
                                        compound=tk.RIGHT, font_size=9, font_weight='bold',
                                        command=self.backFromAuthorToMain, x=265, y=700)
        # Кнопка записаться
        self.buttonZapisat = MyButton(self.frameMain, text='Записаться', width=80, height=60, image=self.imageWrite,
                                      compound=tk.TOP, font_size=9, font_weight='bold', command=self.GetResults, x=645,
                                      y=100)

        # Кнопка очистить ввод
        self.buttonClear = MyButton(self.frameMain, text="Очистить\nввод", width=80, height=60, image=self.imageClear,
                                    compound=tk.TOP, font_size=9, font_weight='bold', command=self.clearEntry, x=645,
                                    y=210)

        # Кнопка удалить данные
        self.buttonDeleteData = MyButton(self.frameMain, text='Удалить\nданные', width=80, height=60,
                                         image=self.imageDelete,
                                         compound=tk.TOP, font_size=9, font_weight='bold', command=self.deleteAllData,
                                         x=645, y=320)

        # Кнопка выбрать сортировку
        self.buttonChooseSort = MyButton(self.frameMain, text='Выбрать\nсортировку', width=80, height=60,
                                         image=self.imageSort,
                                         compound=tk.TOP, font_size=9, font_weight='bold', command=self.chooseSort,
                                         x=645, y=425)

        # Кнопка записи в Excel
        self.buttonToExcel = MyButton(self.frameMain, text='Сохранить\nв Excel', width=80, height=60,
                                         image=self.imageExcelBtn,
                                         compound=tk.TOP, font_size=9, font_weight='bold', command=self.infoToExcel,
                                         x=645, y=530)

        # Кнопка очистки Excel файла по умолчанию
        self.buttonDeleteExcel = MyButton(self.frameMain, text='Очистить\nфайл', width=80, height=60,
                                      image=self.imageClearExcelBtn,
                                      compound=tk.TOP, font_size=9, font_weight='bold', command = self.clearDataFromExcel,
                                      x=645, y=635)

        # Кнопка сохранения в Docx
        self.buttonToDocx = MyButton(self.frameMain, text='Отправить\nкоменданту', width=80, height=60,
                                          image=self.imageWord,
                                          compound=tk.TOP, font_size=9, font_weight='bold', command = self.infoToDocx,
                                          x=345, y=635)

        #Кнопка сохранения изменений
        self.buttonSaveChanges = MyButton(self.frameMain, text = "Сохранить\nизменения", width=80, height=60,
                                         image=self.imageSaveChanges,
                                         compound=tk.TOP, font_size=9, font_weight='bold', command=self.saveAndPlace,
                                         x=645, y=100)
        self.buttonSaveChanges.place_forget()

        #Кнопка отмена изменений
        self.buttonCancel = MyButton(self.frameMain, text="Отмена", width=80, height=60,
                                          image=self.imageCancel,
                                          compound=tk.TOP, font_size=9, font_weight='bold', command=self.cancelChange,
                                          x=645, y=210)
        self.buttonCancel.place_forget()


        #--------------------------------------------------------------------------------------------------------------МЕНЮ КНОПКИ

        self.buttonMenuFile = tk.Menubutton(self.frameMain, text = 'Файл', relief=tk.RAISED, width = 5)
        self.buttonMenuFile.place(x = 0, y = 0)

        self.fileMenu = tk.Menu(self.buttonMenuFile, tearoff=0)
        self.buttonMenuFile.config(menu = self.fileMenu)

        #Открытие файла (чтение из файла)
        self.fileMenu.add_command(label = "Открыть",command=self.openExcelFile) #Сюда добавь команду

        self.fileMenu.add_command(label = "Сохранить как", command = self.saveAsExcelFile)

        #Помощь
        self.buttonMenuHelp = tk.Menubutton(self.frameMain, text = "Помощь", relief = tk.RAISED, width = 7)
        self.buttonMenuHelp.place(x = 43, y = 0)

        self.helpMenu = tk.Menu(self.buttonMenuHelp, tearoff=0)
        self.buttonMenuHelp.config(menu = self.helpMenu)

        self.helpMenu.add_command(label = "Использование", command = self.openHelpWindow)

        # --------------------------------------------------------------------------------------------------------------МЕНЮ КНОПКИ

    # -----------------------------------------------------------------------------------------------------------------КНОПКИ

    # -----------------------------------------------------------------------------------------------------------------ФУНКЦИИ
    # Закрытие окна
    def exitApp(self):
        self.destroy()

    # Открытие второго окна (забиндить и добавить эл-ты второго окна в фрейм 2)
    def openSecondWindow(self):
        self.frame1.pack_forget()
        self.frameMain.pack()

    # Возвращение со второго окна на первое(с основного в Splash)
    def backFromMainto1st(self):
        self.frameMain.pack_forget()
        self.frame1.pack()

    # Открытие окна об авторе с главного окна
    def openAuthorFromMain(self):
        self.frameMain.pack_forget()
        self.frameAuthor.pack()

    # Возвращение с окна об авторе на главную
    def backFromAuthorToMain(self):
        self.frameAuthor.pack_forget()
        self.frameMain.pack()

    def openAbtProgramm(self):
        self.frameMain.pack_forget()
        self.frameAbtProgramm.pack()

    def backFromAbtProgrammToMain(self):
        self.frameAbtProgramm.pack_forget()
        self.frameMain.pack()

    def openHelpWindow(self):

        self.helpWindow = HelpWindow(self.frameMain, self)

    # Очистка всех полей при нажатии кнопки очистить ввод
    def clearEntry(self):

        self.comboboxMonth.config(state="normal")
        self.comboboxDay.config(state="normal")
        self.comboboxTime.config(state="normal")
        self.comboboxRoomInBlock.config(state="normal")
        self.comboboxYear.config(state="normal")

        self.entry1.delete(0, tk.END)
        self.entry2.delete(0, tk.END)
        self.entry3.delete(0, tk.END)
        self.entry4.delete(0, tk.END)
        self.comboboxMonth.delete(0, tk.END)
        self.comboboxDay.delete(0, tk.END)
        self.comboboxTime.delete(0, tk.END)
        self.comboboxRoomInBlock.delete(0, tk.END)
        self.comboboxYear.delete(0, tk.END)

        self.comboboxMonth.config(state="readonly")
        self.comboboxDay.config(state="readonly")
        self.comboboxTime.config(state="readonly")
        self.comboboxRoomInBlock.config(state="readonly")
        self.comboboxYear.config(state="readonly")

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

        self.monthForSort = 0

        # Цикл для вывода даты и записи занятых дат в словарь
        for nameOfMonth,MonthNumber in self.DictForMonthes.items():
            if self.Month == nameOfMonth:
                if int(self.Day) < 10:
                    self.concatenateString = "0" + self.Day + "." + MonthNumber
                else:
                    self.concatenateString = self.Day + "." + MonthNumber

                self.monthForSort = int(MonthNumber)

                # Запись занятой даты в словарь
                self.recordDuty(self.concatenateString, self.Time)  # Вызов функции записи данных

        # Для сортировки
        self.fullStringSort = (f"{(str(self.lastName)).capitalize()}\t\t  {(str(self.firstName)).capitalize()}\t\t"
                               f"{str(self.Block) + str(self.Room)}\t\t{str(self.TelNumber)}\t\t   {str(self.concatenateString)}\t\t{str(self.Time)}\n")

        self.textPlace.insert(tk.END,
                              f"{(str(self.lastName)).capitalize()}\t\t  {(str(self.firstName)).capitalize()}\t\t"
                              f"{str(self.Block) + str(self.Room)}\t\t{str(self.TelNumber)}\t\t   {str(self.concatenateString)}\t\t{str(self.Time)}\n")

        self.clearEntry()

        # Для сортировки и записи данных для Excel
        self.recordData()

        self.textPlace.config(state='disabled')

        print(self.recordFirstNames)
        print(self.recordDates)
        print(self.recordRooms)
        print(self.recordSecondNames)
        print(self.recordBlocks)

    def UpdateDays(self, event):
        self.monthUser = self.comboboxMonth.get()
        current_day = self.comboboxDay.get()

        # Изначально устанавливаем 31 день
        self.days = [str(i) for i in range(1, 32)]

        if self.monthUser:
            days_in_month = self.dictForMonthDays[self.monthUser]
            self.days = [str(i) for i in range(1, days_in_month + 1)]

            if current_day and int(current_day) > days_in_month:
                self.comboboxDay.set("")
            elif current_day:
                self.comboboxDay.set(current_day)

        self.comboboxDay['values'] = self.days
        self.updateAvailableTimes()

    def getInfoComboboxDay(self, event):
        self.updateAvailableTimes()

    def updateAvailableTimes(self):
        stringForDay = self.comboboxDay.get()
        stringForMonth = self.comboboxMonth.get()

        if not stringForDay:  # Проверка на пустую строку
            self.comboboxTime['values'] = self.comboboxTimeValues
            return

        self.FullConcatenate = None
        for nameOfMonth, MonthNumber in self.DictForMonthes.items():
            if stringForMonth == nameOfMonth:
                if int(stringForDay) < 10:
                    self.FullConcatenate = "0" + stringForDay + "." + MonthNumber
                else:
                    self.FullConcatenate = stringForDay + "." + MonthNumber

        if self.FullConcatenate:
            if self.FullConcatenate in self.dictZapisanye:
                reserved_times = self.dictZapisanye[self.FullConcatenate]
                available_times = [time for time in self.comboboxTimeValues if time not in reserved_times]
                self.comboboxTime['values'] = available_times
            else:
                self.comboboxTime['values'] = self.comboboxTimeValues

    def recordDuty(self, date, time):
        if date in self.dictZapisanye:
            if time not in self.dictZapisanye[date]:
                self.dictZapisanye[date].append(time)
        else:
            self.dictZapisanye[date] = [time]

    # Запись данных в списки (типо для блоков к примеру свой отдельный список и тд) (для реализации сортировки)
    def recordData(self):

        self.recordBlocks.append(self.Block)
        self.recordSecondNames.append(self.lastName)
        self.recordFullStrings.append(self.fullStringSort)
        self.recordDays.append(int(self.Day))
        self.recordMonthes.append(self.monthForSort)
        self.recordRooms.append(self.Room)
        self.recordFirstNames.append(self.firstName)
        self.recordTelNumber.append(self.TelNumber)
        self.recordDates.append(self.concatenateString)
        self.recordTimes.append(self.Time)

    # Удаление всех данных
    def deleteAllData(self):

        self.textPlace.config(state='normal')

        result = tk.messagebox.askyesno('Удаление данных',
                                        'Все данные о дежурстве будут безвозвратно удалены.\nПродолжить?')

        if result:
            self.textPlace.delete('3.0', tk.END)
            self.textPlace.insert('2.0', '\n')

        self.textPlace.config(state='disabled')

        self.recordBlocks.clear()
        self.recordSecondNames.clear()
        self.recordFullStrings.clear()
        self.recordTelNumber.clear()
        self.recordRooms.clear()
        self.recordFirstNames.clear()
        self.recordDates.clear()
        self.recordTimes.clear()

        self.dictZapisanye.clear()

    # Открывает окно выбора сортировки
    def chooseSort(self):

        SortWindow(self.frameMain, self)

    #Функция отмены изменений
    def cancelChange(self):
        self.clearEntry()

        # Скрытие кнопок сохранения изменений и отмены
        self.buttonSaveChanges.place_forget()
        self.buttonCancel.place_forget()

        # Возвращение кнопок на исходные позиции
        self.buttonZapisat.place(x=645, y=100)
        self.buttonToExcel.place(x=645, y=530)
        self.buttonDeleteData.place(x=645, y=320)
        self.buttonChooseSort.place(x=645, y=425)
        self.buttonDeleteExcel.place(x=645, y=635)
        self.buttonClear.place(x=645, y=210)

        # Удаление стрелочки и подсветки с текущей выбранной строки
        if self.selected_index is not None:
            self.removeArrowAndHighlight(self.selected_index)
            self.selected_index = None

    #Функция убирающая кнопки при нажатии на редактирование строки
    def switchButtons(self):

        self.buttonZapisat.place_forget()
        self.buttonToExcel.place_forget()
        self.buttonDeleteData.place_forget()
        self.buttonChooseSort.place_forget()
        self.buttonDeleteExcel.place_forget()
        self.buttonClear.place_forget()

        self.buttonSaveChanges.place(x = 645, y = 100)
        self.buttonCancel.place(x = 645, y = 210)

    #Функция получения данных при нажатии на текст
    def onTextClick(self, event):
        # Сохраняем индекс предыдущей выбранной строки
        if hasattr(self, 'selected_index') and self.selected_index is not None:
            previous_index = self.selected_index
        else:
            previous_index = None

        self.textPlace.config(state='normal')
        index = self.textPlace.index("@%s,%s" % (event.x, event.y))
        line_content = self.textPlace.get("%s linestart" % index, "%s lineend" % index).strip()
        self.textPlace.config(state='disabled')

        print(f"Index: {index}, Line Content: {line_content.strip()}")

        parts = line_content.split()
        if len(parts) >= 7:
            try:
                row_index = int(index.split('.')[0]) - 3  # Корректировка для учета первых двух строк
                print(f"Row Index: {row_index}")
                print(
                    f"List Lengths: {len(self.recordSecondNames)}, {len(self.recordFirstNames)}, {len(self.recordBlocks)}, {len(self.recordRooms)}, {len(self.recordTelNumber)}, {len(self.recordDays)}, {len(self.recordMonthes)}, {len(self.recordDates)}, {len(self.recordTimes)}")

                if row_index >= 0 and row_index < len(self.recordSecondNames):
                    lastName = self.recordSecondNames[row_index].strip()
                    firstName = self.recordFirstNames[row_index].strip()
                    block = self.recordBlocks[row_index].strip()
                    room = self.recordRooms[row_index].strip()
                    telNumber = self.recordTelNumber[row_index].strip()
                    day = str(self.recordDays[row_index]) if row_index < len(
                        self.recordDays) else '01'  # Проверка на наличие дня
                    month = self.recordMonthes[row_index] if row_index < len(
                        self.recordMonthes) else 1  # Проверка на наличие месяца
                    date_str = self.recordDates[row_index].strip()
                    time = self.recordTimes[row_index].strip()

                    # Форматируем месяц как строку с двумя цифрами
                    month_str = f"{month:02}"
                    print(f"Day: {day}, Month: {month_str}, Date: {date_str}, Time: {time}")

                    self.comboboxDay.set(day)
                    month_name = next((key for key, value in self.DictForMonthes.items() if value == month_str), None)
                    if month_name:
                        self.comboboxMonth.set(month_name)
                    else:
                        print(f"Не удалось найти соответствие для месяца: {month_str}")

                    self.comboboxTime.set(time)
                    self.updateAvailableTimes()

                    self.entry1.config(state='normal')
                    self.entry1.delete(0, tk.END)
                    self.entry1.insert(0, lastName)

                    self.entry2.config(state='normal')
                    self.entry2.delete(0, tk.END)
                    self.entry2.insert(0, firstName)

                    self.entry3.config(state='normal')
                    self.entry3.delete(0, tk.END)
                    self.entry3.insert(0, block)

                    self.comboboxRoomInBlock.config(state='normal')
                    self.comboboxRoomInBlock.set(room)

                    self.entry4.config(state='normal')
                    self.entry4.delete(0, tk.END)
                    self.entry4.insert(0, telNumber)

                    # Удаление стрелочки и подсветки с предыдущей строки
                    if previous_index is not None and previous_index != index:
                        previous_line_content = self.textPlace.get("%s linestart" % previous_index,
                                                                   "%s lineend" % previous_index).strip()
                        if previous_line_content.endswith("←"):
                            updated_previous_line = previous_line_content[:-1].rstrip()
                            self.textPlace.config(state='normal')
                            self.textPlace.delete("%s linestart" % previous_index, "%s lineend" % previous_index)
                            self.textPlace.insert("%s linestart" % previous_index, updated_previous_line)
                            self.textPlace.tag_remove("highlight", "%s linestart" % previous_index,
                                                      "%s lineend" % previous_index)
                            self.textPlace.config(state='disabled')

                    # Добавление стрелочки "←" к новой выбранной строке и подсветка
                    if not line_content.endswith("←"):
                        updated_line = line_content + " ←"
                        self.textPlace.config(state='normal')
                        self.textPlace.delete("%s linestart" % index, "%s lineend" % index)
                        self.textPlace.insert("%s linestart" % index, updated_line)
                    self.textPlace.tag_add("highlight", "%s linestart" % index, "%s lineend" % index)
                    self.textPlace.tag_configure("highlight", background="#dca2f5")
                    self.textPlace.config(state='disabled')

                    # Обновление индекса выбранной строки
                    self.selected_index = index
                    self.switchButtons()
                else:
                    print("Индекс строки выходит за пределы списка")
            except ValueError:
                print("Не удалось преобразовать индекс строки в целое число")

    def removeDuty(self, date, time):
        if date in self.dictZapisanye:
            if time in self.dictZapisanye[date]:
                self.dictZapisanye[date].remove(time)
                if not self.dictZapisanye[date]:  # Удаление даты, если нет занятых временных слотов
                    del self.dictZapisanye[date]

    #Функция сохранения изменений
    def saveChanges(self):
        # Получаем обновленные данные из полей
        lastName = self.entry1.get().strip()
        firstName = self.entry2.get().strip()
        block = self.entry3.get().strip()
        room = self.comboboxRoomInBlock.get().strip()
        telNumber = self.entry4.get().strip()
        day = self.comboboxDay.get().strip().zfill(2)  # Сохраняем ведущий ноль
        month = next(value for key, value in self.DictForMonthes.items() if self.comboboxMonth.get() == key)
        time = self.comboboxTime.get().strip()
        date_str = f"{day}.{month}"

        # Проверка на корректность ввода фамилии
        if any(symbol.isdigit() or symbol in "_-+=!@#$%*^()&?~/., " for symbol in lastName):
            messagebox.showerror("Некорректный ввод", "Фамилия введена некорректно!")
            self.entry1.delete(0, tk.END)
            self.removeArrowAndHighlight(self.selected_index)
            return

        # Проверка на корректность ввода имени
        if any(symbol.isdigit() or symbol in "_-+=!@#$%*^()&?~/., " for symbol in firstName):
            messagebox.showerror("Некорректный ввод", "Имя введено некорректно!")
            self.entry2.delete(0, tk.END)
            self.removeArrowAndHighlight(self.selected_index)
            return

        # Проверка на корректность ввода блока
        if any(not symbol.isdigit() for symbol in block):
            messagebox.showerror("Некорректный ввод", "Блок введен некорректно!")
            self.entry3.delete(0, tk.END)
            self.removeArrowAndHighlight(self.selected_index)
            return

        # Проверка на корректность ввода номера телефона
        if any(not symbol.isdigit() and symbol not in "+" for symbol in telNumber):
            messagebox.showerror("Некорректный ввод", "Номер введен некорректно!")
            self.entry4.delete(0, tk.END)
            self.removeArrowAndHighlight(self.selected_index)
            return

        # Проверка все ли поля введены
        if not all([lastName, firstName, block, telNumber, month, day, time, room]):
            messagebox.showerror("Ошибка", "Введите все данные!")
            self.removeArrowAndHighlight(self.selected_index)
            return

        # Формирование обновленной строки
        updated_line = f"{lastName.capitalize()}\t\t  {firstName.capitalize()}\t\t{block}{room}\t\t{telNumber}\t\t   {date_str}\t\t{time} ←"

        # Определение текущего индекса строки
        if self.selected_index:
            index = self.selected_index
            row_index = int(index.split('.')[0]) - 3  # Корректировка для учета первых двух строк

            # Получение старого времени и даты
            old_date_str = self.recordDates[row_index]
            old_time = self.recordTimes[row_index]

            # Удаление старого времени из словаря dictZapisanye
            self.removeDuty(old_date_str, old_time)

            # Удаление старой строки и вставка обновленной строки
            self.textPlace.config(state='normal')
            self.textPlace.delete("%s linestart" % index, "%s lineend" % index)
            self.textPlace.insert("%s linestart" % index, updated_line)
            self.textPlace.config(state='disabled')

            # Обновление данных в соответствующих списках
            if row_index >= 0 and row_index < len(self.recordSecondNames):
                self.recordSecondNames[row_index] = lastName
                self.recordFirstNames[row_index] = firstName
                self.recordBlocks[row_index] = block
                self.recordRooms[row_index] = room
                self.recordTelNumber[row_index] = telNumber
                self.recordDays[row_index] = int(day)
                self.recordMonthes[row_index] = month
                self.recordDates[row_index] = date_str
                self.recordTimes[row_index] = time
                self.recordFullStrings[row_index] = updated_line
            else:
                print("Индекс строки выходит за пределы списка")

            # Запись нового времени в словарь dictZapisanye
            self.recordDuty(date_str, time)

            # Удаление стрелки и подсветки после сохранения данных
            self.removeArrowAndHighlight(index)
            self.selected_index = None

    #Функция убирающая подсветку текста и стрелочку
    def removeArrowAndHighlight(self, index):
        if index is not None:
            line_content = self.textPlace.get("%s linestart" % index, "%s lineend" % index).strip()
            if line_content.endswith("←"):
                updated_line = line_content[:-1].rstrip()
                self.textPlace.config(state='normal')
                self.textPlace.delete("%s linestart" % index, "%s lineend" % index)
                self.textPlace.insert("%s linestart" % index, updated_line)
                self.textPlace.tag_remove("highlight", "%s linestart" % index, "%s lineend" % index)
                self.textPlace.config(state='disabled')

    #Функция вызывающая функцию сохранения изменений и позиционирующая кнопки
    def saveAndPlace(self):

        self.saveChanges()
        self.cancelChange()


    #Функция записывающая данные в Excel
    def infoToExcel(self):

        result = tk.messagebox.askyesno("Изменить путь к файлу", "Хотите ли вы выбрать другой файл для записи?")

        if result:

            filePath = tk.filedialog.askopenfilename(title="Выберите файл", filetypes = [("Excel файлы", "*.xls"), ("Excel файлы", ".xlsx")])

            if filePath:

                self.excelObject = WorkExcel(self, filePath)

                self.excelObject.infoInFile(filePath)

        if not result:

            filePath = "excelFiles//Book1.xlsx"

            if filePath:

                self.excelObject = WorkExcel(self, filePath)

                self.excelObject.infoInFile(filePath)

    #Функция очистки файла по умолчанию
    def clearDataFromExcel(self):

        result = tk.messagebox.askyesno("Очистить файл", "Все данные из Excel файла по умолчанию будут удалены. Продолжить?")

        if result:

            filePath = "excelFiles//Book1.xlsx"

            self.excelObject = WorkExcel(self, filePath)

        else:

            return

    #Открытие файла(считывание)
    def openExcelFile(self):
        filePath = tk.filedialog.askopenfilename(title="Выберите файл",
                                                 filetypes=[("Excel файлы", "*.xls"), ("Excel файлы", "*.xlsx")])

        self.wb = load_workbook(filePath)
        self.activeList = self.wb.active

        expectedHeaders = ['Фамилия', 'Имя', 'Блок', 'Комната', 'Номер', 'Дата', 'Время']
        actualHeaders = [self.activeList.cell(row=1, column=i + 1).value for i in range(len(expectedHeaders))]

        if actualHeaders != expectedHeaders:
            tk.messagebox.showerror("Ошибка файла", "Вы открыли файл с посторонними данными")
            return

        self.recordBlocks.clear()
        self.recordSecondNames.clear()
        self.recordFullStrings.clear()
        self.recordTelNumber.clear()
        self.recordRooms.clear()
        self.recordFirstNames.clear()
        self.recordDates.clear()
        self.recordTimes.clear()
        self.recordDays.clear()  # Добавим очистку списка recordDays
        self.recordMonthes.clear()  # Добавим очистку списка recordMonthes
        self.dictZapisanye.clear()

        for row in self.activeList.iter_rows(min_row=2, max_row=self.activeList.max_row, min_col=1,
                                             max_col=self.activeList.max_column, values_only=True):
            self.recordSecondNames.append(row[0])
            self.recordFirstNames.append(row[1])
            self.recordBlocks.append(row[2])
            self.recordRooms.append(row[3])
            self.recordTelNumber.append(row[4])
            self.recordDates.append(row[5])
            self.recordTimes.append(row[6])

            # Добавим извлечение дня и месяца из даты и занесем в соответствующие списки
            day, month = map(int, row[5].split('.'))
            self.recordDays.append(day)
            self.recordMonthes.append(month)

            # Формирование полной строки и добавление в recordFullStrings
            full_string = f"{row[0].capitalize()}\t\t  {row[1].capitalize()}\t\t{row[2]}{row[3]}\t\t{row[4]}\t\t   {row[5]}\t\t{row[6]}"
            self.recordFullStrings.append(full_string)

            # Обновление словаря записанных временных слотов
            if row[5] not in self.dictZapisanye:
                self.dictZapisanye[row[5]] = []
            self.dictZapisanye[row[5]].append(row[6])

        self.textPlace.config(state='normal')
        self.textPlace.delete('3.0', tk.END)
        self.textPlace.insert('2.0', '\n')

        for i in range(len(self.recordBlocks)):
            self.textPlace.insert(tk.END,
                                  f"{(str(self.recordSecondNames[i])).capitalize()}\t\t  {(str(self.recordFirstNames[i])).capitalize()}\t\t"
                                  f"{str(self.recordBlocks[i]) + str(self.recordRooms[i])}\t\t"
                                  f"{str(self.recordTelNumber[i])}\t\t   {str(self.recordDates[i])}\t\t{str(self.recordTimes[i])}\n")

        self.textPlace.config(state='disabled')

        tk.messagebox.showinfo('Чтение из файла', "Успешно!")

    def infoToDocx(self):
        self.filePath = 'wordFiles//Word1.docx'
        self.document = Document()

        # Добавляем заголовок документа
        self.document.add_heading("Данные о записях на дежурство", 0)

        # Добавляем таблицу в документ
        table = self.document.add_table(rows=1, cols=7)

        # Определяем заголовки таблицы
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Фамилия'
        hdr_cells[1].text = 'Имя'
        hdr_cells[2].text = 'Блок'
        hdr_cells[3].text = 'Комната'
        hdr_cells[4].text = 'Номер'
        hdr_cells[5].text = 'Дата'
        hdr_cells[6].text = 'Время'

        # Заполняем таблицу данными
        for i in range(len(self.recordBlocks)):
            row_cells = table.add_row().cells
            rowData = [
                self.recordSecondNames[i],
                self.recordFirstNames[i],
                self.recordBlocks[i],
                self.recordRooms[i],
                self.recordTelNumber[i],
                self.recordDates[i],
                self.recordTimes[i]
            ]

            if all(isinstance(item, (int, float, str)) for item in rowData):
                for j, cell in enumerate(row_cells):
                    cell.text = str(rowData[j])

        # Сохраняем документ
        self.document.save(self.filePath)
        print(f"Файл '{self.filePath}' успешно создан.")

    #Дописать
    def saveAsExcelFile(self):

        filePath = tk.filedialog.asksaveasfilename(title="Сохранить файл", filetypes = [("Excel файлы", "*.xls"), ("Excel файлы", ".xlsx")])

        print(filePath)

    # ---------------------------------------------------------------------------------------------------------------ФУНКЦИИ

# region Класс для создания окна выбора сортировки
class SortWindow():

    def __init__(self, frame, parent):

        self.parent = parent

        self.sortWindow = tk.Toplevel(frame)

        #self.inactivity = Inactivity(frame)

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

        self.comboboxSort = ttk.Combobox(self.sortWindow, values=['В алфавитном порядке (от А до Я)', 'В алфавитном порядке (от Я до А)', 'По возрастанию блоков',
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

        if (self.selectedSort == "В алфавитном порядке (от А до Я)"):

            self.textExampleSort.insert('0.0', "До")

            self.textExampleSort.insert("end", "- Жоров\n- Ловчиновский\n- Бататкин\n\n")
            self.textExampleSort.insert("end", "После\n\n")
            self.textExampleSort.insert("end", "- Бататкин\n- Жоров\n- Ловчиновский")

            self.textExampleSort.config(state='disabled')

        elif (self.selectedSort == 'В алфавитном порядке (от Я до А)'):

            self.textExampleSort.insert('0.0', "До")

            self.textExampleSort.insert("end", "- Жоров\n- Ловчиновский\n- Бататкин\n\n")
            self.textExampleSort.insert("end", "После\n\n")
            self.textExampleSort.insert("end", "- Ловчиновский\n- Жоров\n- Бататкин")

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

    # Проверка не пустой ли комбобокс в сортировке
    def checkIsEntry(self):

        if (self.comboboxSort.get()):

            self.comboboxSort.config(state='readonly')

            return True

        else:
            return False


    # Функция которая проверяет не одна ли строка в тексте (для сортировки нужно минимум 2)
    def checkIsNoOneLine(self):

        numberLines = int(self.parent.textPlace.index("end-1c").split(".")[0])

        # Ну типо просто проверил считается ли кол-во строк
        if (numberLines >= 5):

            messagebox.showinfo("Результат", "Успешно!", parent=self.sortWindow)

            return True

        else:
            messagebox.showerror("Ошибка", "Недостаточно записей для сортировки\n(минимум 2)", parent=self.sortWindow)

            return False

    # Функция сортировки
    def funcOfSort(self):

        resChecking = self.checkIsEntry()

        if not resChecking:
            # parent для того, чтобы при вызове ошибки sortWindow не закрывалось
            messagebox.showerror("Ошибка", "Выберите способ сортировки!", parent=self.sortWindow)

        # Здесь сама сортировка через else (просто напиши else и там вызывай функции сортировки)
        elif resChecking:

            NoOneLine = self.checkIsNoOneLine()

            # Здесь вызываем функции сортировки
            if NoOneLine:

                if self.selectedSort == "По возрастанию блоков":

                    self.sortUpBlocks()

                elif self.selectedSort == "По убыванию блоков":

                    self.sortDownBlocks()

                elif self.selectedSort == "В алфавитном порядке (от А до Я)":

                    self.sortByAlphabetUp()

                elif self.selectedSort == "В алфавитном порядке (от Я до А)":
                    self.sortByAlphabetDown()

                elif self.selectedSort == "По дате":

                    self.sortByData()

    # Функция сортировки по возрастанию блоков
    def sortUpBlocks(self):
        # Создание списка кортежей (блок, индекс)
        combined_list = list(enumerate(self.parent.recordBlocks))
        # Сортировка по блокам
        sorted_combined_list = sorted(combined_list, key=lambda x: x[1])

        # Очищаем textPlace
        self.parent.textPlace.config(state='normal')
        self.parent.textPlace.delete('3.0', tk.END)
        self.parent.textPlace.insert('2.0', "\n")

        # Вставляем отсортированные строки в textPlace
        for index, _ in sorted_combined_list:
            fullString = self.parent.recordFullStrings[index] + "\n"
            self.parent.textPlace.insert(tk.END, fullString)

        self.parent.textPlace.config(state='disabled')

    # Функция сортировки по возрастанию блоков
    def sortDownBlocks(self):
        # Создание списка кортежей (блок, индекс)
        combined_list = list(enumerate(self.parent.recordBlocks))
        # Сортировка по блокам
        sorted_combined_list = sorted(combined_list, key=lambda x: x[1], reverse=True)

        # Очищаем textPlace
        self.parent.textPlace.config(state='normal')
        self.parent.textPlace.delete('3.0', tk.END)
        self.parent.textPlace.insert('2.0', "\n")

        # Вставляем отсортированные строки в textPlace
        for index, _ in sorted_combined_list:
            fullString = self.parent.recordFullStrings[index] + "\n"
            self.parent.textPlace.insert(tk.END, fullString)

        self.parent.textPlace.config(state='disabled')

    # Сортировка в алфавитном порядке
    def sortByAlphabetUp(self):
        # Сортируем полный список строк по алфавиту
        sortedAlphabetical = sorted(self.parent.recordFullStrings)

        self.parent.textPlace.config(state='normal')
        self.parent.textPlace.delete('3.0', tk.END)
        self.parent.textPlace.insert('2.0', "\n")

        for string in sortedAlphabetical:
            self.parent.textPlace.insert(tk.END, string + "\n")

        self.parent.textPlace.config(state='disabled')

    def sortByAlphabetDown(self):
        # Сортируем полный список строк по алфавиту в обратном порядке
        sortedAlphabetical = sorted(self.parent.recordFullStrings, reverse=True)

        self.parent.textPlace.config(state='normal')
        self.parent.textPlace.delete('3.0', tk.END)
        self.parent.textPlace.insert('2.0', "\n")

        for string in sortedAlphabetical:
            self.parent.textPlace.insert(tk.END, string + "\n")

        self.parent.textPlace.config(state='disabled')

    def sortByData(self):
        # Создание списка кортежей (месяц, день, индекс)
        combined_list = list(enumerate(zip(self.parent.recordMonthes, self.parent.recordDays)))

        # Преобразование данных к типу int для корректной сортировки
        combined_list = [(i, (int(month), int(day))) for i, (month, day) in combined_list]

        # Сортировка по месяцу и дню
        sorted_combined_list = sorted(combined_list, key=lambda x: (x[1][0], x[1][1]))

        # Очищаем textPlace
        self.parent.textPlace.config(state='normal')
        self.parent.textPlace.delete('3.0', tk.END)
        self.parent.textPlace.insert('2.0', "\n")

        # Вставляем отсортированные строки в textPlace
        for index, _ in sorted_combined_list:
            fullString = self.parent.recordFullStrings[index] + "\n"
            self.parent.textPlace.insert(tk.END, fullString)

        self.parent.textPlace.config(state='disabled')


#endregion

#Класс окна помощи
class HelpWindow():

    def __init__(self, frame, parent):

        self.parent = parent

        self.helpWindow = tk.Toplevel(frame)

        self.helpWindow.title("Использование")
        self.helpWindow.geometry("600x600")
        self.helpWindow.geometry("+630+270")
        self.helpWindow.resizable(width = False, height=False)

        self.frameHelp1 = tk.Frame(self.helpWindow, width = 600, height = 600)
        self.frameHelp1.place(x = 0, y = 0)

        self.frameHelp2 = tk.Frame(self.helpWindow, width=600, height=600)

        self.canvasHelp1 = tk.Canvas(self.frameHelp1, width=130, height=600, bg='#9681F0')
        self.canvasHelp1.place(x = 470, y = 0)

        self.labelHelp1 = tk.Label(self.frameHelp1, text = "Шаг 1", font = ("Montserrat", 14, 'bold'))
        self.labelHelp1.place(x = 180, y = 25)

        self.labelHelp1 = tk.Label(self.frameHelp1, text="Заполните все поля корректными данными и нажмите кнопку 'Записаться'\n"
                                                         "В случае ввода некорректных данных будет выдана ошибка.\n"
                                                         "В случае незаполнения всех полей так же будет выдана ошибка", font=("Montserrat", 10), justify = "left")
        self.labelHelp1.place(x=10, y=400)

        if self.frameHelp1:

            self.buttonNextHelp = MyButton(self.frameHelp1, text = "Далее", width=10, height=2,  x = 350, y = 500, command = self.fromHelp1ToHelp2)

    def fromHelp1ToHelp2(self):
        self.frameHelp1.pack_forget()
        self.frameHelp2.pack()


#Класс для работы с Excel
class WorkExcel():

    def __init__(self, parent, filepath):

        self.parent = parent

        self.filepath = filepath

        self.wb = Workbook()

        self.activeList = self.wb.active

        self.activeList.title = "Sheet1"

        self.boldFont = Font(bold=True)

        self.activeList['A1'] = "Фамилия"
        self.activeList['A1'].font = self.boldFont

        self.activeList['B1'] = "Имя"
        self.activeList['B1'].font = self.boldFont

        self.activeList['C1'] = "Блок"
        self.activeList['C1'].font = self.boldFont

        self.activeList['D1'] = "Комната"
        self.activeList['D1'].font = self.boldFont

        self.activeList['E1'] = "Номер"
        self.activeList['E1'].font = self.boldFont

        self.activeList['F1'] = "Дата"
        self.activeList['F1'].font = self.boldFont

        self.activeList['G1'] = "Время"
        self.activeList['G1'].font = self.boldFont

        self.wb.save(self.filepath)

    #Функция записывающая данные в таблицу
    def infoInFile(self, filepath):

        self.filepath = filepath

        for i in range(len(self.parent.recordBlocks)):

            rowData = [self.parent.recordSecondNames[i], self.parent.recordFirstNames[i],self.parent.recordBlocks[i],
                                    self.parent.recordRooms[i], self.parent.recordTelNumber[i], self.parent.recordDates[i], self.parent.recordTimes[i]]

            if all(isinstance(item, (int, float, str)) for item in rowData):

                self.activeList.append(rowData)

                self.wb.save(self.filepath)

app = MyApp()
app.mainloop()