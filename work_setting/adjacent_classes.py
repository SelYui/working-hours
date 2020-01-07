# -*- coding: cp1251 -*-
'''
Модули с описанием смежных классов
'''
from PyQt5.QtWidgets import QWidget, QStyle, QTextBrowser, QVBoxLayout, QLabel, QPushButton, qApp
from PyQt5.QtGui import QFont, QTextCursor
from PyQt5.QtCore import QThread, Qt

from work_lib import work_time, web_time, shutdown_lib
from work_setting import module, dialog
import datetime, webbrowser, time, sys

#организуем многопоточнсть (считываем с сайта или ловим выключение компьютера)
class ShowShutOrWeb(QThread):
    def __init__(self):
        QThread.__init__(self)
        
    def run(self):
        #если выставлена галочка работы с сайтом - считываем сайт
        if(module.read_check()):
            #запускаем функцию чтения данных с сайта марса во втором потоке
            web_time.web_main() 
            
        #иначе работам через "отлов" включения/выключение компьютера
        else:
            #получаем текущую дату и время компа
            tekdateandtimeStart = datetime.datetime.now()
    
            tekyear = tekdateandtimeStart.year   #Текущий год
            tekmonth = tekdateandtimeStart.month #текущий месяц
            tekday = tekdateandtimeStart.day     #текущее число
            tekhour = tekdateandtimeStart.hour   #текущий час
            tekminute = tekdateandtimeStart.minute    #текущая минута
            #записываем в Exel файл время последнего выключения компьютера
            work_time.write_exit()
    
            #записываем время включения компьютера
            work_time.start_work(tekminute, tekhour, tekday, tekmonth, tekyear)
            
            #запускаем бесконечный цикл для опроса сигналов виндовс
            #shutdown_lib.shutdown_lib()
            #2ой вариант, просто записываем каждую минуту в файл текущее время (так тратим меньше ресурсов и не надо "ловить" выключение компьютера)
            while True:
                #получаем текущее время
                timeExit = datetime.datetime.now()
                #записываем текущее время в файл
                module.write_timeExit(timeExit.strftime("%d %m %Y %H:%M"))
                time.sleep(60)

#создаем окно с подсказкой
class AdjWindow(QWidget):
   
    def __init__(self):
        # Метод super() возвращает объект родителя класса MainWindow и мы вызываем его конструктор.
        # Метод __init__() - это конструктор класса в языке Python.
        super(AdjWindow, self).__init__()
        #создаем пвлитру окна
        #appearance = self.palette()
        #appearance.setColor(QPalette.Normal, QPalette.Window, QColor("white"))
                  
        self.resize(350,500)                                # Устанавливаем фиксированные размеры окна
        self.setWindowTitle("Как узнать свой индивидуальный номер")  # Устанавливаем заголовок окна
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_TitleBarContextHelpButton))   #устанавливаем одну из стандартных иконку
        #self.setPalette(appearance)                         #Применяем палитру к нашему окну
        '''
        #создаем текст с инструкцией
        self.le_help = QTextEdit(self)
        self.le_help.setFont(QFont('Arial', 12))        #Шрифт
        self.le_help.setText('Здесь напишем инструкцию')
        self.le_help.move(0, 0)                      #расположение в окне
        self.le_help.adjustSize()                           #адаптивный размер в зависимости от содержимого
        '''
        #создаем текст с инструкцией
        
        text = module.read_help()#'Зайдите на сайт: <br><a href="http://www.mars/asu/report/enterexit/">www.mars/asu/report/enterexit/</a><br> текст'
        #чтобы наше поле занимало все окно
        vbox = QVBoxLayout(self)
        #создаем поле с текстом инструкции и ссылкой
        self.pole_vivod = QTextBrowser(self)
        self.pole_vivod.setFont(QFont('Arial', 14))        #Шрифт
        self.pole_vivod.anchorClicked['QUrl'].connect(self.linkClicked)
        self.pole_vivod.setOpenLinks(False)     #Запрет удаления ссылки
        #self.pole_vivod.move(0, 0)
        vbox.addWidget(self.pole_vivod)
        self.setLayout(vbox)
        
        self.pole_vivod.append(text)
        self.pole_vivod.moveCursor(QTextCursor.Start)
        
    #обрабатываем клик по ссылке
    def linkClicked(self, url):
        webbrowser.open(url.toString()) 
        
#класс для таймера выключения
class ShutWindow(QWidget):
   
    def __init__(self):
        # Метод super() возвращает объект родителя класса MainWindow и мы вызываем его конструктор.
        # Метод __init__() - это конструктор класса в языке Python.
        super(ShutWindow, self).__init__()
        #создаем пвлитру окна
                  
        self.resize(200,200)                                # Устанавливаем фиксированные размеры окна
        self.setWindowFlags(Qt.FramelessWindowHint|Qt.WindowStaysOnTopHint)         # окно без рамки
        #self.setWindowOpacity(0.6)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        
        self.ldl = QLabel(self)
        self.ldl.setFont(QFont('Arial', 12))        #Шрифт
        self.ldl.setText('До выключения остальсь:')
        self.ldl.move(5, 10)                      #расположение в окне
        self.ldl.adjustSize()                           #адаптивный размер в зависимости от содержимого
        
        self.timer = QLabel(self)
        self.timer.setFont(QFont('Arial', 100))        #Шрифт
        self.timer.setText('60')
        self.timer.move(25, 20)                      #расположение в окне
        self.timer.adjustSize()                           #адаптивный размер в зависимости от содержимого
        self.timer.setStyleSheet('color: red')                 #цвет текста красный
        
        self.btn_stop = QPushButton('Остановить\nвыключение', self)
        self.btn_stop.setFont(QFont('Arial', 12))        #Шрифт
        self.btn_stop.move(50, 150)                      #расположение в окне кнопки
        self.btn_stop.clicked.connect(self.closeprogramm)      #действие по нажатию
        
    def closeprogramm(self):
        self.hide()
        qApp.quit()