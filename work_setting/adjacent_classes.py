# -*- coding: cp1251 -*-
'''
Модули с описанием смежных классов
'''
from PyQt5.QtWidgets import QWidget, QLabel, QPushButton, qApp, QApplication
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QThread, Qt, QBasicTimer, QObject, pyqtSignal, pyqtSlot

from work_lib import work_time, web_time, shutdown_lib
from work_setting import module, dialog
import datetime, time, sys, threading
from win32com.test.testIterators import SomeObject

#организуем многопоточнсть (считываем с сайта или ловим выключение компьютера в отдельном потоке)
class ShowShutOrWeb(QThread):
    def __init__(self):
        QThread.__init__(self)
        
    def run(self):
        #если выставлена галочка работы с сайтом - считываем сайт
        print(1)
        if(module.read_check()):
            #запускаем функцию чтения данных с сайта марса во втором потоке
            print(2)
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
        
#класс для таймера выключения
class ShutWindow(QWidget):
   
    def __init__(self):
        # Метод super() возвращает объект родителя класса MainWindow и мы вызываем его конструктор.
        # Метод __init__() - это конструктор класса в языке Python.
        super(ShutWindow, self).__init__()
                  
        self.ldl = QLabel(self)                 #лейбл приветствия
        self.ldl.setFont(QFont('Arial', 12))        #Шрифт
        self.ldl.setText('До выключения остальсь:')
        self.ldl.move(5, 10)                      #расположение в окне
        self.ldl.adjustSize()                           #адаптивный размер в зависимости от содержимого
        
        self.lbl_timer = QLabel(self)                   #лейбл со счетчиком
        self.lbl_timer.setFont(QFont('Arial', 100))        #Шрифт
        self.lbl_timer.setText('60')
        self.lbl_timer.move(25, 20)                      #расположение в окне
        self.lbl_timer.adjustSize()                           #адаптивный размер в зависимости от содержимого
        self.lbl_timer.setStyleSheet('color: red')                 #цвет текста красный
        
        self.btn_stop = QPushButton('Остановить\nвыключение', self) #остановки счетчика
        self.btn_stop.setFont(QFont('Arial', 12))        #Шрифт
        self.btn_stop.move(50, 150)                      #расположение в окне кнопки
        self.btn_stop.clicked.connect(self.closeprogramm)      #действие по нажатию

        #создаем поток внутри формы
        self.obj = ShutTimer()
        self.thread = QThread()
        #соединяем сигналы со слотами формы для вывода данных
        self.obj.intReady.connect(self.onIntReady)
        #перемещаем worker в thread
        self.obj.moveToThread(self.thread)
        #подключаем сигналы worker к слотам потока
        self.obj.finished.connect(self.thread.quit)
        #сигнал потокового подключения к методу worker
        self.thread.started.connect(self.obj.CountTime)
        #сигнал завершения потока закроет приложение
        #self.thread.finished.connect(app.exit)
        
        #запуск потока
        self.thread.start()
        #запуск формы
        self.initUI()
        
    def initUI(self):
        
        self.resize(200,200)                                # Устанавливаем фиксированные размеры окна
        self.setWindowFlags(Qt.FramelessWindowHint|Qt.WindowStaysOnTopHint)         # окно без рамки
        #self.setWindowOpacity(0.6)
        self.setAttribute(Qt.WA_TranslucentBackground)
    
    def onIntReady(self, i):
        self.lbl_timer.setText("{}".format(i))
        print(i)
    
    #по нажатию кнопки 
    def closeprogramm(self):
        qApp.quit()

#организуем многопоточнсть для таймера
class ShutTimer(QObject):
    
    finished = pyqtSignal()
    intReady = pyqtSignal(int)
    
    @pyqtSlot()
    def CountTime(self):
        maxtime = 10
        for count in range(maxtime):
            print('count = ', count)
            step = maxtime - count
            self.intReady.emit(step)
            time.sleep(1)
        
        self.finished.emit()
            
#вызываем окно с таймером
def app_ShutWindow():
    ''' не заработало...
    #в цикле считываем признак штатного завершения в потоке
    while True:
        run = module.read_setting(28)
        if int(run) == 0:
            break
        time.sleep(1)
    '''
    app = QApplication(sys.argv)
    ex = ShutWindow()
    ex.show()
    
    sys.exit(app.exec_())       
