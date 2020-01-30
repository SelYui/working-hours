# -*- coding: utf-8 -*-
'''
Модули с описанием смежных классов
'''
from PyQt5.QtWidgets import QWidget, QDialog, QLabel, QPushButton, qApp, QApplication, QHBoxLayout, QVBoxLayout, QMessageBox, QProgressBar
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QThread, Qt, QBasicTimer, QObject, pyqtSignal, pyqtSlot

from work_lib import work_time, web_time, shutdown_lib
from work_setting import module, dialog
import datetime, time, sys, threading
#from win32com.test.testIterators import SomeObject

#организуем многопоточнсть (считываем с сайта или ловим выключение компьютера в отдельном потоке)
class ShowShutOrWeb(QObject):
    #объявляем все сигналы
    finished = pyqtSignal()
    finished_global = pyqtSignal()
    intReady = pyqtSignal(int)
    start_shut = pyqtSignal(int)
    show_wnd = pyqtSignal()
    
    def ShutOrWeb(self):
        #если выставлена галочка работы с сайтом - считываем сайт
        if(int(module.read_setting(16))):
            print(1234)
            #запускаем функцию чтения данных с сайта марса во втором потоке
            self.RunWeb()
            
        #иначе работам через "отлов" включения/выключение компьютера
        else:
            print(1235)
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
                module.write_setting(timeExit.strftime("%d %m %Y %H:%M"),25)
                time.sleep(60)
            self.finished_global.emit()

    #если функция вернет 1234, то запустим окно с таймером на выключение ПК
    @pyqtSlot()
    def RunWeb(self):
        print(12345)
        flg_shut = web_time.web_main()
        module.log_info("flg_shut: %s" % flg_shut)
        if flg_shut == True:
            module.write_setting(0, 28)    #ставим признак штатного завершения
            self.start_shut.emit(flg_shut)   #посылаем сигнал на запуск таймера для выключения
            print(4563)
            self.finished_global.emit()
            print(4564)
            shutdown_lib.signal_shutdown()
    
    #Основной метод счетчика
    @pyqtSlot()
    def CountTime(self):
        self.show_wnd.emit()
        maxtime = 60
        for count in range(maxtime+1):
            print('count = ', count)
            step = maxtime - count
            self.intReady.emit(step)
            time.sleep(1)
        self.finished.emit()
     
        
#класс для таймера выключения
class ShutWindow(QWidget):
   
    def __init__(self):
        # Метод super() возвращает объект родителя класса MainWindow и мы вызываем его конструктор.
        # Метод __init__() - это конструктор класса в языке Python.
        super(ShutWindow, self).__init__()

        #запуск формы
        self.initUI()
        
    def initUI(self):
        
        self.resize(200,200)                                # Устанавливаем фиксированные размеры окна
        self.setWindowFlags(Qt.FramelessWindowHint|Qt.WindowStaysOnTopHint)         # окно без рамки
        #self.setWindowOpacity(0.6)
        self.setAttribute(Qt.WA_TranslucentBackground)                          #окно прозрачное
        
        self.lbl = QLabel(self)                 #лейбл приветствия
        self.lbl.setFont(QFont('Arial', 12))        #Шрифт
        self.lbl.setText('До выключения остальсь:')
        #self.lbl.move(5, 10)                      #расположение в окне
        self.lbl.adjustSize()                           #адаптивный размер в зависимости от содержимого
        
        self.lbl_timer = QLabel(self)                   #лейбл со счетчиком
        self.lbl_timer.setFont(QFont('Arial', 150))        #Шрифт
        self.lbl_timer.setText('60')
        #self.lbl_timer.move(25, 20)                      #расположение в окне
        #self.lbl_timer.adjustSize()                           #адаптивный размер в зависимости от содержимого
        self.lbl_timer.setStyleSheet('color: red')                 #цвет текста красный
        
        self.btn_stop = QPushButton('Остановить\nвыключение', self) #остановки счетчика
        self.btn_stop.setFont(QFont('Arial', 12))        #Шрифт
        #self.btn_stop.move(50, 150)                      #расположение в окне кнопки
        self.btn_stop.clicked.connect(self.close_programm)      #действие по нажатию
        
        self.v_box = QVBoxLayout()
        self.v_box.addWidget(self.lbl)
        self.v_box.addWidget(self.lbl_timer)
        self.v_box.addWidget(self.btn_stop)
        
        self.setLayout(self.v_box)
    
    def onShutReady(self, count):
        self.lbl_timer.setText(str(count).rjust(2, '0'))
        print(count)
    
    def on_show_wnd(self):
        self.show()
    
    #по нажатию кнопки 
    def close_programm(self):
        ex = dialog.MainWindow()
        ex.cleanUp()
        #сам выход
        #sys.exit(0)
'''     
#вызываем окно с таймером (вроде не нужно)
def app_ShutWindow():

    app = QApplication(sys.argv)
    ex = ShutWindow()
    ex.show()
    
    sys.exit(app.exec_())       
'''
##############################################################################################################
#объекты для потока с расчетом прогресса пересчета
class ThreadProgressRecount(QObject):
    finished = pyqtSignal()
    show_act = pyqtSignal()
    count_changed = pyqtSignal(int)          #сигнал для вывода прогресса перезаписи
    finished_progress = pyqtSignal()
    
        #функция для подсчета прогресса пересчета Exel файла
    def ThreadRecount(self):
        print(741)
        self.CountRecount()
    
    @pyqtSlot()
    def CountRecount(self):
        #получаем массив годов
        exel_year = work_time.exel_year()
        step = 100/len(exel_year)
        count = 0
        self.show_act.emit()
        self.count_changed.emit(count)
        print('exel_year = ', exel_year)
        #в цикле вычисляем количество рабочих часов в каждом из месяцев в году
        for i in range(len(exel_year)):
            result = work_time.year_recount(int(exel_year[i]))
            print('result = ', result)
            '''
            if result == 1:
                #module.log_info("exel_recount exception: %s" % e)
                message = 'Ошибка пересчета файла\nПожалуйста проверьте, что:\n    - файл составлен корректно (имеет все названия месяцев, все времена прихода и ухода)\n    - файл закрыт\nИ повторите попытку'
                reply = QMessageBox.warning(self, 'Ошибка', message, QMessageBox.Retry | QMessageBox.Cancel, QMessageBox.Retry)
                # если нажато "Повтор", запускаемся еще раз
                if reply == QMessageBox.Retry:
                    self.run()
            '''
            count = count + step
            self.count_changed.emit(count)
        #self.countChanged.emit(100)
        print('Пересчитал2')
        self.finished.emit()
        self.finished_progress.emit()
    
#окно с прогрессом пересчета
class ProgressRecount(QDialog):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        #окно без рамки
        self.resize(400, 50)
        self.setWindowFlags(Qt.FramelessWindowHint)         # окно без рамки
        self.setAttribute(Qt.WA_TranslucentBackground)      #окно прозрачное
        #создаем ползунок прогресса
        self.pbar = QProgressBar(self)
        self.pbar.setFont(QFont('Arial', 14))
        self.pbar.setValue(0) 
        #запихиваем его в окно
        self.vbox = QVBoxLayout()
        self.vbox.addWidget(self.pbar)
        self.setLayout(self.vbox)
    
    #функция для которой расчитывается прогресс
    def doAction(self, value):
        self.show()
        self.pbar.setValue(value)
        if value >= 100:
            print('Пересчитал', value)
    #показываем окно, блокируя другие
    def on_show_act(self):
        self.exec()

