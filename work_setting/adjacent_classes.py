# -*- coding: utf-8 -*-
'''
Модули с описанием смежных классов
'''
from PyQt5.QtWidgets import QWidget, QDialog, QLabel, QPushButton, QVBoxLayout, QProgressBar
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QObject, pyqtSignal, pyqtSlot

from work_lib import work_time, web_time, shutdown_lib
from work_setting import module, dialog
import datetime, time
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
            #запускаем функцию чтения данных с сайта марса во втором потоке
            self.RunWeb()
            
        #иначе работам через "отлов" включения/выключение компьютера
        else:
            #получаем текущую дату и время компа
            tekdateandtimeStart = datetime.datetime.now()
    
            tekyear = tekdateandtimeStart.year   #Текущий год
            tekmonth = tekdateandtimeStart.month #текущий месяц
            tekday = tekdateandtimeStart.day     #текущее число
            tekhour = tekdateandtimeStart.hour   #текущий час
            tekminute = tekdateandtimeStart.minute    #текущая минута
    
            #учитываем смещение
            min_offset = int(module.read_setting(10))   #получаем смещение
            #вычитаем смещение из минут
            if tekminute - min_offset >= 0:
                tekminute = tekminute - min_offset
            else:
                tekhour = tekhour - 1
                tekminute = 60 + (tekminute - min_offset)
            
            #записываем время включения компьютера
            work_time.start_work(tekminute, tekhour, tekday, tekmonth, tekyear)
            
            #просто записываем каждую минуту в файл текущее время (так тратим меньше ресурсов и не надо "ловить" выключение компьютера)
            while True:
                #получаем текущее время
                timeExit = datetime.datetime.now()
                #записываем текущее время в файл
                module.write_setting(timeExit.strftime("%d %m %Y %H:%M"), 25)
                time.sleep(60)
            self.finished_global.emit()

    #если функция вернет флаг выключения, то запустим окно с таймером на выключение ПК
    @pyqtSlot()
    def RunWeb(self):
        flg_shut = web_time.web_main()
        module.log_info("flg_shut: %s" % flg_shut)
        if flg_shut == True:
            module.write_setting(0, 28)    #ставим признак штатного завершения
            self.start_shut.emit(flg_shut)   #посылаем сигнал на запуск таймера для выключения
            self.finished_global.emit()
            shutdown_lib.signal_shutdown()
    
    #Основной метод счетчика выключения
    @pyqtSlot()
    def CountTime(self):
        self.show_wnd.emit()
        maxtime = 60
        for count in range(maxtime+1):
            step = maxtime - count
            self.intReady.emit(step)
            time.sleep(1)
        self.finished.emit()
     
        
#класс для таймера выключения
class ShutWindow(QWidget):
   
    def __init__(self):
        super(ShutWindow, self).__init__()

        #запуск формы
        self.initUI()
        
    def initUI(self):
        
        self.resize(200,200)                                # Устанавливаем фиксированные размеры окна
        self.setWindowFlags(Qt.FramelessWindowHint|Qt.WindowStaysOnTopHint)         # окно без рамки
        self.setAttribute(Qt.WA_TranslucentBackground)                          #окно прозрачное
        
        self.lbl = QLabel(self)                 #лейбл приветствия
        self.lbl.setFont(QFont('Arial', 12))        #Шрифт
        self.lbl.setText('До выключения остальсь:')
        self.lbl.adjustSize()                           #адаптивный размер в зависимости от содержимого
        
        self.lbl_timer = QLabel(self)                   #лейбл со счетчиком
        self.lbl_timer.setFont(QFont('Arial', 150))        #Шрифт
        self.lbl_timer.setText('60')
        self.lbl_timer.setStyleSheet('color: red')                 #цвет текста красный
        
        self.btn_stop = QPushButton('Остановить\nвыключение', self) #остановки счетчика
        self.btn_stop.setFont(QFont('Arial', 12))        #Шрифт
        self.btn_stop.clicked.connect(self.close_programm)      #действие по нажатию
        #расположение в окне
        self.v_box = QVBoxLayout()
        self.v_box.addWidget(self.lbl)
        self.v_box.addWidget(self.lbl_timer)
        self.v_box.addWidget(self.btn_stop)
        
        self.setLayout(self.v_box)
    
    #запись счетчика в лейбл
    def onShutReady(self, count):
        self.lbl_timer.setText(str(count).rjust(2, '0'))
    #отображение окна по сигналу
    def on_show_wnd(self):
        self.show()
    #по нажатию кнопки выключаем программу
    def close_programm(self):
        ex = dialog.MainWindow()
        ex.cleanUp()

##############################################################################################################
#объекты для потока с расчетом прогресса пересчета
class ThreadProgressRecount(QObject):
    finished = pyqtSignal()
    show_act = pyqtSignal()
    count_changed = pyqtSignal(int)          #сигнал для вывода прогресса перезаписи
    not_recount = pyqtSignal()
    donot_open = pyqtSignal()
    finished_progress = pyqtSignal()
    
    #функция для подсчета прогресса пересчета Exel файла
    def ThreadRecount(self):
        self.CountRecount()
    
    @pyqtSlot()
    def CountRecount(self):
        #получаем массив годов
        try:
            exel_year = work_time.exel_year()   #на существование файла
        except:
            self.donot_open.emit()
            self.finished.emit()
            self.finished_progress.emit()
            return
        step = 100/len(exel_year)
        count = 0
        self.show_act.emit()
        self.count_changed.emit(count)
        #в цикле вычисляем количество рабочих часов в каждом из месяцев в году
        for i in range(len(exel_year)):
            try:
                result = work_time.year_recount(int(exel_year[i]))
                #если пересчет не удался
                if result == False:
                    self.not_recount.emit()
                    break
            except:
                self.not_recount.emit()
                break
            
            count = count + step
            self.count_changed.emit(count)
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
        self.pbar.setValue(value)
        if value >= 100:
            time.sleep(1)   #для того что бы было видно 100%
        
    #показываем окно, блокируя другие
    def on_show_act(self):
        self.exec()

