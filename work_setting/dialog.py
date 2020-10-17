# -*- coding: utf-8 -*-
'''
Класс для вывода акна с настройками
'''
import os, sys, subprocess, webbrowser, datetime
from work_setting import module, adjacent_classes
from work_lib import work_time, shutdown_lib

from PyQt5.QtWidgets import (QPushButton, QLineEdit, QLabel, QDesktopWidget, QToolTip, QSystemTrayIcon, QDialog, QMessageBox, QAction,
    QFileDialog, QApplication, QMenu, QSpinBox, QCheckBox, QWidget, QStyle, QTextBrowser, QHBoxLayout, QVBoxLayout, QGridLayout)
from PyQt5.QtGui import QIcon, QFont, QTextCursor
from PyQt5.QtCore import Qt, QThread, QTimer
from PyQt5.Qt import QIntValidator, QRegExp, QRegExpValidator

url_mars = 'http://www.mars/asu/report/enterexit/'

#создаем окно наших настроек
class MainWindow(QWidget):

    def __init__(self):
        # Метод super() возвращает объект родителя класса MainWindow и мы вызываем его конструктор.
        # Метод __init__() - это конструктор класса в языке Python.
        super().__init__()
        
        #инициализация потоков
        self.initThread()
        # инициализация GUI
        self.initUI()
        # вычисление текущего времени
        self.NowYourTime()
        
    def initThread(self):
        #создаем поток внутри формы
        self.obj = adjacent_classes.ShowShutOrWeb()
        self.obj_progress = adjacent_classes.ThreadProgressRecount()
        self.thread = QThread()
        self.thread_progress = QThread()
        self.wind = adjacent_classes.ShutWindow()
        self.act = adjacent_classes.ProgressRecount()
        #перемещаем в thread
        self.obj.moveToThread(self.thread)
        self.obj_progress.moveToThread(self.thread_progress)
        #подключаем сигналы к слотам потока и к слотами формы для вывода данных
        self.obj.start_shut.connect(self.obj.CountTime)     #сигнал на запуск таймера выключения
        self.obj.intReady.connect(self.wind.onShutReady)    #сигнал для вывода счетчика таймера
        self.obj.show_wnd.connect(self.wind.on_show_wnd)    #показываем окно с счетчиком
        self.obj.finished_global.connect(self.thread.quit)  #конец потока
        self.obj_progress.count_changed.connect(self.act.doAction)    #сигнал для вывода прогресса пересчета
        self.obj_progress.show_act.connect(self.act.on_show_act)    #показываем прогресс
        self.obj_progress.not_recount.connect(self.do_not_rec)      #не удалось пересчитать
        self.obj_progress.donot_open.connect(self.donot_open)       #не удалось открыть файл
        self.obj_progress.finished_progress.connect(self.thread_progress.quit)
        #подключаем сигнал потокового подключения к методу
        self.thread.started.connect(self.obj.ShutOrWeb)
        self.thread.finished.connect(self.cleanUp)
        self.thread_progress.started.connect(self.obj_progress.ThreadRecount)
        self.thread_progress.finished.connect(self.off_show_act)
        
    def initUI(self):
        #получаем настройки из файла
        WorkPath = module.read_setting(4)
        WorkName = module.read_setting(1)
        AbsPath = os.path.abspath(WorkPath + '/' + WorkName)
        WorkDiner = module.read_setting(7)
        WorkOffset = int(module.read_setting(10))
        WorkReload = int(module.read_setting(13))
        YouNumber = module.read_setting(19)
        CheckNum = int(module.read_setting(16))
        WorkShut = module.read_setting(22)
        #признак необходимости перезагрузки 0-не надо, 5577 - надо
        self.flg_shutdown = 0
        self.now_time = 0
     
    #Создаем само окно
        self.resize(495, 360)                            # Устанавливаем начальные размеры окна
        self.setWindowTitle("Подсчёт рабочего времени")  # Устанавливаем заголовок окна
        self.setWindowIcon(QIcon('icon\Bill_chipher.jpg'))         # Устанавливаем иконку
        self.center()               # помещаем окно в центр экрана
    
    
    #виджеты для имени файла
        self.lblN = QLabel(self)                    #создаем лейбел с именем файла
        self.lblN.setFont(QFont('Arial', 12))        #Шрифт
        self.lblN.setText('Настройки файла ' + WorkName)

        
    #виджеты для пути к файлу
        self.lblI = QLabel(self)                    #cоздаем строку с инструкцией
        self.lblI.setFont(QFont('Arial', 12))        #Шрифт
        self.lblI.setText("Путь к файлу:")
        
        self.le = QLineEdit(self)                   #создаем строку для ввода пути к файлу
        self.le.setFont(QFont('Arial', 12))         #Шрифт
        self.le.setText(AbsPath)  #пишем путь из настроечного файла
        
        self.btnI = QPushButton('Изменить', self)       #создаем кнопку для изменения расположения файла
        self.btnI.setFont(QFont('Arial', 12))        #Шрифт
        #self.btnI.clicked.connect(self.getfile)      #действие по нажатию
        self.btnI.clicked.connect(self.savefile)
        self.btnI.setAutoDefault(True)               # click on <Enter>

        self.btnO = QPushButton('Открыть', self)    #создаем кнопку для открытия директории/файла
        self.btnO.setFont(QFont('Arial', 12))        #Шрифт
        self.btnO.clicked.connect(self.opendirectory)      #действие по нажатию
        self.btnO.setAutoDefault(True)               # click on <Enter>
        
        self.btnRec = QPushButton('Пересчитать', self)    #создаем кнопку для пересчета времени в файле
        self.btnRec.setFont(QFont('Arial', 12))        #Шрифт
        self.btnRec.clicked.connect(self.exel_recount)      #действие по нажатию
        self.btnRec.setAutoDefault(True)               # click on <Enter>
        
        
    #виджеты для времени обеда
        self.lblO = QLabel(self)                        #cоздаем строку с инструкцией для смещения
        self.lblO.setFont(QFont('Arial', 12))        #Шрифт
        self.lblO.setText("Обед (мин.):")
        
        self.spbO = QSpinBox(self)                   #создаем SpinBox для выбора времени
        self.spbO.setFont(QFont('Arial', 12))        #Шрифт
        self.spbO.setMaximum(60)                     #верхняя граница счетчика
        self.spbO.setMinimum(0)                      #нижняя граница счетчика
        self.spbO.setSingleStep(5)                   #шаг
        self.spbO.setValue(int(WorkDiner))
        
    #cоздаем виджеты для смещения
        self.lblS = QLabel(self)                        #cоздаем строку с инструкцией для смещения
        self.lblS.setFont(QFont('Arial', 12))        #Шрифт
        self.lblS.setText("Смещение (мин.):")
        
        self.spb = QSpinBox(self)                   #создаем SpinBox для выбора времени
        self.spb.setFont(QFont('Arial', 12))        #Шрифт
        self.spb.setMaximum(60)                     #верхняя граница счетчика
        self.spb.setMinimum(0)                      #нижняя граница счетчика
        self.spb.setValue(WorkOffset)


    #cоздаем виджеты для времени ухода
        self.lblU = QLabel(self)                        #cоздаем строку с инструкцией для времени безопасного ухода
        self.lblU.setFont(QFont('Arial', 12))        #Шрифт
        self.lblU.setText("Возможный уход (мин.):")
        
        self.spblered = QSpinBox(self)                  #создаем SpinBox для выбора времени ухода
        self.spblered.setFont(QFont('Arial', 12))        #Шрифт
        self.spblered.setMaximum(60)                     #верхняя граница счетчика
        self.spblered.setMinimum(0)                      #нижняя граница счетчика
        self.spblered.setSingleStep(5)                        #шаг
        self.spblered.setValue(WorkReload)

        
    #cоздаем виджеты для индивидуального номера
        self.lblCh = QLabel(self)                   #cоздаем строку с инструкцией для индивидуального номера
        self.lblCh.setFont(QFont('Arial', 12))        #Шрифт
        self.lblCh.setText("Ваш номер на сайте:")
        
        self.lenum = QLineEdit(self)                #создаем строку для ввода индивидуального номера сотрудника
        self.lenum.setFont(QFont('Arial', 12))         #Шрифт
        self.lenum.setValidator(QIntValidator(0,9999))
        self.lenum.setText(YouNumber)          #пишем путь из настроечного файла
        self.lenum.returnPressed.connect(self.save_setting_btn) # click on <Enter>
        self.lenum.setEnabled(False)        #делаем строку неактивной

        
    #виджеты для времени выключения
        self.lblSh = QLabel(self)                   #cоздаем строку с инструкцией для времени выключения
        self.lblSh.setFont(QFont('Arial', 12))        #Шрифт
        self.lblSh.setText("Выключать компьютер после:")
        
        #создаем Валидатор для строки времени
        hour = '(2[0123]|([0-1][0-9]))'
        minute = '[0-5][0-9]'
        simbol = '([0-5][0-9]|:)'
        timeRange = QRegExp('^' + hour + simbol + minute + '$')
        timeVali = QRegExpValidator(timeRange, self)
        
        self.leshut = QLineEdit(self)                   #создаем строку для ввода выключения компьютера
        self.leshut.setFont(QFont('Arial', 12))         #Шрифт
        self.leshut.setText(str(WorkShut))          #пишем путь из настроечного файла
        self.leshut.setValidator(timeVali)
        self.leshut.textChanged.connect(self.time_shutdow)      #сигнал по изменению текста
        self.leshut.returnPressed.connect(self.save_setting_btn)    # click on <Enter>
        self.leshut.setEnabled(False)        #делаем строку неактивной

    #виджеты для выбора режима работы (по вкл/выкл, по сайту)
        self.chweb = QCheckBox('Брать время с сайта Марса', self)   #создаем checkbox для выбора подчета времени с сайта Марса
        self.chweb.setFont(QFont('Arial', 12))          #Шрифт
        self.chweb.stateChanged.connect(self.webtime)           #действие по нажатию

        #выставляем в соответствии с настройками
        if(CheckNum):
            self.chweb.setChecked(True)
        else:
            self.chweb.setChecked(False)
            
        self.btnch = QPushButton('?', self)         #создаем кнопку для подсказки
        self.btnch.setFont(QFont('Arial', 12))        #Шрифт
        try:
            self.btnch.clicked.connect(self.openhelp)      #действие по нажатию
        except Exception as e:
            module.log_info('Error openhelp: %s' %e)
        self.btnch.setAutoDefault(True)               # click on <Enter>
    
         
    #виджеты для сохранения
        self.btn_save = QPushButton('Сохранить', self)  #создаем кнопку для всплывающего диалогового окна
        self.btn_save.setFont(QFont('Arial', 12))        #Шрифт
        self.btn_save.clicked.connect(self.save_setting_btn)      #действие по нажатию
        self.btn_save.setDefault(True)                      #значально будет выделена
        self.btn_save.setAutoDefault(True)               # click on <Enter>
        
        self.le.returnPressed.connect(self.btn_save.click)  #действия в строке по интеру
        self.lenum.returnPressed.connect(self.btn_save.click)  #действия в строке по интеру

    #раскладываем виджеты в главном окне
        self.layout_in_main()
    
    # Инициализируем иконку Tray
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon('icon\Bill_chipher.jpg')) #устанавливаем пользовательскую иконку

        show_action = QAction(QIcon('icon\Programming-Show.png'), "Настройки", self)
        quit_action = QAction(QIcon('icon\exit.png'), "Выход", self)
        show_action.triggered.connect(self.show)        #при нажатии на show окно открывается
        quit_action.triggered.connect(self.cleanUp)        #при нажатии на quit приложение закрывается qApp.quit
        tray_menu = QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
    #создаем подсказки
        QToolTip.setFont(QFont('Arial', 10))    # метод устанавливает шрифт, используемый для показа всплывающих подсказок.
        self.setToolTip('Это окно выбора основных настроек программы')  # создаем подсказку для окна
        self.lblN.setToolTip('Текущее имя файла, в котором хранится Ваше рабочее время')
        self.le.setToolTip('Файл с Вашим рабочим временем находится по этому пути\n' + AbsPath)
        self.lblI.setToolTip('Файл с Вашим рабочим временем находится по этому пути')
        self.btnI.setToolTip('Выберете файл рабочего времени')    # создаем подсказку для кнопки
        #self.btnS.setToolTip('Создать новый файл рабочего времени')    # создаем подсказку для кнопки
        self.btnO.setToolTip('Открыть папку с файлом')    # создаем подсказку для кнопки
        self.btnRec.setToolTip('Пересчет рабочего времени в файле\n' + WorkName)    # создаем подсказку для кнопки
        self.btn_save.setToolTip('Сохранить выставленные настройки')    # создаем подсказку для кнопки
        self.lblO.setToolTip('Укажите время Вашего обеда в мин.')
        self.spbO.setToolTip('Укажите время Вашего обеда в мин.')
        self.spb.setToolTip('Укажите смещение от времени вкл/выкл ПК')    # создаем подсказку
        self.lblS.setToolTip('Сколько Вам идти от КПП до рабочего места?')    # создаем подсказку для кнопки
        #self.lblSm.setToolTip('Укажите смещение в минутах')
        self.lblU.setToolTip('Если компьютер выключится на заданное время,\nто в файле Вашего рабочего времени уход не зафиксируется')
        self.spblered.setToolTip('Введите время неучетного выхода в мин.')
        self.lblCh.setToolTip('Введите Вашь номер на сайте')
        self.chweb.setToolTip('Время Вашего прихода фиксируется на сайте:\n' + url_mars + '\nбрать время Вашего прихода от туда?')
        self.lenum.setToolTip('Вашь номер на сайте: ' + YouNumber)
        self.btnch.setToolTip('Нажмите, чтобы понять как узнать свой номер на сайте')
        self.lblSh.setToolTip('Если Вы вышли за КПП после этого времени, выключаю компьютер')
        self.leshut.setToolTip('Введите время в формате:\n00:00')
        
    #если нет имени файла, то показать окно
        if WorkName == '':
            self.show()

    #Функция для расположения виджетов в окне
    def layout_in_main(self):
    #создаем слои и сетки
        self.h_boxN = QHBoxLayout()
        self.h_boxN.addWidget(self.lblN)
        
        self.h_boxI = QHBoxLayout()
        self.h_boxI.addWidget(self.lblI)
        #сетка для кнопок
        self.grid_btnIO = QGridLayout()
        self.grid_btnIO.addWidget(self.le, 1, 0, 1, 4)
        self.grid_btnIO.addWidget(self.btnI, 1, 5)
        self.grid_btnIO.addWidget(self.btnO, 2, 5)
        
        self.h_boxO = QHBoxLayout()
        self.h_boxO.addWidget(self.lblO)
        self.h_boxO.addWidget(self.spbO)
        self.h_boxO.addStretch(1)
        
        self.h_boxS = QHBoxLayout()
        self.h_boxS.addWidget(self.lblS)
        self.h_boxS.addWidget(self.spb)
        self.h_boxS.addStretch(1)
        
        self.h_boxU = QHBoxLayout()
        self.h_boxU.addWidget(self.lblU)
        self.h_boxU.addWidget(self.spblered)
        self.h_boxU.addStretch(1)
        
        self.h_box_web = QHBoxLayout()
        self.h_box_web.addWidget(self.chweb)
        self.h_box_web.addWidget(self.btnch)
        self.h_box_web.addStretch(1)
        #сетка для остальных виджетов
        self.grid_other = QGridLayout()
        self.grid_other.addWidget(self.lblCh, 1, 0)
        self.grid_other.addWidget(self.lenum, 1, 1)
        self.grid_other.addWidget(self.lblSh, 2, 0, 1, 4)
        self.grid_other.addWidget(self.leshut, 2, 5)
        self.grid_other.addWidget(self.btnRec, 3, 6)
        self.grid_other.addWidget(self.btn_save, 3, 7)
    #заносим все в горисонтальный слой
        self.v_box = QVBoxLayout()
        self.v_box.addLayout(self.h_boxN)
        self.v_box.addLayout(self.h_boxI)
        self.v_box.addLayout(self.grid_btnIO)
        self.v_box.addLayout(self.h_boxO)
        self.v_box.addLayout(self.h_boxS)
        self.v_box.addLayout(self.h_boxU)
        self.v_box.addLayout(self.h_box_web)
        self.v_box.addLayout(self.grid_other)
    #помещаем на страницу
        self.setLayout(self.v_box)
        
    #Функция для центрирования окна в экране пользователя
    def center(self):
        qr = self.frameGeometry()           # получаем прямоугольник, точно определяющий форму главного окна.
        cp = QDesktopWidget().availableGeometry().center()  # выясняем разрешение экрана нашего монитора. Из этого разрешения, мы получаем центральную точку.
        qr.moveCenter(cp)                   # устанавливаем центр прямоугольника в центр экрана. Размер прямоугольника не изменяется.
        self.move(qr.topLeft())             # перемещаем верхнюю левую точку окна приложения в верхнюю левую точку прямоугольника qr, таким образом центрируя окно на нашем экране.
    

    #диалоговое окно выбора нового файла (не используется)
    def getfile(self):
        dir_path = module.read_setting(4) + '/' + module.read_setting(1)
        fname = QFileDialog.getOpenFileName(self, 'Выбрать файл', dir_path, 'Exel files (*.xls)')
        #если новый файл выбран, переписываем путь в настройках и в наших текстовых виджетах
        if fname != ('', ''):
            new_dir = os.path.dirname(fname[0])    #путь папки в которой лежит файл
            new_name = os.path.basename(fname[0])   #имя файла
            #запись в настроечный файл нового имени файла
            module.write_setting(new_name, 1)
            #запись в настроечный файл нового пути
            module.write_setting(new_dir, 4)
            self.le.setText(fname[0])
            self.lblN.setText('Настройки файла ' + new_name)
        
    #диалоговое окно сохранения нового файла
    def savefile(self):
        dir_path = module.read_setting(4) + '/' + module.read_setting(1)
        fname = QFileDialog.getSaveFileName(self, 'Выбрать файл', dir_path, 'Exel files (*.xls)',
                                            options = QFileDialog.DontConfirmOverwrite)
        #если новый файл выбран, переписываем путь в настройках и в наших текстовых виджетах
        if fname != ('', ''):
            new_dir = os.path.dirname(fname[0])    #путь папки в которой лежит файл
            new_name = os.path.basename(fname[0])   #имя файла
            #создание нового файла
            if(not os.path.exists(fname[0])):
                if(work_time.new_timework_file(fname[0])!=-1):
                    self.le.setText(fname[0])
                    self.lblN.setText('Настройки файла ' + new_name)
                else:
                    QMessageBox.warning(self, 'Предупреждение','Не удалось создать файл')
                    self.le.setText(dir_path)
                    self.lblN.setText('Настройки файла ' + module.read_setting(1))
            else:
                self.le.setText(fname[0])
                self.lblN.setText('Настройки файла ' + new_name)
    
    #сохраняем настройки с учетом того что введено в строку
    def save_setting_btn(self):
        #получаем путь к файлу, имя файла Exel и номер пользователя из нашего настроечного файла
        WorkPath = module.read_setting(4)
        WorkName = module.read_setting(1)
        WorkNumb = module.read_setting(19)
        WorkShut = module.read_setting(22)
        AbsPath = os.path.abspath(WorkPath + '/' + WorkName)

        dir_path = self.le.text()   #получаем путь к новому файлу
        you_numb = self.lenum.text()
        shut_time = self.leshut.text()
        
        #сохраняем новый путь Exel файла в настроечный файл
        dir_path = os.path.abspath(dir_path)        #получаем абсолютный путь к файлу
        f_name, f_exstension = os.path.splitext(dir_path)   #выделяем название и расширение
        #проверка имени файла на корректность и расширение
        if f_name != '' and f_exstension == '.xls':
            #если выбранный файл существует записываем в настройки путь к нему
            if os.path.exists(dir_path):
                #запись в настроечный файл нового имени файла
                module.write_setting(os.path.basename(dir_path), 1)
                #запись в настроечный файл нового пути и в строку
                module.write_setting(os.path.dirname(dir_path), 4)
                self.le.setText(dir_path)
                self.flg_shutdown = 5577 #после изменения этой настройки предложим перезагрузить компьютер
            #если файла нет, то сздать его?
            else:
                reply = QMessageBox.question(self, 'Сообщение', 'Файл не найден.\nСоздать новый файл?', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                #QMessageBox.warning(self, 'Предупреждение','Файл не найден.\nСоздать новый файл?')
                # если нажато "Да", создаем файл и сохраняем его путь в настроечный файл
                if reply == QMessageBox.Yes:
                    if (work_time.new_timework_file(dir_path)!=-1):
                        #запись в настроечный файл нового имени файла
                        module.write_setting(os.path.basename(dir_path), 1)
                        #запись в настроечный файл нового пути и в строку
                        module.write_setting(os.path.dirname(dir_path), 4)
                        self.le.setText(dir_path)
                        self.flg_shutdown = 5577 #после изменения этой настройки предложим перезагрузить компьютер
                    else:
                        QMessageBox.warning(self, 'Предупреждение','Не удалось создать файл')
                        self.le.setText(AbsPath)
                else:
                    self.le.setText(AbsPath)    #вернуть исходное значение пути
        #если путь пустой выводим сообщение
        else:
            QMessageBox.warning(self, 'Предупреждение','Файл не должен иметь пустое имя\nи должен иметь расширение .xls')
            self.le.setText(AbsPath)
        
        #сохраняем значения из SpinBox в файл
        module.write_setting(self.spb.value(),10)
        module.write_setting(self.spblered.value(),13)
        module.write_setting(self.spbO.value(), 7)
        
        #если галочка выставлена
        if int(module.read_setting(16)):
            #сохраняем новый номер пользователя в файл
            if you_numb != '':
                #если поле не пустое - записываем новое значение в файл
                module.write_setting(you_numb,19)   #по идее после этого не обязательно перезагружать комп
            else:
                #иначе, выводим предупреждение
                QMessageBox.warning(self, 'Предупреждение','Ваш номер не может быть пустым')
                self.lenum.setText(WorkNumb)
            
            #сохраняем новое время выхода
            if shut_time != '' and len(shut_time) == 5:
                #если поле не пустое - записываем новое значение в файл
                module.write_setting(shut_time,22)
            elif (len(shut_time) < 5):
                QMessageBox.warning(self, 'Предупреждение','Время должно иметь формат:\n00:00')
                self.leshut.setText(WorkShut)
            else:
                #иначе, выводим предупреждение
                QMessageBox.warning(self, 'Предупреждение','Время выключения не может быть пустым')
                self.leshut.setText(WorkShut)
        
        #выводим предупреждение что новые настройки заработают после перезагрузки компа
        if self.flg_shutdown == 5577:
            self.flg_shutdown = 0   #после сохранения настроек сбрасываем признак
            reply = QMessageBox.question(self, 'Предупреждение',
                                         'Настройки вступят в силу после перезагрузки компьютера.\nПерезагрузить сейчас?',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            # если нажато "Да", перезагружаем систему
            if reply == QMessageBox.Yes:
                shutdown_lib.restart_system()

    
    #открываем папку с вашим файлом
    def opendirectory(self):
        path_file = module.read_setting(4) + '/' + module.read_setting(1)
        if os.name == 'nt':
            try:
                #os.startfile(os.path.dirname(path_file))    #открыть каталог с файлом
                os.startfile(path_file)                     #запуск файла
            except Exception as e:
                self.donot_open(e)
        else:
            try:
                opener = "open"
                subprocess.call([opener, path_file])
            except Exception as e:
                self.donot_open(e)
        
    #не могу открыть файл
    def donot_open(self, exception):
        QMessageBox.warning(self, 'Предупреждение','Не могу открыть файл\n' + str(exception))
    
    #открываем подсказку для выяснения номера на сайте
    def openhelp(self):
        #откроем дочернее окно с инструкцией
        self.w = AdjWindow()
        #self.w.show()
    
    #когда текст меняется, пишем ":" во второй символ
    def time_shutdow(self):
        text = self.leshut.text()
        if (len(text) >= 3):
            if text[2] != ':':
                text = text[:2] + ':' + text[2:]
                self.leshut.setText(text)
    
    #функция для перерасчета работы во всем Exel файле
    def exel_recount(self):
        #запускаем поток с отображение прогресса пересчета
        self.thread_progress.start()
    
                
    #если пересчет не удался показываем предупреждение
    def do_not_rec(self):
        #module.log_info("exel_recount exception: %s" % e)
        message = 'Ошибка пересчета файла\nПожалуйста проверьте, что:\n    - файл составлен корректно (имеет все названия месяцев, все времена прихода и ухода);\n    - файл закрыт.\nИ повторите попытку'
        reply = QMessageBox.warning(self, 'Ошибка', message, QMessageBox.Retry | QMessageBox.Cancel, QMessageBox.Retry)
        # если нажато "Повтор", запускаемся еще раз
        if reply == QMessageBox.Retry:
            self.exel_recount()
    
    #функция выключения окна прогресса пересчета
    def off_show_act(self):
        self.act.hide()

    #действия при выборе подсчета времени с сайта
    def webtime(self, state):
        chtext = 'Теперь время Вашего прихода и ухода берется с сайта.\nВаш компьютер автоматически запишет время Вашего ухода в файл и\nвыключится!'
        
        #если chekbox устанавили
        if state == Qt.Checked:
            #если изменился статус, высвечиваем сообщение
            if int(module.read_setting(16)) == 0:
                QMessageBox.warning(self, 'Предупреждение',chtext)
                self.flg_shutdown = 5577 #после изменения этой настройки предложим перезагрузить компьютер
            
            #module.write_setting('0',10)        #записываем в настроечный файл нулевое смещение
            #WorkOffset = int(module.read_setting(10))
            module.write_setting('0',13)        #записываем в настроечный файл нулевой выход
            WorkReload = int(module.read_setting(13))
            
            self.lenum.setEnabled(True)    #делаем строку ввода индивидуального номера активной
            self.leshut.setEnabled(True)        #делаем строку активной
            #self.spb.setEnabled(False)      #делаем виджет ввода смещеня неактивным
            self.spblered.setEnabled(False)    #делает виджет ввода возможного ухода неактивным
            
            #self.spb.setValue(WorkOffset)   #обнуляем смещение
            self.spblered.setValue(WorkReload)   #обнуляем reload
            
            #записываем в файл состояние виджета
            module.write_setting('1',16)
        #если checkbox сбросили
        else:
            #если изменился статусе
            if int(module.read_setting(16)) == 1:
                self.flg_shutdown = 5577 #после изменения этой настройки предложим перезагрузить компьютер
            self.lenum.setText(module.read_setting(19))
            self.lenum.setEnabled(False)    #делаем строку ввода индивидуального номера неактивной
            self.leshut.setText(module.read_setting(22))
            self.leshut.setEnabled(False)        #делаем строку неактивной
            #self.spb.setEnabled(True)      #делаем виджет ввода смещеня активным
            self.spblered.setEnabled(True)
            #записываем в файл состояние виджета
            module.write_setting('0',16)
    
    #метод заполнения диалогового окна из файла
    def RestoreFromFile(self):
        self.le.setText(os.path.abspath(str(module.read_setting(4) + '/' + module.read_setting(1))))
        self.spbO.setValue(int(module.read_setting(7)))
        self.spb.setValue(int(module.read_setting(10)))
        self.spblered.setValue(int(module.read_setting(13)))
        self.lenum.setText(module.read_setting(19))
        self.leshut.setText(module.read_setting(22))
        self.flg_shutdown = 0
    
    #наведение на иконку в трее
    def NowYourTime(self):
        day_time = []
        try:
            time = work_time.arr_time_day()    #получаем массив сегодняшних часов
        except:
            time = -1
        if time != -1 and len(time) > 0:
            time_start = str(time[0][0])    #время прихода
            #подменяем время ухода на текущее время по компу
            try:
                time[1][len(time[0])-1] = datetime.datetime.now().time().strftime("%H:%M")
            except:
                time[1].append(datetime.datetime.now().time().strftime("%H:%M"))
            day_time.append(time[0])
            day_time.append(time[1])
            sum_time = work_time.time_in_day(day_time[0], day_time[1]) #считаем количество часов в сегодняшнем дне
            self.tray_icon.setToolTip(f'Отслеживаю Ваше рабочее время.\nВы пришли: {time_start}\nОтработано: {sum_time}')
        else:
            self.tray_icon.setToolTip('Отслеживаю Ваше рабочее время.')
        #просто записываем в файл текущее время (так тратим меньше ресурсов и не надо "ловить" выключение компьютера)
        #получаем текущее время
        timeExit = datetime.datetime.now()
        #записываем текущее время в файл
        module.write_setting(timeExit.strftime("%d %m %Y %H:%M"), 25)
       
    # действие по нажатию на кнопку 'X'
    def closeEvent(self, event):
        #если путь в строке не совпадает с тем что записан в настроечном файле
        setting_dir_path = os.path.abspath(module.read_setting(4) + '/' + module.read_setting(1))
        setting_diner = int(module.read_setting(7))
        setting_offset = int(module.read_setting(10))
        setting_reload = int(module.read_setting(13))
        setting_number = str(module.read_setting(19))
        setting_shtime = str(module.read_setting(22))
        
        dir_path = os.path.abspath(self.le.text())
        work_diner = int(self.spbO.value())
        work_offset = int(self.spb.value())
        work_reload = int(self.spblered.value())
        you_number = str(self.lenum.text())
        shtime = str(self.leshut.text())
        
        if (dir_path != setting_dir_path) or (work_diner != setting_diner) or (work_offset !=
            setting_offset) or (you_number != setting_number) or (work_reload != setting_reload) or (shtime !=
            setting_shtime) or (self.flg_shutdown == 5577):
            
            reply = QMessageBox.question(self, 'Сообщение', "Вы хотите сохранить настройки?", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            # если нажато "Да", сохраняем файл подтверждаем закрытие
            if reply == QMessageBox.Yes:
                self.save_setting_btn()
            # если нет, то востанавливаем значения в виджете из файла
            else:
                self.RestoreFromFile()

        #сворачиваем приложение в Tray
        event.ignore()                          #игнорируем выход из программы
        self.hide()                             #скрываем программу

        self.tray_icon.showMessage(             #выводим сообщение
                "System Tray",
                "Программа свернута",
                QIcon('icon\Bill.ico'),
                1
            )
        #event.accept()                          #'''не забыть закоментировать!!!!'''
    
    #выход из программы
    def cleanUp(self):
        #сохраняем в Exel файл время выхода
        work_time.quit_app()

        #убираем иконку из Tray
        self.tray_icon.hide()
        
        #выключаю поток
        self.thread.quit()
        #сам выход
        sys.exit(0)

#создаем окно с подсказкой
class AdjWindow(QDialog):
   
    def __init__(self):
        super(AdjWindow, self).__init__()
                  
        self.resize(500,500)                                # Устанавливаем фиксированные размеры окна
        self.setWindowTitle("Как узнать свой индивидуальный номер")  # Устанавливаем заголовок окна
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_TitleBarContextHelpButton))   #устанавливаем одну из стандартных иконку
        
        #чтобы наше поле занимало все окно
        self.vbox = QVBoxLayout()
        #создаем поле с текстом инструкции и ссылкой
        self.pole_vivod = QTextBrowser(self)
        self.pole_vivod.setFont(QFont('Arial', 14))        #Шрифт
        self.pole_vivod.anchorClicked['QUrl'].connect(self.linkClicked)     #по клику открываем ссылку в браузере
        self.pole_vivod.setOpenLinks(False)     #Запрет удаления ссылки
        self.vbox.addWidget(self.pole_vivod)
        self.setLayout(self.vbox)
        
        try:
            text = module.read_help()
            self.pole_vivod.append(text)
            self.pole_vivod.moveCursor(QTextCursor.Start)
            self.show()
        except Exception as e:
            QMessageBox.warning(self, 'Предупреждение','Не могу открыть файл с инструкцией\n' + str(e))
        
    #обрабатываем клик по ссылке
    def linkClicked(self, url):
        webbrowser.open(url.toString()) 

#открываем наше окно
def app_main():
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.thread.start()
    # создадим таймер для подсчета и вывода текущего отработанного времени
    timer = QTimer()
    timer.timeout.connect(ex.NowYourTime)
    timer.start(30*1000)        #раз в 30 секунд
    #ex.show()                   #не забыть закоментировать
    sys.exit(app.exec_())
