# -*- coding: cp1251 -*-
'''
класс для вывода акна с настройками
'''
import os, sys, time
from work_setting import module, adjacent_classes
from work_lib import work_time, web_time

from PyQt5.QtWidgets import (QMainWindow, QPushButton, QLineEdit, QLabel, QDesktopWidget, QToolTip, QSystemTrayIcon,
    QMessageBox, QAction, QFileDialog, QApplication, qApp, QMenu, QSpinBox, QCheckBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt, QSize
from PyQt5.Qt import QIntValidator, QRegExp, QRegExpValidator

url_mars = 'http://www.mars/asu/report/enterexit/'

#создаем окно наших настроек
class MainWindow(QMainWindow):

    def __init__(self):
        # Метод super() возвращает объект родителя класса MainWindow и мы вызываем его конструктор.
        # Метод __init__() - это конструктор класса в языке Python.
        super().__init__()
        # Создание GUI поручено методу initUI().
        self.initUI()


    def initUI(self):
        #получаем путь к файлу и имя файла Exel из нашего настроечного файла
        WorkPath = module.read_path()
        WorkName = module.read_name()
        WorkOffset = module.read_offset()
        WorkReload = module.read_reload()
        YouNumber = module.read_number()
        CheckNum = module.read_check()
        WorkShut = module.read_timeShut()
        
        #окно стало неизменяемых размеров
        self.setFixedSize(QSize(495, 330))             # Устанавливаем фиксированные размеры окна
        self.setWindowTitle("Подсчёт рабочего времени")  # Устанавливаем заголовок окна
        self.setWindowIcon(QIcon('icon\Bill_chipher.jpg'))         # Устанавливаем иконку
        self.center()               # помещаем окно в центр экрана
        
        #создаем лейбел с именем файла
        self.lblN = QLabel(self)
        self.lblN.setFont(QFont('Arial', 12))        #Шрифт
        self.lblN.setText('Настройки файла ' + WorkName)
        self.lblN.move(10, 10)                      #расположение в окне
        self.lblN.adjustSize()                           #адаптивный размер в зависимости от содержимого
        
        #cоздаем строку с инструкцией
        self.lblI = QLabel(self)
        self.lblI.setFont(QFont('Arial', 12))        #Шрифт
        self.lblI.setText("Путь к файлу:")
        self.lblI.move(10, 50)                      #расположение в окне
        self.lblI.adjustSize()                           #адаптивный размер в зависимости от содержимого
        #self.lblI.resize(200, 30)
        
        #cоздаем строку с инструкцией для смещения
        self.lblS = QLabel(self)
        self.lblS.setFont(QFont('Arial', 12))        #Шрифт
        self.lblS.setText("Смещение:               мин.")
        self.lblS.move(10, 130)                      #расположение в окне
        self.lblS.adjustSize()                       #адаптивный размер в зависимости от содержимого
        '''
        #cоздаем строку с инструкцией для смещения (минуты)
        self.lblSm = QLabel(self)
        self.lblSm.setFont(QFont('Arial', 12))        #Шрифт
        self.lblSm.setText(" мин.")
        self.lblSm.move(150, 135)                      #расположение в окне
        self.lblSm.adjustSize()                       #адаптивный размер в зависимости от содержимого
        '''
        #cоздаем строку с инструкцией для времени безопасного ухода
        self.lblU = QLabel(self)
        self.lblU.setFont(QFont('Arial', 12))        #Шрифт
        self.lblU.setText("Возможный уход:               мин.")
        self.lblU.move(10, 162)                      #расположение в окне
        self.lblU.adjustSize()                       #адаптивный размер в зависимости от содержимого
          
        #cоздаем строку с инструкцией для индивидуального номера
        self.lblCh = QLabel(self)
        self.lblCh.setFont(QFont('Arial', 12))        #Шрифт
        self.lblCh.setText("Ваш номер на сайте:")
        self.lblCh.move(10, 231)                      #расположение в окне
        self.lblCh.adjustSize()                       #адаптивный размер в зависимости от содержимого
        
        #cоздаем строку с инструкцией для времени выключения
        self.lblSh = QLabel(self)
        self.lblSh.setFont(QFont('Arial', 12))        #Шрифт
        self.lblSh.setText("Выключать компьютер после:")
        self.lblSh.move(10, 262)                      #расположение в окне
        self.lblSh.adjustSize()                       #адаптивный размер в зависимости от содержимого
        
        #создаем строку для ввода пути к файлу
        self.le = QLineEdit(self)
        self.le.setFont(QFont('Arial', 12))         #Шрифт
        self.le.move(10, 72)                        #расположение в окне 
        self.le.resize(360,26)                      #размер строки 
        self.le.setText(WorkPath + '/' + WorkName)  #пишем путь из настроечного файла
        #self.le.returnPressed.connect(self.btn_save.click)  # click on <Enter>
        
        #создаем строку для ввода индивидуального номера сотрудника
        self.lenum = QLineEdit(self)
        self.lenum.setFont(QFont('Arial', 12))         #Шрифт
        self.lenum.move(170, 229)                        #расположение в окне 
        self.lenum.resize(45,26)                      #размер строки
        self.lenum.setValidator(QIntValidator(0,9999))
        self.lenum.setText(YouNumber)          #пишем путь из настроечного файла
        self.lenum.returnPressed.connect(self.save_setting_btn) # click on <Enter>
        self.lenum.setEnabled(False)        #делаем строку неактивной
        
        #создаем Валидатор для строки времени
        hour = '(2[0123]|([0-1][0-9]))'
        minute = '[0-5][0-9]'
        timeRange = QRegExp('^' + hour + ':' + minute + '$')
        timeVali = QRegExpValidator(timeRange, self)
        #создаем строку для ввода выключения компьютера
        self.leshut = QLineEdit(self)
        self.leshut.setFont(QFont('Arial', 12))         #Шрифт
        self.leshut.move(235, 260)                        #расположение в окне 
        self.leshut.resize(50,26)                      #размер строки
        self.leshut.setText(str(WorkShut))          #пишем путь из настроечного файла
        self.leshut.setValidator(timeVali)
        #self.leshut.setInputMask('99:99')
        self.leshut.returnPressed.connect(self.save_setting_btn)    # click on <Enter>
        self.leshut.setEnabled(False)        #делаем строку неактивной
        
        #создаем кнопку для изменения расположения файла
        self.btnI = QPushButton('Изменить', self)
        self.btnI.setFont(QFont('Arial', 12))        #Шрифт
        self.btnI.move(385, 70)                      #расположение в окне кнопки
        self.btnI.clicked.connect(self.getfile)      #действие по нажатию
        self.btnI.setAutoDefault(True)               # click on <Enter>
        #будет только кнопка изменения файла

        #создаем кнопку для открытия директории/файла
        self.btnO = QPushButton('Открыть', self)
        self.btnO.setFont(QFont('Arial', 12))        #Шрифт
        self.btnO.move(385, 100)                      #расположение в окне кнопки
        self.btnO.clicked.connect(self.opendirectory)      #действие по нажатию
        self.btnO.setAutoDefault(True)               # click on <Enter>
        
        #создаем кнопку для подсказки
        self.btnch = QPushButton('?', self)
        self.btnch.setFont(QFont('Arial', 18))        #Шрифт
        self.btnch.move(235, 200)                      #расположение в окне кнопки
        self.btnch.resize(20, 26)
        try:
            self.btnch.clicked.connect(self.openhelp)      #действие по нажатию
        except Exception as e:
            module.log_info('Error openhelp: %s' %e)
        self.btnch.setAutoDefault(True)               # click on <Enter>
        
        #создаем SpinBox для выбора времени
        self.spb = QSpinBox(self)
        self.spb.setFont(QFont('Arial', 12))        #Шрифт
        self.spb.move(100, 128)                     #расположение в окне кнопки
        self.spb.resize(45, 25)                     #размер
        self.spb.setMaximum(60)                     #верхняя граница счетчика
        self.spb.setMinimum(0)                      #нижняя граница счетчика
        self.spb.setValue(WorkOffset)
        
        #создаем SpinBox для выбора времени ухода
        self.spblered = QSpinBox(self)
        self.spblered.setFont(QFont('Arial', 12))        #Шрифт
        self.spblered.move(145, 160)                     #расположение в окне кнопки
        self.spblered.resize(45, 25)                     #размер
        self.spblered.setMaximum(60)                     #верхняя граница счетчика
        self.spblered.setMinimum(0)                      #нижняя граница счетчика
        self.spblered.setValue(WorkReload)
        
        #создаем checkbox для выбора подчета времени с сайта Марса
        self.chweb = QCheckBox('Брать время с сайта Марса', self)
        self.chweb.setFont(QFont('Arial', 12))          #Шрифт
        self.chweb.move (10, 200)
        self.chweb.adjustSize()                           #адаптивный размер в зависимости от содержимого
        self.chweb.stateChanged.connect(self.webtime)           #действие по нажатию
        #выставляем в соответствии с настройками
        if(CheckNum):
            self.chweb.setChecked(True)
        else:
            self.chweb.setChecked(False)
        
        #создаем кнопку для всплывающего диалогового окна
        self.btn_save = QPushButton('Сохранить', self)
        self.btn_save.setFont(QFont('Arial', 12))        #Шрифт
        self.btn_save.move(385, 290)                      #расположение в окне кнопки
        self.btn_save.clicked.connect(self.save_setting_btn)      #действие по нажатию
        self.btn_save.setAutoDefault(True)               # click on <Enter>
        
        self.le.returnPressed.connect(self.btn_save.click)  #действия в строке по интеру
        self.lenum.returnPressed.connect(self.btn_save.click)  #действия в строке по интеру
        #self.spb.returnPressed.connect(self.btn_save.click)    #действия в SpinBox по интеру
        
        # Инициализируем иконку Tray
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon('icon\Bill_chipher.jpg')) #устанавливаем пользовательскую иконку
        #self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))   #устанавливаем одну из стандартных иконку
 
        '''
            Объявим и добавим действия для работы с иконкой системного трея
            show - показать окно
            exit - выход из программы
        '''
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
        self.le.setToolTip('Файл с Вашим рабочим временем находится по этому пути\n' + WorkPath + '/' + WorkName)
        self.lblI.setToolTip('Файл с Вашим рабочим временем находится по этому пути')
        self.btnI.setToolTip('Выберете файл рабочего времени')    # создаем подсказку для кнопки
        #self.btnS.setToolTip('Создать новый файл рабочего времени')    # создаем подсказку для кнопки
        self.btnO.setToolTip('Открыть папку с файлом')    # создаем подсказку для кнопки
        self.btn_save.setToolTip('Сохранить выставленные настройки')    # создаем подсказку для кнопки
        self.spb.setToolTip('Это смещение от времени вкл/выкл ПК')    # создаем подсказку
        self.lblS.setToolTip('Сколько Вам идти от КПП до рабочего места?')    # создаем подсказку для кнопки
        #self.lblSm.setToolTip('Укажите смещение в минутах')
        self.lblU.setToolTip('Если компьютер выключится на заданное время,\n то в файле Вашего рабочего времени уход не зафиксируется')
        self.spblered.setToolTip('Введите время неучетного выхода в мин.')
        self.lblCh.setToolTip('Введите вашь номер на сайте')
        self.chweb.setToolTip('Время Вашего прихода фиксируется на сайте:\n' + url_mars + '\n брать время Вашего прихода от туда?')
        self.lenum.setToolTip('Вашь номер на сайте: ' + YouNumber)
        self.btnch.setToolTip('Как узнать свой номер на сайте?')
        self.lblSh.setToolTip('Если ваше время выхода на КПП после этого времени, выключаю компьютер')
        self.leshut.setToolTip('Введите время в формате:\n00:00')
        self.tray_icon.setToolTip('Отслеживаю Ваше рабочее время')
        #self.show()    #показываем окно/показывать будем в основном модуле

    #диалоговое окно выбора нового файла
    def getfile(self):
        dir_path = module.read_path() + '/' + module.read_name()
        fname = QFileDialog.getOpenFileName(self, 'Выбрать файл', dir_path, 'Exel files (*.xls)')
        #если новый файл выбран, переписываем путь в настройках и в наших текстовых виджетах
        if fname != ('', ''):
            new_dir = os.path.dirname(fname[0])    #путь папки в которой лежит файл
            new_name = os.path.basename(fname[0])   #имя файла
            #запись в настроечный файл нового имени файла
            module.write_name(new_name)
            #запись в настроечный файл нового пути
            module.write_path(new_dir)
            self.le.setText(fname[0])
            self.lblN.setText('Настройки файла ' + new_name)
        
    #диалоговое окно сохранения нового файла
    def savefile(self):
        dir_path = module.read_path() + '/' + module.read_name()
        fname = QFileDialog.getSaveFileName(self, 'Выбрать файл', dir_path, 'Exel files (*.xls)')
        #если новый файл выбран, переписываем путь в настройках и в наших текстовых виджетах
        if fname != ('', ''):
            new_dir = module.save_setting(fname[0],'Repace')
            self.le.setText(new_dir[0] + '/' + new_dir[1])
            self.lblN.setText('Настройки файла ' + new_dir[1])
    
    #сохраняем настройки с учетом того что введено в строку
    def save_setting_btn(self):
        #получаем путь к файлу, имя файла Exel и номер пользователя из нашего настроечного файла
        WorkPath = module.read_path()
        WorkName = module.read_name()
        WorkNumb = module.read_number()
        WorkShut = module.read_timeShut()

        dir_path = self.le.text()   #получаем путь к новому файлу
        you_numb = self.lenum.text()
        shut_time = self.leshut.text()
        
        #сохраняем новй путь Exel файла в настроечный файл
        if dir_path != '':
            #если выбранный файл существует записываем в настройки путь к нему
            if os.path.exists(dir_path):
                #запись в настроечный файл нового имени файла
                module.write_name(os.path.basename(dir_path))
                #запись в настроечный файл нового пути
                module.write_path(os.path.dirname(dir_path))
            #если файла нет, тосздать его?
            else:
                reply = QMessageBox.question(self, 'Сообщение', 'Файл не найден.\nСоздать новый файл?', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                #QMessageBox.warning(self, 'Предупреждение','Файл не найден.\nСоздать новый файл?')
                # если нажато "Да", создаем файл и сохраняем его путь в настроечный файл
                if reply == QMessageBox.Yes:
                    module.new_timework_file(dir_path)
                    #запись в настроечный файл нового имени файла
                    module.write_name(os.path.basename(dir_path))
                    #запись в настроечный файл нового пути
                    module.write_path(os.path.dirname(dir_path))
                #self.le.setText(WorkPath + '/' + WorkName)    #вернуть исходное значение пути
        #если путь пустой выводим сообщение
        else:
            QMessageBox.warning(self, 'Предупреждение','Путь к файлу не может быть пустым')
            self.le.setText(WorkPath + '/' + WorkName)
        '''
        мы теперь ничего не сохраняем Просто выбираем другой файл
        #получаем новое значение из строки
        dir_path = self.le.text()
        #если строка не пустая
        if dir_path != '':
            new_dir = module.save_setting(dir_path)             #редактируем настроечный файл
            self.le.setText(new_dir[0] + '/' + new_dir[1])      #обновляем строку
            self.lblN.setText('Настройки файла ' + new_dir[1])  #обновляем лейбл
            self.lblN.adjustSize()                              #обновляем размер лейбла
        #если строка пустая выводится предупреждающее сообщение и возвращается текст
        else:
            QMessageBox.warning(self, 'Предупреждение','Путь к файлу не может быть пустым')
            self.le.setText(WorkPath + '/' + WorkName)
        '''
        
        #сохраняем значение из SpinBox в файл
        module.write_offset(self.spb.value())
        module.write_reload(self.spblered.value())
        
        #сохраняем новый номер пользователя в файл
        if you_numb != '':
            #если поле не пустое - записываем новое значение в файл
            module.write_number(you_numb)
        else:
            #иначе, выводим предупреждение
            QMessageBox.warning(self, 'Предупреждение','Ваш номер не может быть пустым')
            self.lenum.setText(WorkNumb)
        
        #сохраняем новое время выхода
        print(len(shut_time))
        if shut_time != '' and len(shut_time) == 5:
            #если поле не пустое - записываем новое значение в файл
            module.write_timeShut(shut_time)
        elif (len(shut_time) < 5):
            QMessageBox.warning(self, 'Предупреждение','Время должно иметь формат:\n00:00')
            self.leshut.setText(WorkShut)
        else:
            #иначе, выводим предупреждение
            QMessageBox.warning(self, 'Предупреждение','Время выключения не может быть пустым')
            self.leshut.setText(WorkShut)
        
        #выводим предупреждение что новые настройки заработают после перезагрузки компа
        #QMessageBox.warning(self, 'Предупреждение','Некоторые настройки вступят в силу после перезапуска программы')
    
    #открываем папку с вашим файлом
    def opendirectory(self):
        path_file = module.read_path() + '/' + module.read_name()
        os.startfile(os.path.dirname(path_file))    #открыть каталог с файлом
        os.startfile(path_file)                     #запуск файла
        
    #открываем подсказку для выяснения номера на сайте
    def openhelp(self):
        #откроем дочернее окно м инструкцией
        self.w = adjacent_classes.AdjWindow()
        self.tw = adjacent_classes.ShutWindow()
        self.tw.show()
        
    #Функция для центрирования окна в экране пользователя
    def center(self):
        qr = self.frameGeometry()           # получаем прямоугольник, точно определяющий форму главного окна.
        cp = QDesktopWidget().availableGeometry().center()  # выясняем разрешение экрана нашего монитора. Из этого разрешения, мы получаем центральную точку.
        qr.moveCenter(cp)                   # устанавливаем центр прямоугольника в центр экрана. Размер прямоугольника не изменяется.
        self.move(qr.topLeft())             # перемещаем верхнюю левую точку окна приложения в верхнюю левую точку прямоугольника qr, таким образом центрируя окно на нашем экране.
    
    #действия при выборе подсчета времени с сайта
    def webtime(self, state):
        chtext = 'Теперь время вашего прихода и ухода берется с сайта.\nВаш компьютер автоматически запишет время вашего ухода в файл и\nвыключится!'
        
        #если chekbox устанавили
        if state == Qt.Checked:
            #если изменился статус, высвечиваем сообщение
            if module.read_check() == 0:
                QMessageBox.warning(self, 'Предупреждение',chtext)
            
            #module.write_offset('0')        #записываем в настроечный файл нулевое смещение
            WorkOffset = module.read_offset()
            #module.write_reload('0')        #записываем в настроечный файл нулевой выход
            WorkReload = module.read_reload()
            
            self.lenum.setEnabled(True)    #делаем строку ввода индивидуального номера активной
            self.leshut.setEnabled(True)        #делаем строку активной
            self.spb.setEnabled(False)      #делаем виджет ввода смещеня неактивным
            self.spblered.setEnabled(False)    #делает виджет ввода возможного ухода неактивным
            
            self.spb.setValue(WorkOffset)   #обнуляем смещение
            self.spblered.setValue(WorkReload)   #обнуляем reload
            
            #записываем в файл состояние виджета
            module.write_checkt('1')
        #если checkbox сбросили
        else:
            self.lenum.setEnabled(False)    #делаем строку ввода индивидуального номера неактивной
            self.leshut.setEnabled(False)        #делаем строку неактивной
            self.spb.setEnabled(True)      #делаем виджет ввода смещеня активным
            self.spblered.setEnabled(True)
            #записываем в файл состояние виджета
            module.write_checkt('0')
        
    # действие по нажатию на кнопку 'X'
    def closeEvent(self, event):
        # показываем сообщение с двумя кнопками: «Yes» и «No».
        # Первая строка появляется в строке заголовка. Вторая строка – это текст сообщения, отображаемый с помощью диалогового окна.
        # Третий аргумент указывает комбинацию кнопок, появляющихся в диалоге. Последний параметр – кнопка по умолчанию.
        # Это кнопка, которая первоначально имеет на себе указатель клавиатуры.
        # Возвращаемое значение хранится в переменной reply.
        
        #если путь в строке не совпадает с тем что записан в настроечном файле
        setting_dir_path = module.read_path() + '/' + module.read_name()
        setting_offset = module.read_offset()
        setting_reload = module.read_reload()
        setting_number = module.read_number()
        
        dir_path = self.le.text()
        work_offset = self.spb.value()
        work_reload = self.spblered.value()
        you_number = self.lenum.text()
        
        if (dir_path != setting_dir_path) or (work_offset != setting_offset) or (you_number != setting_number) or (work_reload != setting_reload):
            reply = QMessageBox.question(self, 'Сообщение', "Вы хотите сохранить настройки?", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            # если нажато "Да", сохраняем файл подтверждаем закрыти
            if reply == QMessageBox.Yes:
                self.save_setting_btn()
        #сворачиваем приложение в Tray
        event.ignore()                          #игнорируем выход из программы
        self.hide()                             #скрываем программу
        self.tray_icon.showMessage(             #выводим сообщение
                "System Tray",
                "Программа свернута",
                QIcon('icon\Bill.jpg'),
                1
            )
        event.accept()                          #'''не забыть закоментировать!!!!'''
    
    #выход из программы
    @staticmethod
    def cleanUp(self):
    #def work_exit(self):
        #записываю в лог файл
        module.log_info('Выключаюсь!!!')
        #сохраняем в Exel файл время выхода
        work_time.quit_app()
        #убираем иконку из Tray
        self.tray_icon.hide()
        #сам выход
        qApp.quit()
        
#открываем наше окно
#if __name__ == '__main__':
def app_main():
    app = QApplication(sys.argv)
    ex = MainWindow()
    app.aboutToQuit.connect(ex.cleanUp)
    ex.show()                   #не забыть закоментировать
    
    sys.exit(app.exec_())
