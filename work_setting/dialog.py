# -*- coding: cp1251 -*-
'''
����� ��� ������ ���� � �����������
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

#������� ���� ����� ��������
class MainWindow(QMainWindow):

    def __init__(self):
        # ����� super() ���������� ������ �������� ������ MainWindow � �� �������� ��� �����������.
        # ����� __init__() - ��� ����������� ������ � ����� Python.
        super().__init__()
        # �������� GUI �������� ������ initUI().
        self.initUI()


    def initUI(self):
        #�������� ���� � ����� � ��� ����� Exel �� ������ ������������ �����
        WorkPath = module.read_path()
        WorkName = module.read_name()
        WorkOffset = module.read_offset()
        WorkReload = module.read_reload()
        YouNumber = module.read_number()
        CheckNum = module.read_check()
        WorkShut = module.read_timeShut()
        
        #���� ����� ������������ ��������
        self.setFixedSize(QSize(495, 330))             # ������������� ������������� ������� ����
        self.setWindowTitle("������� �������� �������")  # ������������� ��������� ����
        self.setWindowIcon(QIcon('icon\Bill_chipher.jpg'))         # ������������� ������
        self.center()               # �������� ���� � ����� ������
        
        #������� ������ � ������ �����
        self.lblN = QLabel(self)
        self.lblN.setFont(QFont('Arial', 12))        #�����
        self.lblN.setText('��������� ����� ' + WorkName)
        self.lblN.move(10, 10)                      #������������ � ����
        self.lblN.adjustSize()                           #���������� ������ � ����������� �� �����������
        
        #c������ ������ � �����������
        self.lblI = QLabel(self)
        self.lblI.setFont(QFont('Arial', 12))        #�����
        self.lblI.setText("���� � �����:")
        self.lblI.move(10, 50)                      #������������ � ����
        self.lblI.adjustSize()                           #���������� ������ � ����������� �� �����������
        #self.lblI.resize(200, 30)
        
        #c������ ������ � ����������� ��� ��������
        self.lblS = QLabel(self)
        self.lblS.setFont(QFont('Arial', 12))        #�����
        self.lblS.setText("��������:               ���.")
        self.lblS.move(10, 130)                      #������������ � ����
        self.lblS.adjustSize()                       #���������� ������ � ����������� �� �����������
        '''
        #c������ ������ � ����������� ��� �������� (������)
        self.lblSm = QLabel(self)
        self.lblSm.setFont(QFont('Arial', 12))        #�����
        self.lblSm.setText(" ���.")
        self.lblSm.move(150, 135)                      #������������ � ����
        self.lblSm.adjustSize()                       #���������� ������ � ����������� �� �����������
        '''
        #c������ ������ � ����������� ��� ������� ����������� �����
        self.lblU = QLabel(self)
        self.lblU.setFont(QFont('Arial', 12))        #�����
        self.lblU.setText("��������� ����:               ���.")
        self.lblU.move(10, 162)                      #������������ � ����
        self.lblU.adjustSize()                       #���������� ������ � ����������� �� �����������
          
        #c������ ������ � ����������� ��� ��������������� ������
        self.lblCh = QLabel(self)
        self.lblCh.setFont(QFont('Arial', 12))        #�����
        self.lblCh.setText("��� ����� �� �����:")
        self.lblCh.move(10, 231)                      #������������ � ����
        self.lblCh.adjustSize()                       #���������� ������ � ����������� �� �����������
        
        #c������ ������ � ����������� ��� ������� ����������
        self.lblSh = QLabel(self)
        self.lblSh.setFont(QFont('Arial', 12))        #�����
        self.lblSh.setText("��������� ��������� �����:")
        self.lblSh.move(10, 262)                      #������������ � ����
        self.lblSh.adjustSize()                       #���������� ������ � ����������� �� �����������
        
        #������� ������ ��� ����� ���� � �����
        self.le = QLineEdit(self)
        self.le.setFont(QFont('Arial', 12))         #�����
        self.le.move(10, 72)                        #������������ � ���� 
        self.le.resize(360,26)                      #������ ������ 
        self.le.setText(WorkPath + '/' + WorkName)  #����� ���� �� ������������ �����
        #self.le.returnPressed.connect(self.btn_save.click)  # click on <Enter>
        
        #������� ������ ��� ����� ��������������� ������ ����������
        self.lenum = QLineEdit(self)
        self.lenum.setFont(QFont('Arial', 12))         #�����
        self.lenum.move(170, 229)                        #������������ � ���� 
        self.lenum.resize(45,26)                      #������ ������
        self.lenum.setValidator(QIntValidator(0,9999))
        self.lenum.setText(YouNumber)          #����� ���� �� ������������ �����
        self.lenum.returnPressed.connect(self.save_setting_btn) # click on <Enter>
        self.lenum.setEnabled(False)        #������ ������ ����������
        
        #������� ��������� ��� ������ �������
        hour = '(2[0123]|([0-1][0-9]))'
        minute = '[0-5][0-9]'
        timeRange = QRegExp('^' + hour + ':' + minute + '$')
        timeVali = QRegExpValidator(timeRange, self)
        #������� ������ ��� ����� ���������� ����������
        self.leshut = QLineEdit(self)
        self.leshut.setFont(QFont('Arial', 12))         #�����
        self.leshut.move(235, 260)                        #������������ � ���� 
        self.leshut.resize(50,26)                      #������ ������
        self.leshut.setText(str(WorkShut))          #����� ���� �� ������������ �����
        self.leshut.setValidator(timeVali)
        #self.leshut.setInputMask('99:99')
        self.leshut.returnPressed.connect(self.save_setting_btn)    # click on <Enter>
        self.leshut.setEnabled(False)        #������ ������ ����������
        
        #������� ������ ��� ��������� ������������ �����
        self.btnI = QPushButton('��������', self)
        self.btnI.setFont(QFont('Arial', 12))        #�����
        self.btnI.move(385, 70)                      #������������ � ���� ������
        self.btnI.clicked.connect(self.getfile)      #�������� �� �������
        self.btnI.setAutoDefault(True)               # click on <Enter>
        #����� ������ ������ ��������� �����

        #������� ������ ��� �������� ����������/�����
        self.btnO = QPushButton('�������', self)
        self.btnO.setFont(QFont('Arial', 12))        #�����
        self.btnO.move(385, 100)                      #������������ � ���� ������
        self.btnO.clicked.connect(self.opendirectory)      #�������� �� �������
        self.btnO.setAutoDefault(True)               # click on <Enter>
        
        #������� ������ ��� ���������
        self.btnch = QPushButton('?', self)
        self.btnch.setFont(QFont('Arial', 18))        #�����
        self.btnch.move(235, 200)                      #������������ � ���� ������
        self.btnch.resize(20, 26)
        try:
            self.btnch.clicked.connect(self.openhelp)      #�������� �� �������
        except Exception as e:
            module.log_info('Error openhelp: %s' %e)
        self.btnch.setAutoDefault(True)               # click on <Enter>
        
        #������� SpinBox ��� ������ �������
        self.spb = QSpinBox(self)
        self.spb.setFont(QFont('Arial', 12))        #�����
        self.spb.move(100, 128)                     #������������ � ���� ������
        self.spb.resize(45, 25)                     #������
        self.spb.setMaximum(60)                     #������� ������� ��������
        self.spb.setMinimum(0)                      #������ ������� ��������
        self.spb.setValue(WorkOffset)
        
        #������� SpinBox ��� ������ ������� �����
        self.spblered = QSpinBox(self)
        self.spblered.setFont(QFont('Arial', 12))        #�����
        self.spblered.move(145, 160)                     #������������ � ���� ������
        self.spblered.resize(45, 25)                     #������
        self.spblered.setMaximum(60)                     #������� ������� ��������
        self.spblered.setMinimum(0)                      #������ ������� ��������
        self.spblered.setValue(WorkReload)
        
        #������� checkbox ��� ������ ������� ������� � ����� �����
        self.chweb = QCheckBox('����� ����� � ����� �����', self)
        self.chweb.setFont(QFont('Arial', 12))          #�����
        self.chweb.move (10, 200)
        self.chweb.adjustSize()                           #���������� ������ � ����������� �� �����������
        self.chweb.stateChanged.connect(self.webtime)           #�������� �� �������
        #���������� � ������������ � �����������
        if(CheckNum):
            self.chweb.setChecked(True)
        else:
            self.chweb.setChecked(False)
        
        #������� ������ ��� ������������ ����������� ����
        self.btn_save = QPushButton('���������', self)
        self.btn_save.setFont(QFont('Arial', 12))        #�����
        self.btn_save.move(385, 290)                      #������������ � ���� ������
        self.btn_save.clicked.connect(self.save_setting_btn)      #�������� �� �������
        self.btn_save.setAutoDefault(True)               # click on <Enter>
        
        self.le.returnPressed.connect(self.btn_save.click)  #�������� � ������ �� ������
        self.lenum.returnPressed.connect(self.btn_save.click)  #�������� � ������ �� ������
        #self.spb.returnPressed.connect(self.btn_save.click)    #�������� � SpinBox �� ������
        
        # �������������� ������ Tray
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon('icon\Bill_chipher.jpg')) #������������� ���������������� ������
        #self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))   #������������� ���� �� ����������� ������
 
        '''
            ������� � ������� �������� ��� ������ � ������� ���������� ����
            show - �������� ����
            exit - ����� �� ���������
        '''
        show_action = QAction(QIcon('icon\Programming-Show.png'), "���������", self)
        quit_action = QAction(QIcon('icon\exit.png'), "�����", self)
        show_action.triggered.connect(self.show)        #��� ������� �� show ���� �����������
        quit_action.triggered.connect(self.cleanUp)        #��� ������� �� quit ���������� ����������� qApp.quit
        tray_menu = QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
        #������� ���������
        QToolTip.setFont(QFont('Arial', 10))    # ����� ������������� �����, ������������ ��� ������ ����������� ���������.
        self.setToolTip('��� ���� ������ �������� �������� ���������')  # ������� ��������� ��� ����
        self.lblN.setToolTip('������� ��� �����, � ������� �������� ���� ������� �����')
        self.le.setToolTip('���� � ����� ������� �������� ��������� �� ����� ����\n' + WorkPath + '/' + WorkName)
        self.lblI.setToolTip('���� � ����� ������� �������� ��������� �� ����� ����')
        self.btnI.setToolTip('�������� ���� �������� �������')    # ������� ��������� ��� ������
        #self.btnS.setToolTip('������� ����� ���� �������� �������')    # ������� ��������� ��� ������
        self.btnO.setToolTip('������� ����� � ������')    # ������� ��������� ��� ������
        self.btn_save.setToolTip('��������� ������������ ���������')    # ������� ��������� ��� ������
        self.spb.setToolTip('��� �������� �� ������� ���/���� ��')    # ������� ���������
        self.lblS.setToolTip('������� ��� ���� �� ��� �� �������� �����?')    # ������� ��������� ��� ������
        #self.lblSm.setToolTip('������� �������� � �������')
        self.lblU.setToolTip('���� ��������� ���������� �� �������� �����,\n �� � ����� ������ �������� ������� ���� �� �������������')
        self.spblered.setToolTip('������� ����� ���������� ������ � ���.')
        self.lblCh.setToolTip('������� ���� ����� �� �����')
        self.chweb.setToolTip('����� ������ ������� ����������� �� �����:\n' + url_mars + '\n ����� ����� ������ ������� �� ����?')
        self.lenum.setToolTip('���� ����� �� �����: ' + YouNumber)
        self.btnch.setToolTip('��� ������ ���� ����� �� �����?')
        self.lblSh.setToolTip('���� ���� ����� ������ �� ��� ����� ����� �������, �������� ���������')
        self.leshut.setToolTip('������� ����� � �������:\n00:00')
        self.tray_icon.setToolTip('���������� ���� ������� �����')
        #self.show()    #���������� ����/���������� ����� � �������� ������

    #���������� ���� ������ ������ �����
    def getfile(self):
        dir_path = module.read_path() + '/' + module.read_name()
        fname = QFileDialog.getOpenFileName(self, '������� ����', dir_path, 'Exel files (*.xls)')
        #���� ����� ���� ������, ������������ ���� � ���������� � � ����� ��������� ��������
        if fname != ('', ''):
            new_dir = os.path.dirname(fname[0])    #���� ����� � ������� ����� ����
            new_name = os.path.basename(fname[0])   #��� �����
            #������ � ����������� ���� ������ ����� �����
            module.write_name(new_name)
            #������ � ����������� ���� ������ ����
            module.write_path(new_dir)
            self.le.setText(fname[0])
            self.lblN.setText('��������� ����� ' + new_name)
        
    #���������� ���� ���������� ������ �����
    def savefile(self):
        dir_path = module.read_path() + '/' + module.read_name()
        fname = QFileDialog.getSaveFileName(self, '������� ����', dir_path, 'Exel files (*.xls)')
        #���� ����� ���� ������, ������������ ���� � ���������� � � ����� ��������� ��������
        if fname != ('', ''):
            new_dir = module.save_setting(fname[0],'Repace')
            self.le.setText(new_dir[0] + '/' + new_dir[1])
            self.lblN.setText('��������� ����� ' + new_dir[1])
    
    #��������� ��������� � ������ ���� ��� ������� � ������
    def save_setting_btn(self):
        #�������� ���� � �����, ��� ����� Exel � ����� ������������ �� ������ ������������ �����
        WorkPath = module.read_path()
        WorkName = module.read_name()
        WorkNumb = module.read_number()
        WorkShut = module.read_timeShut()

        dir_path = self.le.text()   #�������� ���� � ������ �����
        you_numb = self.lenum.text()
        shut_time = self.leshut.text()
        
        #��������� ���� ���� Exel ����� � ����������� ����
        if dir_path != '':
            #���� ��������� ���� ���������� ���������� � ��������� ���� � ����
            if os.path.exists(dir_path):
                #������ � ����������� ���� ������ ����� �����
                module.write_name(os.path.basename(dir_path))
                #������ � ����������� ���� ������ ����
                module.write_path(os.path.dirname(dir_path))
            #���� ����� ���, �������� ���?
            else:
                reply = QMessageBox.question(self, '���������', '���� �� ������.\n������� ����� ����?', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                #QMessageBox.warning(self, '��������������','���� �� ������.\n������� ����� ����?')
                # ���� ������ "��", ������� ���� � ��������� ��� ���� � ����������� ����
                if reply == QMessageBox.Yes:
                    module.new_timework_file(dir_path)
                    #������ � ����������� ���� ������ ����� �����
                    module.write_name(os.path.basename(dir_path))
                    #������ � ����������� ���� ������ ����
                    module.write_path(os.path.dirname(dir_path))
                #self.le.setText(WorkPath + '/' + WorkName)    #������� �������� �������� ����
        #���� ���� ������ ������� ���������
        else:
            QMessageBox.warning(self, '��������������','���� � ����� �� ����� ���� ������')
            self.le.setText(WorkPath + '/' + WorkName)
        '''
        �� ������ ������ �� ��������� ������ �������� ������ ����
        #�������� ����� �������� �� ������
        dir_path = self.le.text()
        #���� ������ �� ������
        if dir_path != '':
            new_dir = module.save_setting(dir_path)             #����������� ����������� ����
            self.le.setText(new_dir[0] + '/' + new_dir[1])      #��������� ������
            self.lblN.setText('��������� ����� ' + new_dir[1])  #��������� �����
            self.lblN.adjustSize()                              #��������� ������ ������
        #���� ������ ������ ��������� ��������������� ��������� � ������������ �����
        else:
            QMessageBox.warning(self, '��������������','���� � ����� �� ����� ���� ������')
            self.le.setText(WorkPath + '/' + WorkName)
        '''
        
        #��������� �������� �� SpinBox � ����
        module.write_offset(self.spb.value())
        module.write_reload(self.spblered.value())
        
        #��������� ����� ����� ������������ � ����
        if you_numb != '':
            #���� ���� �� ������ - ���������� ����� �������� � ����
            module.write_number(you_numb)
        else:
            #�����, ������� ��������������
            QMessageBox.warning(self, '��������������','��� ����� �� ����� ���� ������')
            self.lenum.setText(WorkNumb)
        
        #��������� ����� ����� ������
        print(len(shut_time))
        if shut_time != '' and len(shut_time) == 5:
            #���� ���� �� ������ - ���������� ����� �������� � ����
            module.write_timeShut(shut_time)
        elif (len(shut_time) < 5):
            QMessageBox.warning(self, '��������������','����� ������ ����� ������:\n00:00')
            self.leshut.setText(WorkShut)
        else:
            #�����, ������� ��������������
            QMessageBox.warning(self, '��������������','����� ���������� �� ����� ���� ������')
            self.leshut.setText(WorkShut)
        
        #������� �������������� ��� ����� ��������� ���������� ����� ������������ �����
        #QMessageBox.warning(self, '��������������','��������� ��������� ������� � ���� ����� ����������� ���������')
    
    #��������� ����� � ����� ������
    def opendirectory(self):
        path_file = module.read_path() + '/' + module.read_name()
        os.startfile(os.path.dirname(path_file))    #������� ������� � ������
        os.startfile(path_file)                     #������ �����
        
    #��������� ��������� ��� ��������� ������ �� �����
    def openhelp(self):
        #������� �������� ���� � �����������
        self.w = adjacent_classes.AdjWindow()
        self.tw = adjacent_classes.ShutWindow()
        self.tw.show()
        
    #������� ��� ������������� ���� � ������ ������������
    def center(self):
        qr = self.frameGeometry()           # �������� �������������, ����� ������������ ����� �������� ����.
        cp = QDesktopWidget().availableGeometry().center()  # �������� ���������� ������ ������ ��������. �� ����� ����������, �� �������� ����������� �����.
        qr.moveCenter(cp)                   # ������������� ����� �������������� � ����� ������. ������ �������������� �� ����������.
        self.move(qr.topLeft())             # ���������� ������� ����� ����� ���� ���������� � ������� ����� ����� �������������� qr, ����� ������� ��������� ���� �� ����� ������.
    
    #�������� ��� ������ �������� ������� � �����
    def webtime(self, state):
        chtext = '������ ����� ������ ������� � ����� ������� � �����.\n��� ��������� ������������� ������� ����� ������ ����� � ���� �\n����������!'
        
        #���� chekbox ����������
        if state == Qt.Checked:
            #���� ��������� ������, ����������� ���������
            if module.read_check() == 0:
                QMessageBox.warning(self, '��������������',chtext)
            
            #module.write_offset('0')        #���������� � ����������� ���� ������� ��������
            WorkOffset = module.read_offset()
            #module.write_reload('0')        #���������� � ����������� ���� ������� �����
            WorkReload = module.read_reload()
            
            self.lenum.setEnabled(True)    #������ ������ ����� ��������������� ������ ��������
            self.leshut.setEnabled(True)        #������ ������ ��������
            self.spb.setEnabled(False)      #������ ������ ����� ������� ����������
            self.spblered.setEnabled(False)    #������ ������ ����� ���������� ����� ����������
            
            self.spb.setValue(WorkOffset)   #�������� ��������
            self.spblered.setValue(WorkReload)   #�������� reload
            
            #���������� � ���� ��������� �������
            module.write_checkt('1')
        #���� checkbox ��������
        else:
            self.lenum.setEnabled(False)    #������ ������ ����� ��������������� ������ ����������
            self.leshut.setEnabled(False)        #������ ������ ����������
            self.spb.setEnabled(True)      #������ ������ ����� ������� ��������
            self.spblered.setEnabled(True)
            #���������� � ���� ��������� �������
            module.write_checkt('0')
        
    # �������� �� ������� �� ������ 'X'
    def closeEvent(self, event):
        # ���������� ��������� � ����� ��������: �Yes� � �No�.
        # ������ ������ ���������� � ������ ���������. ������ ������ � ��� ����� ���������, ������������ � ������� ����������� ����.
        # ������ �������� ��������� ���������� ������, ������������ � �������. ��������� �������� � ������ �� ���������.
        # ��� ������, ������� ������������� ����� �� ���� ��������� ����������.
        # ������������ �������� �������� � ���������� reply.
        
        #���� ���� � ������ �� ��������� � ��� ��� ������� � ����������� �����
        setting_dir_path = module.read_path() + '/' + module.read_name()
        setting_offset = module.read_offset()
        setting_reload = module.read_reload()
        setting_number = module.read_number()
        
        dir_path = self.le.text()
        work_offset = self.spb.value()
        work_reload = self.spblered.value()
        you_number = self.lenum.text()
        
        if (dir_path != setting_dir_path) or (work_offset != setting_offset) or (you_number != setting_number) or (work_reload != setting_reload):
            reply = QMessageBox.question(self, '���������', "�� ������ ��������� ���������?", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            # ���� ������ "��", ��������� ���� ������������ �������
            if reply == QMessageBox.Yes:
                self.save_setting_btn()
        #����������� ���������� � Tray
        event.ignore()                          #���������� ����� �� ���������
        self.hide()                             #�������� ���������
        self.tray_icon.showMessage(             #������� ���������
                "System Tray",
                "��������� ��������",
                QIcon('icon\Bill.jpg'),
                1
            )
        event.accept()                          #'''�� ������ ���������������!!!!'''
    
    #����� �� ���������
    @staticmethod
    def cleanUp(self):
    #def work_exit(self):
        #��������� � ��� ����
        module.log_info('����������!!!')
        #��������� � Exel ���� ����� ������
        work_time.quit_app()
        #������� ������ �� Tray
        self.tray_icon.hide()
        #��� �����
        qApp.quit()
        
#��������� ���� ����
#if __name__ == '__main__':
def app_main():
    app = QApplication(sys.argv)
    ex = MainWindow()
    app.aboutToQuit.connect(ex.cleanUp)
    ex.show()                   #�� ������ ���������������
    
    sys.exit(app.exec_())
