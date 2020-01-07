# -*- coding: cp1251 -*-
'''
������ � ��������� ������� �������
'''
from PyQt5.QtWidgets import QWidget, QStyle, QTextBrowser, QVBoxLayout, QLabel, QPushButton, qApp
from PyQt5.QtGui import QFont, QTextCursor
from PyQt5.QtCore import QThread, Qt

from work_lib import work_time, web_time, shutdown_lib
from work_setting import module, dialog
import datetime, webbrowser, time, sys

#���������� �������������� (��������� � ����� ��� ����� ���������� ����������)
class ShowShutOrWeb(QThread):
    def __init__(self):
        QThread.__init__(self)
        
    def run(self):
        #���� ���������� ������� ������ � ������ - ��������� ����
        if(module.read_check()):
            #��������� ������� ������ ������ � ����� ����� �� ������ ������
            web_time.web_main() 
            
        #����� ������� ����� "�����" ���������/���������� ����������
        else:
            #�������� ������� ���� � ����� �����
            tekdateandtimeStart = datetime.datetime.now()
    
            tekyear = tekdateandtimeStart.year   #������� ���
            tekmonth = tekdateandtimeStart.month #������� �����
            tekday = tekdateandtimeStart.day     #������� �����
            tekhour = tekdateandtimeStart.hour   #������� ���
            tekminute = tekdateandtimeStart.minute    #������� ������
            #���������� � Exel ���� ����� ���������� ���������� ����������
            work_time.write_exit()
    
            #���������� ����� ��������� ����������
            work_time.start_work(tekminute, tekhour, tekday, tekmonth, tekyear)
            
            #��������� ����������� ���� ��� ������ �������� �������
            #shutdown_lib.shutdown_lib()
            #2�� �������, ������ ���������� ������ ������ � ���� ������� ����� (��� ������ ������ �������� � �� ���� "������" ���������� ����������)
            while True:
                #�������� ������� �����
                timeExit = datetime.datetime.now()
                #���������� ������� ����� � ����
                module.write_timeExit(timeExit.strftime("%d %m %Y %H:%M"))
                time.sleep(60)

#������� ���� � ����������
class AdjWindow(QWidget):
   
    def __init__(self):
        # ����� super() ���������� ������ �������� ������ MainWindow � �� �������� ��� �����������.
        # ����� __init__() - ��� ����������� ������ � ����� Python.
        super(AdjWindow, self).__init__()
        #������� ������� ����
        #appearance = self.palette()
        #appearance.setColor(QPalette.Normal, QPalette.Window, QColor("white"))
                  
        self.resize(350,500)                                # ������������� ������������� ������� ����
        self.setWindowTitle("��� ������ ���� �������������� �����")  # ������������� ��������� ����
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_TitleBarContextHelpButton))   #������������� ���� �� ����������� ������
        #self.setPalette(appearance)                         #��������� ������� � ������ ����
        '''
        #������� ����� � �����������
        self.le_help = QTextEdit(self)
        self.le_help.setFont(QFont('Arial', 12))        #�����
        self.le_help.setText('����� ������� ����������')
        self.le_help.move(0, 0)                      #������������ � ����
        self.le_help.adjustSize()                           #���������� ������ � ����������� �� �����������
        '''
        #������� ����� � �����������
        
        text = module.read_help()#'������� �� ����: <br><a href="http://www.mars/asu/report/enterexit/">www.mars/asu/report/enterexit/</a><br> �����'
        #����� ���� ���� �������� ��� ����
        vbox = QVBoxLayout(self)
        #������� ���� � ������� ���������� � �������
        self.pole_vivod = QTextBrowser(self)
        self.pole_vivod.setFont(QFont('Arial', 14))        #�����
        self.pole_vivod.anchorClicked['QUrl'].connect(self.linkClicked)
        self.pole_vivod.setOpenLinks(False)     #������ �������� ������
        #self.pole_vivod.move(0, 0)
        vbox.addWidget(self.pole_vivod)
        self.setLayout(vbox)
        
        self.pole_vivod.append(text)
        self.pole_vivod.moveCursor(QTextCursor.Start)
        
    #������������ ���� �� ������
    def linkClicked(self, url):
        webbrowser.open(url.toString()) 
        
#����� ��� ������� ����������
class ShutWindow(QWidget):
   
    def __init__(self):
        # ����� super() ���������� ������ �������� ������ MainWindow � �� �������� ��� �����������.
        # ����� __init__() - ��� ����������� ������ � ����� Python.
        super(ShutWindow, self).__init__()
        #������� ������� ����
                  
        self.resize(200,200)                                # ������������� ������������� ������� ����
        self.setWindowFlags(Qt.FramelessWindowHint|Qt.WindowStaysOnTopHint)         # ���� ��� �����
        #self.setWindowOpacity(0.6)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        
        self.ldl = QLabel(self)
        self.ldl.setFont(QFont('Arial', 12))        #�����
        self.ldl.setText('�� ���������� ��������:')
        self.ldl.move(5, 10)                      #������������ � ����
        self.ldl.adjustSize()                           #���������� ������ � ����������� �� �����������
        
        self.timer = QLabel(self)
        self.timer.setFont(QFont('Arial', 100))        #�����
        self.timer.setText('60')
        self.timer.move(25, 20)                      #������������ � ����
        self.timer.adjustSize()                           #���������� ������ � ����������� �� �����������
        self.timer.setStyleSheet('color: red')                 #���� ������ �������
        
        self.btn_stop = QPushButton('����������\n����������', self)
        self.btn_stop.setFont(QFont('Arial', 12))        #�����
        self.btn_stop.move(50, 150)                      #������������ � ���� ������
        self.btn_stop.clicked.connect(self.closeprogramm)      #�������� �� �������
        
    def closeprogramm(self):
        self.hide()
        qApp.quit()