# -*- coding: cp1251 -*-
'''
������ � ��������� ������� �������
'''
from PyQt5.QtWidgets import QWidget, QLabel, QPushButton, qApp, QApplication
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QThread, Qt, QBasicTimer, QObject, pyqtSignal, pyqtSlot

from work_lib import work_time, web_time, shutdown_lib
from work_setting import module, dialog
import datetime, time, sys, threading
from win32com.test.testIterators import SomeObject

#���������� �������������� (��������� � ����� ��� ����� ���������� ���������� � ��������� ������)
class ShowShutOrWeb(QThread):
    def __init__(self):
        QThread.__init__(self)
        
    def run(self):
        #���� ���������� ������� ������ � ������ - ��������� ����
        print(1)
        if(module.read_check()):
            #��������� ������� ������ ������ � ����� ����� �� ������ ������
            print(2)
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
        
#����� ��� ������� ����������
class ShutWindow(QWidget):
   
    def __init__(self):
        # ����� super() ���������� ������ �������� ������ MainWindow � �� �������� ��� �����������.
        # ����� __init__() - ��� ����������� ������ � ����� Python.
        super(ShutWindow, self).__init__()
                  
        self.ldl = QLabel(self)                 #����� �����������
        self.ldl.setFont(QFont('Arial', 12))        #�����
        self.ldl.setText('�� ���������� ��������:')
        self.ldl.move(5, 10)                      #������������ � ����
        self.ldl.adjustSize()                           #���������� ������ � ����������� �� �����������
        
        self.lbl_timer = QLabel(self)                   #����� �� ���������
        self.lbl_timer.setFont(QFont('Arial', 100))        #�����
        self.lbl_timer.setText('60')
        self.lbl_timer.move(25, 20)                      #������������ � ����
        self.lbl_timer.adjustSize()                           #���������� ������ � ����������� �� �����������
        self.lbl_timer.setStyleSheet('color: red')                 #���� ������ �������
        
        self.btn_stop = QPushButton('����������\n����������', self) #��������� ��������
        self.btn_stop.setFont(QFont('Arial', 12))        #�����
        self.btn_stop.move(50, 150)                      #������������ � ���� ������
        self.btn_stop.clicked.connect(self.closeprogramm)      #�������� �� �������

        #������� ����� ������ �����
        self.obj = ShutTimer()
        self.thread = QThread()
        #��������� ������� �� ������� ����� ��� ������ ������
        self.obj.intReady.connect(self.onIntReady)
        #���������� worker � thread
        self.obj.moveToThread(self.thread)
        #���������� ������� worker � ������ ������
        self.obj.finished.connect(self.thread.quit)
        #������ ���������� ����������� � ������ worker
        self.thread.started.connect(self.obj.CountTime)
        #������ ���������� ������ ������� ����������
        #self.thread.finished.connect(app.exit)
        
        #������ ������
        self.thread.start()
        #������ �����
        self.initUI()
        
    def initUI(self):
        
        self.resize(200,200)                                # ������������� ������������� ������� ����
        self.setWindowFlags(Qt.FramelessWindowHint|Qt.WindowStaysOnTopHint)         # ���� ��� �����
        #self.setWindowOpacity(0.6)
        self.setAttribute(Qt.WA_TranslucentBackground)
    
    def onIntReady(self, i):
        self.lbl_timer.setText("{}".format(i))
        print(i)
    
    #�� ������� ������ 
    def closeprogramm(self):
        qApp.quit()

#���������� �������������� ��� �������
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
            
#�������� ���� � ��������
def app_ShutWindow():
    ''' �� ����������...
    #� ����� ��������� ������� �������� ���������� � ������
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
