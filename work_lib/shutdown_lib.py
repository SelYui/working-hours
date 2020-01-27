# -*- coding: utf-8 -*-
'''
одуль для ловли выключения компьютера
'''

import os
import time, datetime
from work_setting import module
from work_lib import work_time

if os.name == "nt":
    import win32con
    import win32api
    import win32gui
else:
    import signal
    
class MyWindow:
    def __init__(self):
        win32gui.InitCommonControls()
        self.hinst = win32api.GetModuleHandle(None)
        className = 'MyWndClass'
        messageMap = {win32con.WM_QUERYENDSESSION : self.OnDestroy,
                      win32con.WM_ENDSESSION : self.OnDestroy,
                      win32con.WM_QUIT : self.OnDestroy,
                      win32con.WM_DESTROY : self.OnDestroy,
                      win32con.WM_CLOSE : self.OnDestroy }
        wc = win32gui.WNDCLASS()
        #wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.lpfnWndProc = messageMap
        wc.lpszClassName = className
        win32gui.RegisterClass(wc)
        style = win32con.WS_OVERLAPPEDWINDOW
        self.hwnd = win32gui.CreateWindow(
            className,
            'My win32api app',
            style,
            win32con.CW_USEDEFAULT,
            win32con.CW_USEDEFAULT,
            0,
            0,
            0,#win32con.HWND_MESSAGE,
            0,
            self.hinst,
            None)
        win32gui.UpdateWindow(self.hwnd)
        #win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
    
    #действия при ловле сигнала выключения компьютера
    #@staticmethod
    def OnDestroy(self, hwd, message, wparam, lparam):
        #work_time.quit_app()
        #получаем текущее время
        timeExit = datetime.datetime.now()
        #записываем время выключения в файл
        module.write_timeExit(timeExit.strftime("%d %m %Y %H:%M"))
        win32gui.PostQuitMessage(0)
        return True

#основная функция для использования  
def shutdown_lib():
    module.log_info("system: %s" % os.name)
    #определяем текущую ОС
    #если виндовс
    if os.name == "nt":
        #создаем окно windows и ловим сообщение выключения компьютера
        w= MyWindow()
        win32gui.PumpMessages()
    #если нет (Linux или Mac)
    else:
        #ловим сообщение выключения в другом формате
        while True:
            #signal.signal(signal.SIGTERM, MyWindow().OnDestroy)
            time.sleep(1)

#выключение компьютера
def signal_shutdown():
    if os.name == 'nt':
        os.system('shutdown -s')
    else:
        os.system('sudo shutdown now')
        #os.system('systemctrl poweroff')