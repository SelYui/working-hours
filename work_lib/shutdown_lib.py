# -*- coding: utf-8 -*-
'''
Модуль для выключения компьютера
'''

import os

#выключение компьютера
def signal_shutdown():
    if os.name == 'nt':
        os.system('shutdown -s')
    else:
        os.system('sudo shutdown now')
        #os.system('systemctrl poweroff')