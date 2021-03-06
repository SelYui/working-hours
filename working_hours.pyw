# -*- coding: utf-8 -*-
'''
Основной модуль выполнения программ
'''

from work_setting import dialog, module
from work_lib import work_time

if __name__ == '__main__':
    #если не записано вчерашнее время ухода по какой-то причине - записываем его из настроечного файла
    err_exit = module.read_setting(28)
    if int(err_exit) != 0:
        work_time.write_exit()   
    #выставляем признак нештатного завершения (что бы в случае чего записать время завтра)
    module.write_setting(1, 28)                    #не забыть раскомментировать!!!
    
    #открываем окно приложения из dialog
    dialog.app_main()
    