# -*- coding: cp1251 -*-
'''
основной модуль выполнения программ
'''

from work_setting import dialog, adjacent_classes

if __name__ == '__main__':
    
    #запускаем в отдельном потоке класс ShowWeb
    thread = adjacent_classes.ShowShutOrWeb()
    thread.start()
    
    #открываем окно приложения из dialog
    dialog.app_main()
