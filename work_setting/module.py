# -*- coding: cp1251 -*-
'''
Модули для записи и чтения из настроечного файла work_setting.txt
'''
import os

#запись в лог файл сбоев
def log_info(msg):
    ''' пример вызова
    module.log_info("hwnd: %s" % hwnd)
    '''
    f = open("work_setting\working_hour.log", "a")
    f.write(msg + "\n")
    f.close()
    
#получаем имя файла Exel из нашего настроечного файла
def read_name():
    #задаем стандартное имя фйла
    WorkFile = 'Рабочее_время.xls'
    
    #читаем файл построчно и возвращаем первую строку
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines_name = f.readlines()
        try:
            WorkFile = lines_name[1]
        except:
            None
    #без символа переноса строки
    return WorkFile[:-1]

#запись в настроечный файл нового имени файла
def write_name(new_name):
    #читаем файл построчно
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #заменяем первую строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= 1:
        lines[1] = new_name + '\n'
    #добавляем в конец списка новый путь
    else:
        lines.append(new_name)
    #сохраняем весь список строк в файл
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#получаем путь к файлу Exel из нашего настроечного файла
def read_path():
    #узнаем текущий каталог для работы
    WorkPath = os.path.dirname(os.path.realpath(__file__)) + '\Рабочее_время.xls'
    
    #читаем файл построчно и возвращаем первую строку
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            WorkPath = lines[4]
        except:
            None
    #без символа переноса строки
    return WorkPath[:-1]

#запись в настроечный файл нового пути
def write_path(new_path):
    #читаем файл построчно
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #заменяем четверткю строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= 4:
        lines[4] = new_path + '\n'
    #добавляем в конец списка новый путь
    else:
        lines.append(new_path)
    #сохраняем весь список строк в файл
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#считываем из файла смещение
def read_offset():
    #узнаем текущий каталог для работы
    WorkOffset = 0
    
    #читаем файл построчно и возвращаем первую строку
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            WorkOffset = lines[7]
        except:
            None
    #без символа переноса строки
    return int(WorkOffset)
    
#записываем смещение в файл
def write_offset(new_offset):
    #читаем файл построчно
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #заменяем седьмую строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= 7:
        lines[7] = str(new_offset) + '\n'
    #добавляем в конец списка новый путь
    else:
        lines.append(str(new_offset))
    #сохраняем весь список строк в файл
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#считываем из файла допустимое время ухода
def read_reload():
    #обнуляем
    WorkReload = 0
    
    #читаем файл построчно и возвращаем 10 строку
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            WorkReload = lines[10]
        except:
            None
    
    return int(WorkReload)
    
#записываем смещение в файл
def write_reload(new_reload):
    #читаем файл построчно
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #заменяем 10 строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= 10:
        lines[10] = str(new_reload) + '\n'
    #добавляем в конец списка новый путь
    else:
        lines.append(str(new_reload))
    #сохраняем весь список строк в файл
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()


#считываем из файла выставлен или нет флаг чтения с сайта
def read_check():
    #узнаем текущий каталог для работы
    CheckNum = 0 #не выставлен
    
    #читаем файл построчно и возвращаем 13 строку
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            CheckNum = lines[13]
        except:
            None
    #без символа переноса строки
    return int(CheckNum)
    
#записываем флаг чтения с сайта в файл
def write_checkt(new_check):
    #читаем файл построчно
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #заменяем седьмую строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= 13:
        lines[13] = str(new_check) + '\n'
    #добавляем в конец списка новый путь
    else:
        lines.append(str(new_check))
    #сохраняем весь список строк в файл
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#считываем из файла номр сотрудника
def read_number():
    #узнаем текущий каталог для работы
    CheckNum = 0 #не выставлен
    
    #читаем файл построчно и возвращаем 16 строку
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            CheckNum = lines[16]
        except:
            None
    #без символа переноса строки
    return CheckNum[:-1]
    
#записываем в файл номер сотрудника
def write_number(new_number):
    #читаем файл построчно
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #заменяем 16 строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= 16:
        lines[16] = str(new_number) + '\n'
    #добавляем в конец списка новый путь
    else:
        lines.append(str(new_number))
    #сохраняем весь список строк в файл
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#запись в конец настроечного файла
def write_timeShut(timeShut):
    #читаем файл построчно
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #заменяем 19 строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= 19:
        lines[19] = str(timeShut) + '\n'
    #добавляем в конец списка новый путь
    else:
        lines.append(str(timeShut))
    #сохраняем весь список строк в файл
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()
    
#считываем из файла последнюю строку (строку ухода с марса)
def read_timeShut():
    #обнуляем переменную
    timeShut = 0 #не выставлен
    
    #читаем файл построчно и возвращаем 19 строку
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            timeShut = lines[19]
        except:
            None
    #без символа переноса строки
    return timeShut[:-1]

#запись в конец настроечного файла
def write_timeExit(timeExit):
    #читаем файл построчно
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #заменяем 22 строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= 22:
        lines[22] = str(timeExit) + '\n'
    #добавляем в конец списка новый путь
    else:
        lines.append(str(timeExit))
    #сохраняем весь список строк в файл
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()
    
#считываем из файла последнюю строку (строку ухода с марса)
def read_timeExit():
    #обнуляем переменную
    timeExit = 0 #не выставлен
    
    #читаем файл построчно и возвращаем 22 строку
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            timeExit = lines[22]
        except:
            None
    #без символа переноса строки
    return timeExit[:-1]

#читаем из файла инструкцию
def read_help():
    #читаем файл построчно
    f = open('work_setting\work_help.txt', 'r')
    text = f.read()
    f.close()
    
    return str(text)

#создание нового exel файла
def new_timework_file(path):
    f = open(path, 'w')
    f.close()

#сохранение в файл
def save_setting(new_path, mode):
    
    old_WorkPath = read_path()              #читаем старое значение пути
    old_WorkName = read_name()              #читаем старое значение имени файла
    WorkPath = os.path.dirname(new_path)    #путь папки в которой лежит файл
    WorkName = os.path.basename(new_path)   #имя файла
    
    #перезапись
    if(mode == 'Repace'):
        #если выбранный каталог существует перемещаем туда файл с новым именем
        if os.path.exists(WorkPath):        
            #копировать файл, даже если такое имя уже существует
            os.replace((old_WorkPath + '/' + old_WorkName),(WorkPath + '/' + WorkName))
            save_warning = 2
            #запись в настроечный файл нового имени файда
            write_name(WorkName)
            #запись в настроечный файл нового пути
            write_path(WorkPath)
    elif (mode == 'New'):
        save_warning = 1
        #создаем файл
        new_file = os.open((WorkPath + '/' + WorkName),os.O_CREAT)
        os.close(new_file)
    #возвращаем в вызываемый модуль читанные из настроечного файла путь и имя файла
    #и ошибку, если таковая имеется: 0-все хорошо, 1-недопустимая директория, 2-имя файла не существует)
    return read_path(), read_name(), save_warning
        
        