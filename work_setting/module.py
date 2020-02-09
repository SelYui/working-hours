# -*- coding: utf-8 -*-
'''
Модули для записи и чтения из настроечного файла work_setting.txt
'''
import os, datetime

#запись в лог файл сбоев + даты сбоя
def log_info(msg):
    ''' пример вызова
    module.log_info("date: %s" % date)
    '''
    #получаем время сбоя
    log_time = datetime.datetime.now()
    #записываем сбой в файл
    f = open("work_setting/working_hour.log", "a", encoding = 'utf-8')
    f.write(msg + "            " + str(log_time) + "\n")
    f.close()

#универсальная запись в настроечный файл
def write_setting(date, num_setting):
    #читаем файл построчно
    try:
        f = open('work_setting/work_setting.txt', 'r', encoding = 'utf-8')
        lines = f.readlines()
    except:
        log_info("write_setting: не удалось открыть файл")
    finally:
        f.close()
    
    #заменяем заданную строку на новую
    #проверяем что в файле есть необходимая строка
    if len(lines)-1 >= num_setting:
        lines[num_setting] = str(date) + '\n'
    #сохраняем весь список строк в файл
    try:
        save_f = open('work_setting/work_setting.txt', 'w', encoding = 'utf-8')
        save_f.writelines(lines)
    except:
        log_info("write_setting: не удалось записать в файл %s %s"%(date, num_setting))
    finally:
        save_f.close()
#считываем из файла заданной строки
def read_setting(num_setting):
    #обнуляем переменную
    date = '' #не выставлен
    #делаем 10 попыток чтения
    for cnt_read in range(10):
        #читаем файл построчно и возвращаем 22 строку
        f = open('work_setting/work_setting.txt', 'r', encoding = 'utf-8')
        with f:
            lines = f.readlines()
            try:
                date = lines[num_setting]
                break
            except:
                log_info("read_setting: не считалась строка: %s %d раз" % (num_setting, cnt_read+1))
    else:
        log_info("read_setting: не удалось считать строку: %s" % num_setting)
    #без символа переноса строки, возвращаем тип srt
    return date[:-1]

#читаем из файла инструкцию
def read_help():
    #читаем файл построчно
    f = open('work_setting/work_help.txt', 'r', encoding = 'utf-8')
    text = f.read()
    f.close()
    
    return str(text)

#создание нового exel файла
def new_timework_file(path):
    f = open(path, 'w')
    f.close()

#сохранение в файл (не используется)
def save_setting(new_path, mode):
    
    old_WorkPath = read_setting(4)              #читаем старое значение пути
    old_WorkName = read_setting(1)              #читаем старое значение имени файла
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
            write_setting(WorkName,1)
            #запись в настроечный файл нового пути
            write_setting(WorkPath,4)
    elif (mode == 'New'):
        save_warning = 1
        #создаем файл
        new_file = os.open((WorkPath + '/' + WorkName),os.O_CREAT)
        os.close(new_file)
    #возвращаем в вызываемый модуль читанные из настроечного файла путь и имя файла
    #и ошибку, если таковая имеется: 0-все хорошо, 1-недопустимая директория, 2-имя файла не существует)
    return read_setting(4), read_setting(1), save_warning
        
        