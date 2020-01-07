# -*- coding: cp1251 -*-
'''
Модуль для подсчета рабочего времени и записи результатов в файл
'''
# -*- coding: cp1251 -*-
import datetime
import xlrd, xlwt
from xlutils.copy import copy

from work_setting import module

#Массив месяцев
month_word = ('за Январь','за Февраль','за Март','за Апрель','за Май','за Июнь','за Июль','за Август','за Сентябрь','за Октябрь','за Ноябрь','за Декабрь')
    
#функция для работы при включении компа
def start_work(tekminute, tekhour, tekday, tekmonth, tekyear):
    '''
    tekyear = tekdateandtime.year   #Текущий год
    tekmonth = tekdateandtime.month #текущий месяц
    tekday = tekdateandtime.day     #текущее число
    tekhour = tekdateandtime.hour   #текущий час
    tekminute = tekdateandtime.minute    #текущая минута
    '''
    #получаем путь к файлу и смещение
    wt_filename = module.read_path() + '/' + module.read_name()
    min_offset = module.read_offset()
    #обнуление начальных условий
    flg_dontdata = 0    #обнуляем признак незаполнения даты
    
    #вычитаем смещение из минут
    if tekminute - min_offset >= 0:
        tekminute = tekminute - min_offset
    else:
        tekhour = tekhour - 1
        tekminute = 60 + (tekminute - min_offset)
    
    #открываем наш Exel файл
    read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    write_book = copy(read_book)
    #Переходим на лист текущего года
    try:
        #если лист с текущим годом уже существует
        sheet_index = read_book.sheet_names().index(str(tekyear))
        #выбираем активным лист с нашим годом
        sheet = read_book.sheet_by_index(sheet_index)
        #чтобы не было ошибок
        sheet_nrows = sheet.nrows       #строка в Exel
        sheet_ncols = sheet.ncols       #столбец в Exel
            
    except ValueError:
        #если листа текущего года нет, создаем этот лист
        sheet = write_book.add_sheet(str(tekyear))
        sheet_index = read_book.nsheets
        #т.к. страница пустая
        sheet_nrows = 0
        sheet_ncols = 0

    #если лист не пустой, будем писать в конец
    if (sheet_nrows and sheet_ncols) != 0:
        #получаем последнюю дату
        i = sheet_nrows-1
        while i > 0:
            lastdate = sheet.row_values(i)[0]
            #если строка пустая, то ищем дату выше
            if lastdate == '':
                i=i-1
            else:
                dd = lastdate[0:2]  #день
                mm = lastdate[3:5]  #месяц
                break
        
        #если месяц совпал
        if tekmonth == int(mm):
            #проверяем на то что день на один меньше
            if tekday-1 == int(dd):
                #заполняем текущюю, строку
                i = sheet_nrows
            #за сегодня комп включился не первый раз
            elif tekday == int(dd):
                #сравниваем текущую дату прихода с первым (j) столбцом
                if time_compare(sheet_nrows, sheet, 1, tekhour, tekminute) == 0:    #если такая дата уже записанна, то ничего не делаем
                    return
                #проверяем время выключения, если менее 30 мин, то не записываем время прихода, выходим из программы
                if pc_reload(sheet.row_values(sheet_nrows-1)[2], tekhour, tekminute) == 0:
                    return
                    
                #не заполняем дату (признак незаполнения)
                flg_dontdata = 5826
                i = sheet_nrows
            #началась новая неделя
            else:
                i = sheet_nrows+1
        #начался новый месяц
        else:
            i = sheet_nrows+2
            #пишем месяц
            write_book.get_sheet(sheet_index).write(i,0,month_word[tekmonth-1])
            i=i+1
    #если лист пустой, начинаем заполнять с месяца
    else:
        i=0
        #пишем месяц
        write_book.get_sheet(sheet_index).write(i,0,month_word[tekmonth-1])
        i=i+1
    
    #заполняем строку датой
    if flg_dontdata != 5826:
        if tekday < 10 and tekmonth < 10:
            write_book.get_sheet(sheet_index).write(i,0,'0'+str(tekday)+'.0'+str(tekmonth)+'.'+str(tekyear))
        elif tekday < 10 and tekmonth > 10:
            write_book.get_sheet(sheet_index).write(i,0,'0'+str(tekday)+'.'+str(tekmonth)+'.'+str(tekyear))
        elif tekday > 10 and tekmonth < 10:
            write_book.get_sheet(sheet_index).write(i,0,''+str(tekday)+'.0'+str(tekmonth)+'.'+str(tekyear))
        else:
            write_book.get_sheet(sheet_index).write(i,0,str(tekday)+'.'+str(tekmonth)+'.'+str(tekyear))    
    #заполняем строку временем
    if tekhour < 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,1,'0'+str(tekhour)+':0'+str(tekminute))
    elif tekhour < 10 and tekminute > 10:
        write_book.get_sheet(sheet_index).write(i,1,'0'+str(tekhour)+':'+str(tekminute))
    elif tekhour > 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,1,str(tekhour)+':0'+str(tekminute))
    else:
        write_book.get_sheet(sheet_index).write(i,1,str(tekhour)+':'+str(tekminute))

    #если сейчас число больше 15, а предыдущее меньше либо равно, то вычисляем отработанное время в авансе
    if tekday > 15 and int(dd) <= 15:
        #вычисляем сумму в этом месяце
        i = sheet_nrows-1
        #ищем наименование месяца в файле
        while sheet.row_values(i)[0] != month_word[tekmonth-1]:
            i=i-1
        #суммируем все ячейки часов в месяце
        #начальное значение отработанных часов
        mount_sum = 0
        while i <= sheet_nrows-1:
            #если пустая строка
            if sheet.row_values(i)[3] == '':
                i=i+1
            #получаем часы отработанные в дне
            else:
                mount_sum = mount_sum + float(sheet.row_values(i)[3])
                i=i+1
        #округляем до 3его знака
        mount_sum = round(mount_sum,3)
        #заполняем сумму часов в соответствующую строку
        write_book.get_sheet(sheet_index).write(i,4,'('+str(mount_sum)+')')

    #сохраняем запись
    try:
        write_book.save(wt_filename)
    except Exception as e:
        module.log_info("Не удалось сохранить в Exel время прихода. Exception: %s" % str(e))

#функция для работы при выключении компа     
def exit_work(tekminute, tekhour, tekday, tekmonth, tekyear):
    '''
    tekyear = tekdateandtime.year   #Текущий год
    tekmonth = tekdateandtime.month #текущий месяц
    tekday = tekdateandtime.day     #текущее число
    tekhour = tekdateandtime.hour   #текущий час
    tekminute = tekdateandtime.minute    #текущая минута
    '''

    #получаем путь к файлу и смещение
    wt_filename = module.read_path() + '/' + module.read_name()
    min_offset = module.read_offset()
    
    #вычитаем смещение из минут
    if tekminute + min_offset < 60:
        tekminute = tekminute + min_offset
    else:
        tekhour = tekhour + 1
        tekminute = (tekminute + min_offset) - 60
    
    #открываем наш Exel файл
    read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    write_book = copy(read_book)
    
    #Переходим на лист текущего года
    try:
        #если лист с текущим годом уже существует
        sheet_index = read_book.sheet_names().index(str(tekyear))
    except:
        #если листа текущего года нет значит программа start не сработала
        return
    
    #выбираем активным лист с нашим годом
    sheet = read_book.sheet_by_index(sheet_index)
    
    #получаем последнюю дату
    i = sheet.nrows-1
    while i > 0:
        lastdate = sheet.row_values(i)[0]
        #если строка пустая, то ищем дату выше
        if lastdate == '':
            i=i-1
        else:
            dd = lastdate[0:2]  #день
            mm = lastdate[3:5]  #месяц
            #в этой строке будем писать сумарное количество часов в дне
            time_date_index = i
            break
    
    #если месяц совпал
    if tekmonth == int(mm):
        #проверяем на то что день такой же
        if tekday == int(dd):
            #заполняем текущюю, строку 
            i = sheet.nrows-1
        #Если что-то не совпало, то start не сработал чтобы не заполнить другую строку
        else: return
    else: return
    
    #сравниваем текущую дату прихода со вторым (j) столбцом
    if time_compare(sheet.nrows, sheet, 2, tekhour, tekminute) == 0:    #если такая дата уже записанна, то ничего не делаем
        return

    #заполняем строку временем
    if tekhour < 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,2,'0'+str(tekhour)+':0'+str(tekminute))
    elif tekhour < 10 and tekminute > 10:
        write_book.get_sheet(sheet_index).write(i,2,'0'+str(tekhour)+':'+str(tekminute))
    elif tekhour > 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,2,str(tekhour)+':0'+str(tekminute))
    else:
        write_book.get_sheet(sheet_index).write(i,2,str(tekhour)+':'+str(tekminute))
    
    #вычисляем количество рабочих часов в сегодняшнем дне
    i = sheet.nrows-1
    sum_timework = 0
    count_cycle = 0
    #пока не дошли до первой строки дня
    while i >= time_date_index:
        #получаем время прихода
        time_start = sheet.row_values(i)[1]
        #выделяем часы и минуты прихода
        hr_start = time_start[0:2]  #часы
        min_start = time_start[3:5]  #минуты
        timestart = int(hr_start) + int(min_start)/60
        #выделяем часы и минуты ухода
        if (int(sheet.ncols) < 3) or (sheet.row_values(i)[2] == '') or (i == time_date_index):
            hr_exit = tekhour  #часы
            min_exit = tekminute  #минуты
        else:
            time_exit = sheet.row_values(i)[2]
            hr_exit = time_exit[0:2]  #часы
            min_exit = time_exit[3:5]  #минуты
        timeexit = int(hr_exit) + int(min_exit)/60
        #разность
        timework = timeexit - timestart
        #вычитаем обед. Если одна строка и рабочих часов больше 4, вычитаем час
        if (i == time_date_index) and timework > 4 and count_cycle == 0:
            timework = timework -1
        #сумма рабочих часов в дне
        sum_timework = sum_timework + timework
        count_cycle = count_cycle+1
        i = i-1

    #округляем до 3его знака
    sum_timework = round(sum_timework,3)
    #заполняем строку часов
    write_book.get_sheet(sheet_index).write(time_date_index,3,str(sum_timework))
    #вычисляем количество рабочих часов в текущем месяце
    i = sheet.nrows-1
    #ищем наименование месяца в файле
    while sheet.row_values(i)[0] != month_word[tekmonth-1]:
        i=i-1
    #запоминаем строку, куда запишем сумму
    index_sum = i
    #суммируем все ячейки часов в месяце
    i = time_date_index-1
    #начальное значение отработанных часов - только что вычисленное значение
    mount_sum = sum_timework
    while i > int(index_sum):
        #если пустая строка
        if sheet.row_values(i)[3] == '':
            i=i-1
        else:
            #получаем часы отработанные в дне
            mount_sum = mount_sum + float(sheet.row_values(i)[3])
            i=i-1
    #округляем до 3его знака
    mount_sum = round(mount_sum,3)
    #заполняем сумму часов в соответствующую строку
    write_book.get_sheet(sheet_index).write(index_sum,1,str(mount_sum))
    
    #сохраняем запись
    try:
        write_book.save(wt_filename)
    except Exception as e:
        module.log_info("Не удалось сохранить в Exel время ухода. Exception: %s" % str(e))

#если комп выключался не на долгото не заполняем время прихода (возвращаем заполнять или не заполнять)
def pc_reload(timeexit, starthour, startminute):
    reload = module.read_reload()
    #если строка пустая, выходим из программы, такого не должно быть
    if timeexit == '':
        module.log_info("pc reload = 2")
        return 2    #неизвестная ошибка - продолжаем работу
    else:
        hour_e = timeexit[0:2]      #часы
        minut_e = timeexit[3:5]     #минуты
    #считаем время ухода и прихода
    te = int(hour_e) + int(minut_e)/60
    ts = int(starthour) + int(startminute)/60
    #если разность текущего прихода и ухода менее получаса, то не добавляйте новую строчку прихода
    if ts - te < (reload/60):
        return 0    #уходили не на долго - выходим
    else: return 1  #уходили на долго - продолжаем работу

#сравниваем время текущее с временем в Exel, если совпало, то не будем записывать
def time_compare(sheet_nrows, sheet, j, tekhour, tekminute):
    #получаем последнее время прихода
    i = sheet_nrows-1
    while i > 0:
        lastdate = sheet.row_values(i)[j]
        #если строка пустая, то ищем дату выше
        if lastdate == '':
            i=i-1
        else:
            shour = lastdate[0:2]  #день
            sminute = lastdate[3:5]  #месяц
            break
    #если время совпало с последним записанным временем - ничего не делаем. Выходим из программы
    if (int(shour) == tekhour) and (int(sminute) == tekminute):
        return 0    #текущая дата совпадает с последней записанной
    else: return 1  #не совпадает
        
#действия при выходе из программы по кнопке
def quit_app():
    #получаем текущую дату и время компа
    tekdateandtimeExit = datetime.datetime.now()
    '''
    tekyear = tekdateandtimeExit.year   #Текущий год
    tekmonth = tekdateandtimeExit.month #текущий месяц
    tekday = tekdateandtimeExit.day     #текущее число
    tekhour = tekdateandtimeExit.hour   #текущий час
    tekminute = tekdateandtimeExit.minute    #текущая минута
    '''
    #дату записываем время выключения компьютера в файл
    module.write_timeExit(tekdateandtimeExit.strftime("%d %m %Y %H:%M"))
    
    #записываем время выключения компьютера
    #exit_work(tekminute, tekhour, tekday, tekmonth, tekyear)

def write_exit():
    #записываем в Exel файл время последнего выключения компьютера
    dtimE = module.read_timeExit()
    dtimE = dtimE.split()
    #если приложение закрыл не пользователь
    if (dtimE != ''):
        timE = dtimE[-1]
        exit_work(int(timE[3:5]), int(timE[0:2]), int(dtimE[0]), int(dtimE[1]), int(dtimE[2]))
