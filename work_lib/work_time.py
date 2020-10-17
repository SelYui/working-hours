# -*- coding: utf-8 -*-
'''
Модуль для подсчета рабочего времени и записи результатов в файл
'''
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
    wt_filename = module.read_setting(4) + '/' + module.read_setting(1)
    
    #обнуление начальных условий
    flg_dontdata = False    #обнуляем признак незаполнения даты
    flg_newmonth = False    #флаг начала нового месяца

    dd = mm = 8520  #пишем бред, что бы в случае чего программа не вылетела
    #открываем наш Exel файл
    try:
        read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    except Exception as e:
        module.log_info('start_work: файл Exel отсутствует %s'% str(e))
        return
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
        #ищем текущий месяц
        while i > 0:
            #lastdate = sheet.row_values(i)[0]
            lastdate = read_excel(read_book, sheet, i, 0)
            #если строка пустая, то ищем дату выше
            if lastdate[0] == '':
                i=i-1
            else:
                dd = lastdate[0]  #день
                mm = lastdate[1]  #месяц
                time_date_index = i
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
                if time_compare(read_book, sheet_nrows, sheet, 1, tekhour, tekminute, time_date_index) == 0:    #если такая дата уже записанна, то ничего не делаем
                    return
                #проверяем время выключения, если менее 30 мин, то не записываем время прихода, выходим из программы
                #if pc_reload(sheet.row_values(sheet_nrows-1)[2], tekhour, tekminute) == 0:
                if pc_reload(read_excel(read_book, sheet, sheet_nrows-1, 2), tekhour, tekminute) == 0:
                    return
                    
                #не заполняем дату (признак незаполнения)
                flg_dontdata = True
                i = sheet_nrows
            #началась новая неделя
            else:
                i = sheet_nrows+1
        #начался новый месяц
        else:
            flg_newmonth = True
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
    if flg_dontdata == False:
        if tekday < 10 and tekmonth < 10:
            write_book.get_sheet(sheet_index).write(i,0,'0'+str(tekday)+'.0'+str(tekmonth)+'.'+str(tekyear))
        elif tekday < 10 and tekmonth >= 10:
            write_book.get_sheet(sheet_index).write(i,0,'0'+str(tekday)+'.'+str(tekmonth)+'.'+str(tekyear))
        elif tekday >= 10 and tekmonth < 10:
            write_book.get_sheet(sheet_index).write(i,0,''+str(tekday)+'.0'+str(tekmonth)+'.'+str(tekyear))
        else:
            write_book.get_sheet(sheet_index).write(i,0,str(tekday)+'.'+str(tekmonth)+'.'+str(tekyear))
    #заполняем строку временем
    if tekhour < 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,1,'0'+str(tekhour)+':0'+str(tekminute))
    elif tekhour < 10 and tekminute >= 10:
        write_book.get_sheet(sheet_index).write(i,1,'0'+str(tekhour)+':'+str(tekminute))
    elif tekhour >= 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,1,str(tekhour)+':0'+str(tekminute))
    else:
        write_book.get_sheet(sheet_index).write(i,1,str(tekhour)+':'+str(tekminute))
    #вычисляем отработанное время в авансе
    if (tekday > 15 and int(dd) <= 15) and flg_newmonth == False:
        #вычисляем сумму в этом месяце
        i = sheet_nrows-1
        #ищем наименование месяца в файле
        #while sheet.row_values(i)[0] != month_word[tekmonth-1]:
        temp = read_excel(read_book, sheet, i, 0)
        while temp[6] != month_word[tekmonth-1]:
            i=i-1
            temp = read_excel(read_book, sheet, i, 0)
        #суммируем все ячейки часов в месяце
        #начальное значение отработанных часов
        mount_sum = 0
        while i <= sheet_nrows-1:
            #если пустая строка
            #if sheet.row_values(i)[3] == '':
            temp = read_excel(read_book, sheet, i, 3)
            if temp[0] == '':
                i=i+1
            #получаем часы отработанные в дне
            else:
                #mount_sum = mount_sum + float(sheet.row_values(i)[3])
                mount_sum = mount_sum + float(temp[6])
                i=i+1
        #округляем до 3его знака
        mount_sum = round(mount_sum,3)
        #заполняем сумму часов в соответствующую строку
        write_book.get_sheet(sheet_index).write(time_date_index,4,'('+str(mount_sum)+')')
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
    #получаем путь к файлу
    wt_filename = module.read_setting(4) + '/' + module.read_setting(1)
    
    #открываем наш Exel файл
    try:
        read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    except:
        module.log_info('exit_work: файл Exel отсутствует')
        return
    write_book = copy(read_book)
    #Переходим на лист текущего года
    try:
        #если лист с текущим годом уже существует
        sheet_index = read_book.sheet_names().index(str(tekyear))
    except:
        #если листа текущего года нет значит программа start не сработала
        module.log_info('exit_work: в Exel файле отсутствует страница %s' %tekyear)
        return
    
    #выбираем активным лист с нашим годом
    sheet = read_book.sheet_by_index(sheet_index)
    #получаем последнюю дату
    i = sheet.nrows-1
    if i<0:
        #если лист пустой не записываем время ухода и выходим из программы
        module.log_info('exit_work: в Exel файле страница %s пуста' %tekyear)
        return
    #начальные значения
    dd = mm = 8520
    
    while i > 0:
        #lastdate = sheet.row_values(i)[0]
        lastdate = read_excel(read_book, sheet, i, 0)
        #если строка пустая, то ищем дату выше
        if lastdate[0] == '':
            i=i-1
        else:
            dd = lastdate[0]  #день
            mm = lastdate[1]  #месяц
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
        else:
            module.log_info('exit_work: нет текущей даты %s'% sheet.nrows)
            return
    else:
        module.log_info('exit_work: нет текущего месяца %s'% sheet.nrows)
        return
    #сравниваем текущую дату прихода со вторым (j) столбцом
    if time_compare(read_book, sheet.nrows, sheet, 2, tekhour, tekminute, time_date_index) == 0:    #если такая дата уже записанна, то ничего не делаем
        return
    #заполняем строку временем
    if tekhour < 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,2,'0'+str(tekhour)+':0'+str(tekminute))
    elif tekhour < 10 and tekminute >= 10:
        write_book.get_sheet(sheet_index).write(i,2,'0'+str(tekhour)+':'+str(tekminute))
    elif tekhour >= 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,2,str(tekhour)+':0'+str(tekminute))
    else:
        write_book.get_sheet(sheet_index).write(i,2,str(tekhour)+':'+str(tekminute))
    #сохраняем запись
    try:
        write_book.save(wt_filename)
    except Exception as e:
        module.log_info("Не удалось сохранить в Exel. Exception: %s" % str(e))    
    #получаем массив дней в месяце
    day_in_mount = arr_day_month(tekmonth, tekyear)
    #получаем номер середины месяца для подсчета аванса
    num_center_day = center_month(day_in_mount)
    #получаем количество часов в месяце и до аванса, и массив часов работы в дне
    sum_month = month_recount(day_in_mount, num_center_day)
    #записываем в Exel
    if write_sum_month(sum_month[0], sum_month[1], num_center_day, sum_month[2], tekmonth, tekyear) == False:
        module.log_info("Не удалось сохранить в Exel время прихода.")
    
#функция для чтения даты и времени из Excel файла
def read_excel(read_book, sheet, rowx, colx):
    try:
        #считываем значение ячейки
        type_cell = sheet.cell_type(rowx, colx)
        data = sheet.cell_value(rowx, colx)
    except:
        type_cell = 1234
        data = 0
    #если в ячейке формат xldata:
    if type_cell == 3:
        y, m, d, h, i, s = xlrd.xldate_as_tuple(data, read_book.datemode)
        try:
            date = str_date(d, m, y)
        except:
            date = data
        try:
            time = str_time(h, i)
        except:
            time = data
    #если в ячейке тестовый формат
    elif type_cell == 1:
        d = data[0:2]  #день
        m = data[3:5]  #месяц
        y = data[6:]    #год
        h = data[0:2]   #часы
        i = data[3:5]   #минуты
        s = 0        #секунды
        try:
            date = str_date(d, m, y)
        except:
            date = data
        try:
            time = str_time(h, i)
        except:
            time = data
    #что-то другое
    else:
        y, m, d, h, i, s = data, data, data, data, data, data
        try:
            date = str_date(d, m, y)
        except:
            date = data
        try:
            time = str_time(h, i)
        except:
            time = data
    return d, m, y, h, i, s, date, time

#если комп выключался не на долгото не заполняем время прихода (возвращаем заполнять или не заполнять)
def pc_reload(timeexit, starthour, startminute):
    reload = int(module.read_setting(13))
    #если строка пустая, выходим из программы, такого не должно быть
    if (timeexit[3] == '') or (timeexit[4] == ''):
        module.log_info("pc reload = -2")
        return -2    #неизвестная ошибка - продолжаем работу
    else:
        hour_e = timeexit[3]      #часы
        minut_e = timeexit[4]     #минуты
    #считаем время ухода и прихода
    te = int(hour_e) + int(minut_e)/60
    ts = int(starthour) + int(startminute)/60
    #если разность текущего прихода и ухода менее получаса, то не добавляйте новую строчку прихода
    if ts - te <= (reload/60):
        return 0    #уходили не на долго - выходим
    else:
        return 1  #уходили на долго - продолжаем работу

#сравниваем время текущее с временем в Exel, если совпало или меньше, то не будем записывать
def time_compare(read_book, sheet_nrows, sheet, j, tekhour, tekminute, time_date_index):
    shour = sminute = 0
    #получаем последнее время прихода
    i = sheet_nrows-1
    #пока мы в текущем дне
    while (i >= time_date_index):#i > 0:
        #lastdate = sheet.row_values(i)[j]
        lastdate = read_excel(read_book, sheet, i, j)
        #если строка пустая, то ищем дату выше
        if lastdate[0] == '':
            i=i-1
        else:
            shour = lastdate[0]  #день
            sminute = lastdate[1]  #месяц
            break
    #если время совпало с последним записанным временем - ничего не делаем. Выходим из программы
    if ((int(shour) == tekhour) and (int(sminute) >= tekminute)) or ((int(shour) > tekhour)):
        return 0    #текущая дата совпадает с последней записанной
    else:
        return 1  #не совпадает
        
#действия при выходе из программы по кнопке
def quit_app():
    #получаем текущую дату и время компа
    tekdateandtimeExit = datetime.datetime.now()
    
    #записываем время выключения компьютера в настроечный файл
    module.write_setting(tekdateandtimeExit.strftime("%d %m %Y %H:%M"), 25)

#записываем в Exel файл время последнего выключения компьютера
def write_exit():
        
    dtimE = module.read_setting(25)
    
    #если приложение закрыл не пользователь
    if (dtimE != ''):
        dtimE = dtimE.split()
        
        timE = dtimE[-1]
        
        #получаем смещение
        min_offset = int(module.read_setting(10))
        minute = int(timE[3:5])
        hour = int(timE[0:2])
        #вычитаем смещение из минут
        if minute + min_offset < 60:
            minute = minute + min_offset
        else:
            hour = hour + 1
            minute = (minute + min_offset) - 60
        
        exit_work(minute, hour, int(dtimE[0]), int(dtimE[1]), int(dtimE[2]))
    else:
        msg = "dtimE = %s"% dtimE
        module.log_info(msg)
###############################################################################################
#получаем лист из Exel файла
def year_sheet(year):
    #получаем путь к файлу и смещение
    wt_filename = module.read_setting(4) + '/' + module.read_setting(1)
    
    #открываем наш Exel файл
    read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    write_book = copy(read_book)
    
    #Переходим на лист текущего года
    try:
        #если лист с текущим годом уже существует
        sheet_index = read_book.sheet_names().index(str(year))
    except:
        #если листа текущего года нет значит программа start не сработала
        module.log_info('year_sheet: в Exel файле отсутствует страница %s' %year)
        return 0
    
    #выбираем активным лист с нашим годом
    return read_book.sheet_by_index(sheet_index), write_book, read_book

#функция получения массива времян в заданном дне
def arr_time_day(day=0, month=0, year=0):
    date = datetime.datetime.now()
    tekmonth = date.month #текущий месяц
    tekyear = date.year
    tekday = date.day
    #если параметры не заданы, считаем текущий день
    if (day!=0) and (month!=0) and (year!=0):
    #если пересчитываемый месяц больше текущего
        if year == tekyear and ((month == tekmonth) and (day > tekday)):
            module.log_info("arr_time_day: неверный день = %s"% day)
            return -1
        elif year == tekyear and month > tekmonth:
            module.log_info("arr_time_day: неверный месяц = %s"% month)
            return -1
        elif(year > tekyear):
            module.log_info("arr_time_day: неверный год = %s"% year)
            return -1
    else:
        day = tekday
        month = tekmonth
        year = tekyear

    #выбираем активным лист с нашим годом
    exel_sheet = year_sheet(year)
    sheet = exel_sheet[0]
    read_book = exel_sheet[2]
    #строковая заданная дата и следующая
    dan_date = str_date(day, month, year)
    #ищем наименование заданного дня в файле
    i = sheet.nrows-1
    #while sheet.row_values(i)[0] != dan_date and i > 0:
    temp = read_excel(read_book, sheet, i, 0)
    while temp[6] != dan_date and i > 0:
        i=i-1
        temp = read_excel(read_book, sheet, i, 0)
    if i <= 0:
        return -1

    index_sumS = i  #запоминаем строку, c началом заданного дня
    #ищем наименование следующего дня в файле (или конец файла)
    for i in range(index_sumS, sheet.nrows):
        #if (sheet.row_values(i)[0] != dan_date and sheet.row_values(i)[0] != ''):
        temp = read_excel(read_book, sheet, i, 0)
        if (temp[6] != dan_date and temp[0] != ''):
            index_sumE = i-1      #запоминаем строку, c началом следующего месяца месяца
            break
        else: index_sumE = i    #или запоминаем конец строки, если следующего месяца нет
    
    #обнуление времен
    timestart = []
    timeexit = []
    i = index_sumS
    #записываем времена
    for i in range(index_sumS, index_sumE+1):
        #если строка пустая, то ищем время дальше
        #if sheet.row_values(i)[1] == '':
        temp = read_excel(read_book, sheet, i, 1)
        if temp[0] == '':
            i=i+1
        else:
            #timestart.append(sheet.row_values(i)[1])
            timestart.append(temp[7])
            
    for i in range(index_sumS, index_sumE+1):
        #if sheet.row_values(i)[2] == '':
        temp = read_excel(read_book, sheet, i, 2)
        if temp[0] == '':
            i=i+1
        else:
            #timeexit.append(sheet.row_values(i)[2])
            timeexit.append(temp[7])
    #возвращаем массив времен
    return timestart, timeexit

#перевод даты из int в str, в формате dd.mm.yyyy
def str_date(day, month, year):
    day, month, year = int(day), int(month), int(year)
    if day < 10 and month < 10:
        return ('0'+str(day)+'.0'+str(month)+'.'+str(year))
    elif day < 10 and month >= 10:
        return('0'+str(day)+'.'+str(month)+'.'+str(year))
    elif day >= 10 and month < 10:
        return(''+str(day)+'.0'+str(month)+'.'+str(year))
    else:
        return(str(day)+'.'+str(month)+'.'+str(year))     

#перевод времени из int в str, в формате hh:mm
def str_time(hour, minute):
    if hour < 10 and minute < 10:
        return ('0'+str(hour)+':0'+str(minute))
    elif hour < 10 and minute >= 10:
        return('0'+str(hour)+':'+str(minute))
    elif hour >= 10 and minute < 10:
        return(''+str(hour)+':0'+str(minute))
    else:
        return(str(hour)+':'+str(minute))

#перевод времени в числовой формат с округлением до третьего знака
def convert_time(time):
    #выделяем часы и минуты прихода
    hr_start = time[0:2]  #часы
    min_start = time[3:5]  #минуты
    return round((int(hr_start) + int(min_start)/60),3)

#функция для вычисления часов в заданном дне
def time_in_day(timestart, timeexit):
    diner = int(module.read_setting(7))
    #переводим в численное представление времени
    sum_time = 0
    diner_time = 0
    for i in range(len(timestart)):
        timestart[i] = convert_time(timestart[i])
    for i in range(len(timeexit)):
        timeexit[i] = convert_time(timeexit[i])
    #вычисляем отработанное время в дне
    for i in range(len(timestart)):
        try:
            sum_time += (timeexit[i] - timestart[i])
        except:
            None
        #вычисляем обед
        try:
            diner_time += (timestart[i+1] - timeexit[i])
        except:
            None
    #если в дне отработано больше 4 часов - вычитаем обед из отработанных часов
    sum_time = round(sum_time,3)
    if sum_time > 4:
        if diner_time > diner/60:
            sum_time = round(sum_time,3)
        else:
            sum_time = round(sum_time - (diner/60 - diner_time),3)
    return sum_time

#функция для получения массива дней в месяце
def arr_day_month(month, year):
    #выбираем активным лист с нашим годом
    exel_sheet = year_sheet(year)
    sheet = exel_sheet[0]
    read_book = exel_sheet[2]
    
    #ищем наименование заданного месяца в файле
    i = sheet.nrows-1
    #while sheet.row_values(i)[0] != month_word[month-1]:
    temp = read_excel(read_book, sheet, i, 0)
    while temp[6] != month_word[month-1]:
        i=i-1
        temp = read_excel(read_book, sheet, i, 0)
    index_sumS = i+1  #запоминаем строку, c началом заданного месяца
    index_sumE = sheet.nrows-1
    #ищем наименование следующего месяца в файле (или конец файла)
    for i in range(index_sumS+1, sheet.nrows):
        temp = read_excel(read_book, sheet, i, 0)
        for j in range(len(month_word)):
            #if (sheet.row_values(i)[0] == month_word[j]):
            if (temp[6] == month_word[j]):
                index_sumE = i-1      #запоминаем строку, c началом следующего месяца
                break
            else:
                index_sumE = i    #или запоминаем конец строки, если следующего месяца нет
        #if (sheet.row_values(i)[0] == month_word[j]):  
        if (temp[6] == month_word[j]):
            break
    
    #от начала месяца до конца массив дней в месяце
    month_day = []
    for i in range(index_sumS, index_sumE+1):
        #if sheet.row_values(i)[0] == '':
        temp = read_excel(read_book, sheet, i, 0)
        if temp[0] == '':
            i=i+1
        else:
            #записываем в массив дней в месяце
            #month_day.append(sheet.row_values(i)[0])
            month_day.append(temp[6])
    return month_day

#функция для вывода номера дня аванса
def center_month(month_day):
    if len(month_day) >= 1:
        for i in range(1, len(month_day)):
            day_old = month_day[i-1]
            day = month_day[i]
            if (int(day[:2]) > 15) and (int(day_old[:2]) <= 15):
                return i-1
    else:
        day = month_day[0]
        if (int(day[:2]) >= 15):
            return 0
    return -1

#Функция для подсчета часов в месяце
def month_recount(month_day, num_cent_day):
    sum_month_time = 0
    sum_center_month = 0
    sum_day_time = []
    time = []
    for i in range(len(month_day)):
        #получаем массив времен в дне
        day = month_day[i]
        time = arr_time_day(int(day[:2]), int(day[3:5]), int(day[6:]))
        #если массив получен нормально, вычисляем сумму в месяце и массив рабочих часов в дне
        if time != -1:
            day_time = time_in_day(time[0], time[1])
            sum_day_time.append(day_time)
            sum_month_time = sum_month_time + day_time
            if i == num_cent_day:
                sum_center_month = sum_month_time
    return round(sum_month_time, 3), round(sum_center_month, 3), sum_day_time

#Функция получения массива месяцев
def arr_month_year(year):
    date = datetime.datetime.now()
    tekyear = date.year #текущий месяц
    
    #если пересчитываемый месяц больше текущего
    if year > tekyear:
        module.log_info('arr_month_year: неверный год %s' %year)
        return 1

    #выбираем активным лист с нашим годом
    exel_sheet = year_sheet(year)
    sheet = exel_sheet[0]
    read_book = exel_sheet[2]
    
    #ищем наименование месяцев в файле
    year_month = []
    for i in range(sheet.nrows):
        for j in range(len(month_word)):
            #if sheet.row_values(i)[0] == month_word[j]:
            temp = read_excel(read_book, sheet, i, 0)
            if temp[6] == month_word[j]:
                #записываем числовые значения месяца
                year_month.append(j+1)
                break
            j=j+1
        i=i+1
    return year_month

#Функция для подсчета часов в каждом месяце года  
def year_recount(year):
    #выбираем активным лист с нашим годом
    exel_sheet = year_sheet(year)
    #получаем массив месяцев
    year_month = arr_month_year(year)
    #в цикле вычисляем количество рабочих часов в каждом из месяцев
    for i in range(len(year_month)):
        #получаем массив дней в месяце
        day_in_mount = arr_day_month(year_month[i],year)
        #получаем номер середины месяца для подсчета аванса
        num_center_day = center_month(day_in_mount)
        #получаем количество часов в месяце и до аванса, и массив часов работы в дне
        sum_month = month_recount(day_in_mount, num_center_day)
        #записываем в Exel 
        if(write_sum_month(sum_month[0], sum_month[1], num_center_day, sum_month[2], year_month[i], year) == False):
            return False
    return True    

#запись в Exel файл массива часов работы в месяце и массива часов работы по дням в месяце
def write_sum_month(sum_month, sum_cnt_month, num_cnt_day, day_in_mount, month, year):
    #получаем путь к файлу и смещение
    wt_filename = module.read_setting(4) + '/' + module.read_setting(1)
    #выбираем активным лист с нашим годом
    #открываем наш Exel файл
    read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    write_book = copy(read_book)
    
    #Переходим на лист текущего года
    try:
        #если лист с текущим годом уже существует
        sheet_index = read_book.sheet_names().index(str(year))
    except:
        #если листа текущего года нет значит программа start не сработала
        module.log_info('year_sheet: в Exel файле отсутствует страница %s' %year)
        return False
    
    #выбираем активным лист с нашим годом
    sheet = read_book.sheet_by_index(sheet_index)
    
    #поиск заданного месяца в файле
    i = sheet.nrows-1
    #while sheet.row_values(i)[0] != month_word[month-1]:
    temp = read_excel(read_book, sheet, i, 0)
    while temp[6] != month_word[month-1]:
        i=i-1
        temp = read_excel(read_book, sheet, i, 0)
    index_month = i  #индекс строки месяца
    #заполняем сумму часов в соответствующую строку
    write_book.get_sheet(sheet_index).write(index_month,1,str(sum_month))
    
    #начиная с начала месяца и до конца массива, записываем часы
    i_day = index_month + 1
    i = 0
    while (i < len(day_in_mount)):
    #for i in range(len(day_in_mount)):
        #if sheet.row_values(i_day)[0] == '':
        temp = read_excel(read_book, sheet, i_day, 0)
        if temp[0] == '':
            i_day=i_day+1
        else:
            #записываем часы в дне
            write_book.get_sheet(sheet_index).write(i_day,3,str(day_in_mount[i]))
            #записываем отработанное кол-во часов до аванса
            if i == num_cnt_day:
                write_book.get_sheet(sheet_index).write(i_day,4,'('+ str(sum_cnt_month)+')')
            i_day=i_day+1
            i = i+1
    
    #сохраняем запись
    try:
        write_book.save(wt_filename)
    except Exception as e:
        module.log_info("Не удалось сохранить в Exel. Exception: %s" % str(e))
        return False
    return True

#функция для получения массива с годами, хранящимися в exel файле
def exel_year():
    #получаем путь к файлу и смещение
    wt_filename = module.read_setting(4) + '/' + module.read_setting(1)
    #открываем наш Exel файл
    read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    #получаем массив годов
    return read_book.sheet_names()

#создание Excel файла
def new_timework_file(path):
    try:
        book = xlwt.Workbook('utf-8')   #создали книку
        sheetname = datetime.datetime.now().year
        sheet = book.add_sheet(str(sheetname))  #создали страницу с текущим годом
        sheet.portrain = False  #не альбомная ориентация
        book.save(path)     #сохраняем файл
        return 0
    except Exception as e:
        module.log_info("new_timework_file: Не удалось создать Exel файл. Exception: %s" % str(e))
        return -1
    
