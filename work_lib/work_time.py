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
    print('start', sheet_nrows)
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
                if time_compare(sheet_nrows, sheet, 1, tekhour, tekminute, time_date_index) == 0:    #если такая дата уже записанна, то ничего не делаем
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
        module.log_info('exit_work: в Exel файле отсутствует страница %s' %tekyear)
        return
    
    #выбираем активным лист с нашим годом
    sheet = read_book.sheet_by_index(sheet_index)
    print('sheet123 = ', sheet)
    print('exit', sheet.nrows)
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
        else:
            module.log_info('exit_work: нет текущей даты')
            return
    else:
        module.log_info('exit_work: нет текущего мксяца')
        return
    
    #сравниваем текущую дату прихода со вторым (j) столбцом
    if time_compare(sheet.nrows, sheet, 2, tekhour, tekminute, time_date_index) == 0:    #если такая дата уже записанна, то ничего не делаем
        return
    print(10)
    #заполняем строку временем
    if tekhour < 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,2,'0'+str(tekhour)+':0'+str(tekminute))
    elif tekhour < 10 and tekminute >= 10:
        write_book.get_sheet(sheet_index).write(i,2,'0'+str(tekhour)+':'+str(tekminute))
    elif tekhour >= 10 and tekminute < 10:
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

#сравниваем время текущее с временем в Exel, если совпало или меньше, то не будем записывать
def time_compare(sheet_nrows, sheet, j, tekhour, tekminute, time_date_index):
    shour = sminute = 0
    print("я тута %s"%j)
    #получаем последнее время прихода
    i = sheet_nrows-1
    print('timcomp', sheet_nrows)
    print(i, time_date_index)
    #пока мы в текущем дне
    
    while (i >= time_date_index):#i > 0:
        print(1)
        lastdate = sheet.row_values(i)[j]
        print(2, lastdate)
        #если строка пустая, то ищем дату выше
        if lastdate == '':
            print(3)
            i=i-1
        else:
            print(4)
            shour = lastdate[0:2]  #день
            sminute = lastdate[3:5]  #месяц
            break
    print(shour, sminute)
    #если время совпало с последним записанным временем - ничего не делаем. Выходим из программы
    if ((int(shour) == tekhour) and (int(sminute) >= tekminute)) or ((int(shour) > tekhour)):
        print('вышел 0', shour, sminute, tekhour, tekminute)
        return 0    #текущая дата совпадает с последней записанной
    else:
        print('вышел 1')
        return 1  #не совпадает
        
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

#записываем в Exel файл время последнего выключения компьютера
def write_exit():
        
    dtimE = module.read_timeExit()
    dtimE = dtimE.split()
    #если приложение закрыл не пользователь
    if (dtimE != ''):
        timE = dtimE[-1]
        print('записываю %s'%timE)
        exit_work(int(timE[3:5]), int(timE[0:2]), int(dtimE[0]), int(dtimE[1]), int(dtimE[2]))
    else:
        msg = "dtimE = %s"% dtimE
        module.log_info(msg)

#получаем лист из Exel файла
def year_sheet(year):
    #получаем путь к файлу и смещение
    wt_filename = module.read_path() + '/' + module.read_name()
    
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
    print('sheet0 = ', read_book.sheet_by_index(sheet_index))
    return read_book.sheet_by_index(sheet_index), write_book

#функция получения массива времян в заданном дне
def arr_time_day(day, month, year):
    date = datetime.datetime.now()
    tekmonth = date.month #текущий месяц
    tekyear = date.year
    tekday = date.day
    
    #если пересчитываемый месяц больше текущего
    if year == tekyear and ((month == tekmonth) and (day > tekday)):
        module.log_info("неверный день = %s"% day)
        return 1
    elif year == tekyear and month > tekmonth:
        module.log_info("неверный месяц = %s"% month)
        return 1
    elif(year > tekyear):
        module.log_info("неверный год = %s"% year)
        return 1

    #выбираем активным лист с нашим годом
    exel_sheet = year_sheet(year)
    sheet = exel_sheet[0]
    #строковая заданная дата и следующая
    dan_date = str_date(day, month, year)
    next_date = str_date(day+1, month, year)
    
    #ищем наименование заданного дня в файле
    i = sheet.nrows-1
    while sheet.row_values(i)[0] != dan_date and i > 0:
        i=i-1
    if i == 0:
        return 1
        
    index_sumS = i  #запоминаем строку, c началом заданного дня
    #ищем наименование следующего дня в файле (или конец файла)
    for i in range(index_sumS, sheet.nrows):
        if (sheet.row_values(i)[0] != dan_date and sheet.row_values(i)[0] != ''):
            index_sumE = i-1      #запоминаем строку, c началом следующего месяца месяца
            break
        else: index_sumE = i    #или запоминаем конец строки, если следующего месяца нет
    print(index_sumS, index_sumE)
    
    #обнуление времен
    timestart = []
    timeexit = []
    i = index_sumS
    #записываем времена
    for i in range(index_sumS, index_sumE+1):
        #если строка пустая, то ищем время дальше
        if sheet.row_values(i)[1] == '':
            i=i+1
        else:
            timestart.append(sheet.row_values(i)[1])
    for i in range(index_sumS, index_sumE+1):
        if sheet.row_values(i)[2] == '':
            i=i+1
        else:
            timeexit.append(sheet.row_values(i)[2])
    #возвращаем массив времен
    print('timestart =', timestart)
    print('timeexit =', timeexit)
    return timestart, timeexit

#перевод даты из int в str, в формате dd.mm.yyyy
def str_date(day, month, year):
    if day < 10 and month < 10:
        return ('0'+str(day)+'.0'+str(month)+'.'+str(year))
    elif day < 10 and month >= 10:
        return('0'+str(day)+'.'+str(month)+'.'+str(year))
    elif day >= 10 and month < 10:
        return(''+str(day)+'.0'+str(month)+'.'+str(year))
    else:
        return(str(day)+'.'+str(month)+'.'+str(year))     

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
        except Exception as e:
            #print('time_in_day 1 = %s', e)
            None
        #вычисляем обед
        try:
            diner_time += (timestart[i+1] - timeexit[i])
        except Exception as e:
            #print('time_in_day 2= %s', e)
            None
    #вычитаем обед из отработанных часов
    if diner_time > diner/60:
        sum_time = round(sum_time,3)
    else:
        sum_time = round(sum_time - (diner/60 - diner_time),3)
    print('sum_time =', sum_time)    
    return sum_time

#функция для получения массива дней в месяце
def arr_day_month(month, year):
    date = datetime.datetime.now()
    tekmonth = date.month #текущий месяц
    
    #если пересчитываемый месяц больше текущего
    #if month > tekmonth:
    #    module.log_info('mount_recount: неверный месяц %s' %month)
    #    return 1

    #выбираем активным лист с нашим годом
    exel_sheet = year_sheet(year)
    sheet = exel_sheet[0]
    
    #ищем наименование заданного месяца в файле
    i = sheet.nrows-1
    while sheet.row_values(i)[0] != month_word[month-1]:
        i=i-1
    index_sumS = i+1  #запоминаем строку, c началом заданного месяца
    #ищем наименование следующего месяца в файле (или конец файла)
    for i in range(index_sumS+1, sheet.nrows):
        for j in range(len(month_word)):
            if (sheet.row_values(i)[0] == month_word[j]):
                print('!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
                index_sumE = i-1      #запоминаем строку, c началом следующего месяца
                break
            else:
                index_sumE = i    #или запоминаем конец строки, если следующего месяца нет
        if (sheet.row_values(i)[0] == month_word[j]):
            print('*********************************************')
            break
    
    #от начала месяца до конца массив дней в месяце
    month_day = []
    for i in range(index_sumS, index_sumE+1):
        if sheet.row_values(i)[0] == '':
            i=i+1
        else:
            #записываем в массив дней в месяце
            month_day.append(sheet.row_values(i)[0])
    print('month_day =', month_day)
    return month_day

#функция для вывода номера дня аванса
def center_month(month_day):
    if len(month_day) > 1:
        for i in range(1, len(month_day)):
            day_old = month_day[i-1]
            day = month_day[i]
            if (int(day[:2]) > 15) and (int(day_old[:2]) <= 15):
                print('center_month = ', i-1)
                return i-1
            
    return 0

#Функция для подсчета часов в месяце
def month_recount(month_day, num_cent_day):
    sum_month_time = 0
    sum_day_time = []
    time = []
    print 
    for i in range(len(month_day)):
        print('i=',i)
        #получаем массив времен в дне
        day = month_day[i]
        time = arr_time_day(int(day[:2]), int(day[3:5]), int(day[6:]))
        #если массив получен нормально, вычисляем сумму в месяце и массив рабочих часов в дне
        if time != 1:
            day_time = time_in_day(time[0], time[1])
            sum_day_time.append(day_time)
            sum_month_time = sum_month_time + day_time
            if i == num_cent_day:
                sum_center_month = sum_month_time
    print('sum_month_time =', sum_month_time)   #сумма часов в месяце
    print('sum_day_time =', sum_day_time)       #массив сумм часов в дне
    print('sum_center_month =', sum_center_month)       #сумма часов в центре месяца
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
    
    #ищем наименование месяцев в файле
    year_month = []
    for i in range(sheet.nrows):
        for j in range(len(month_word)):
            if sheet.row_values(i)[0] == month_word[j]:
                #записываем числовые значения месяца
                year_month.append(j+1)
                break
            j=j+1
        i=i+1
    return year_month

#Функция для подсчета часов в каждом месяце года  
def year_recount(year):
    #получаем массив месяцев
    year_month = arr_month_year(year)
    print('year_month =', year_month, len(year_month))
    #в цикле вычисляем количество рабочих часов в каждом из месяцев
    for i in range(len(year_month)):
        print('сейчас считаю %s месяц'% year_month[i])
        #получаем массив дней в месяце
        day_in_mount = arr_day_month(year_month[i],year)
        #получаем номер середины месяца для подсчета аванса
        num_center_day = center_month(day_in_mount)
        #получаем количество часов в месяце и до аванса, и массив часов работы в дне
        sum_month = month_recount(day_in_mount, num_center_day)
        print('sum_month =', sum_month, len(sum_month))
        #записываем в Exel 
        write_sum_month(sum_month[0], sum_month[1], num_center_day, sum_month[2], year_month[i], year)

#запись в Exel файл массива часов работы в месяце и массива часов работы по дням в месяце
def write_sum_month(sum_month, sum_cnt_month, num_cnt_day, day_in_mount, month, year):
    #получаем путь к файлу и смещение
    wt_filename = module.read_path() + '/' + module.read_name()
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
        return 0
    
    #выбираем активным лист с нашим годом
    sheet = read_book.sheet_by_index(sheet_index)
    
    '''
    exel_sheet = year_sheet(year)
    sheet = exel_sheet[0]
    write_book = exel_sheet[1]
    '''
    print('sheet = ', sheet)
    #поиск заданного месяца в файле
    i = sheet.nrows-1
    while sheet.row_values(i)[0] != month_word[month-1]:
        i=i-1
    index_month = i  #индекс строки месяца
    #заполняем сумму часов в соответствующую строку
    write_book.get_sheet(sheet_index).write(index_month,1,str(sum_month))
    
    #начиная с начала месяца и до конца массива, записываем часы
    i_day = index_month + 1
    i = 0
    while (i < len(day_in_mount)):
    #for i in range(len(day_in_mount)):
        if sheet.row_values(i_day)[0] == '':
            i_day=i_day+1
        else:
            #записываем часы в дне
            write_book.get_sheet(sheet_index).write(i_day,3,str(day_in_mount[i]))
            #записываем отработанное кол-во часов до аванса
            if i == num_cnt_day:
                write_book.get_sheet(sheet_index).write(i_day,4,str(sum_cnt_month))
            i_day=i_day+1
            i = i+1
    
    #сохраняем запись
    try:
        write_book.save(wt_filename)
    except Exception as e:
        module.log_info("Не удалось сохранить в Exel. Exception: %s" % str(e))
    return 1
    

