# -*- coding: utf-8 -*-
'''
Модуль для получения рабочего времени с сайта
'''
import time, datetime, requests
from work_setting import module
from work_lib import work_time

#после чтения с сайта удаляем ненужные символы
symbol = ['<table border="1" data-count="','<th class="header" colspan="3">','<tr class="header">','</table>','<tr>','</tr>','<th>', '</th>','</td>','<td>','">',]
#массив для поиска даты
datemass = "Отчёт по проходной - сотрудники на предприятии "
#на этом сайте хранится все
url_mars = 'http://www.mars/asu/report/enterexit/'
#на этом сайте хранится время прихода
url = 'getinfo.php'

#функция для получения строки с данными с сайта Марса
def you_split_time(my_data):
    #отправляем запрос на получение данных для определенного работника
    try:
        rp = requests.post(url_mars+url, data = my_data)
    except:
        return -1
    lines = rp.text
    #удаляем ненужные символы из ответа
    for i in range(len(symbol)):
        lines = lines.replace(symbol[i],'')
    #получаем массив слов
    lines = lines.split()
    #проверяем наличие человека
    if len(lines) < 7:
        return -2
    
    return lines

#получаем массивы времени прихода/ухода (точность до минуты)
def came_left(lines, timestart, timeexit):
    for count in range(6,len(lines)):
        if lines[count] == 'Пришел' or lines[count] == 'Ушел':
            #время прихода
            try:
                start_tek = lines[count+1]
                if (start_tek != 'Пришел') and (start_tek != 'Ушел') and (not start_tek[:5] in timestart):
                    timestart.append(start_tek[:5]) #записываем только часы и минуты
            except IndexError:
                return -1
    
            #время первого ухода
            try:
                exit_tek = lines[count+2]
                if (exit_tek != 'Пришел') and (exit_tek != 'Ушел') and (not exit_tek[:5] in timeexit):
                    timeexit.append(exit_tek[:5])   #записываем только часы и минуты
            except IndexError:
                return -2
    
    return count

#основная функция для работы с данными с сайта
def web_main():
    #return True
    # обнуление счетчика опроса сайта
    count = 0
    
    #получаем дату с сайта
    date = url_date()
    if date == -1:
        module.log_info('на сайте дата не найдена')
        #используем дату компьютера
        data = datetime.datetime.now()
        tekyear = data.year   #Текущий год
        tekmonth = data.month #текущий месяц
        tekday = data.day     #текущее число
    else:
        tekday = int(date[8:10])    #текущий день
        tekmonth = int(date[5:7])   #текущий месяц
        tekyear = int(date[0:4])    #текущий год
    
    #получаем время после которого выключим компьютер
    max_timE = module.read_setting(22)
    #составляем запрос для сайта
    name_id = module.read_setting(19)
    my_data = {'type': 'search', 'info': name_id}
    
    while True:
        #также записываем текущее время в файл, что бы в случае сбоя записать его в Exel файл
        CurrentTime = datetime.datetime.now()      #получаем текущее время
        #обнуляем массивы перед чтением с сайта
        timestart = []
        timeexit = []
        
        if CurrentTime.strftime("%H:%M") < max_timE:
            count = 60         #заряжаем новый таймер на час
        else:
            count = 10         #заряжаем новый таймер на 10 минут
        lines = you_split_time(my_data)
        if lines == -1:
            module.log_info('сервер не отвечает')
        elif lines == -2:
            msg = ('пропуск человека с номером %s не пробит'% name_id)
            module.log_info(msg)
        else:
            #статус наличия человека на предприятии (это лишнее, т.к. есть другие методы проверки)
            tek_status = int(lines[0])
            #получаем массивы прихода - ухода
            status_availability = came_left(lines, timestart, timeexit)
            #записываем массив времен прихода (если вдруг комп включили после нескольких заходов на марс)
            #технология для отладки
            #timestart = ['07:00','08:00','09:00','10:00','11:00','12:00']
            #timeexit = ['07:30','08:30','09:30','10:30','11:30','12:30']
            for i in range(len(timestart)):
                #получаем время прихода, записываем его в Exel
                try:
                    timeS = timestart[i]
                    #записываем время прихода на работу
                    work_time.start_work(int(timeS[3:5]), int(timeS[0:2]), tekday, tekmonth, tekyear)
                except:
                    None
                    #module.log_info('не удалось записать время прихода')
                try:
                    timeE = timeexit[i]
                    #записываем время ухода
                    work_time.exit_work(int(timeE[3:5]), int(timeE[0:2]), tekday, tekmonth, tekyear)
                except:
                    None
                    #module.log_info('не удалось записать время ухода')
                i=i+1
    
            try:
                #проверяем условие выключения компьютера
                if(timeexit[-1] > max_timE):
                    #выставляем признак завершения (для потока с таймером завершения)
                    return True
                else:
                    None
            except:
                None
        time.sleep(60*count)          #пауза в 60 секунд что бы каждую минуту записывать время компьютера
    
#функция для получения даты с сайта
def url_date():
    try:
        rp = requests.post(url_mars)     #копируем сайт
    except:
        return -1
    lines = rp.text                 
    num = lines.find(datemass)      #ищем строку с датой
    return lines[num+47:num+57]     #выделяем дату
    