# -*- coding: cp1251 -*-
'''
Модуль для получения рабочего времени с сайта марса
'''
import time, requests
from work_setting import module
from work_lib import work_time, shutdown_lib
import work_lib

#после чтения с сайта удаляем ненужные символы
symbol = ['<table border="1" data-count="','<th class="header" colspan="3">','<tr class="header">','</table>','<tr>','</tr>','<th>', '</th>','</td>','<td>','">',]
#массив для поиска даты
datemass = "Отчёт по проходной - сотрудники на предприятии "
#на этом сайте хранится все
url_mars = 'http://www.mars/asu/report/enterexit/'
#на этом сайте хранится время прихода
url = 'getinfo.php'

#функция для получения строки с моими данными с сайта Марса
def you_split_time(my_data):
    #отправляем запрос на получение данных для определенного работника
    try:
        rp = requests.post(url_mars+url, data = my_data)
    except:
        return 1
    lines = rp.text
    
    #удаляем ненужные символы из ответа
    for i in range(len(symbol)):
        lines = lines.replace(symbol[i],'')
    
    #получаем массив слов
    lines = lines.split()
    
    #проверяем наличие человека
    if len(lines) < 7:
        return 2
    
    return lines

#получаем массивы времени прихода/ухода (точность до минуты)
def came_left(lines, timestart, timeexit):
    count = 7
    while count < len(lines):
        #время прихода
        try:
            start_tek = lines[count]
            timestart.append(start_tek[:5]) #записываем только часы и минуты
        except IndexError:
            return 1
    
        #время первого ухода
        try:
            exit_tek = lines[count+1]
            timeexit.append(exit_tek[:5])   #записываем только часы и минуты
        except IndexError:
            return 2
    
        count = count+3
    
    return count

def web_main():
    #получаем дату с сайта
    rp = requests.post(url_mars)     #копируем сайт
    lines = rp.text                 
    num = lines.find(datemass)      #ищем строку с датой
    date = lines[num+47:num+57]     #выделяем дату
    tekday = int(date[8:10])
    tekmonth = int(date[5:7])
    tekyear = int(date[0:4])
    
    #получаем время после которого выключим компьютер
    max_timE = module.read_timeShut()
    
    #начальные условия
    name_id = module.read_number()
    my_data = {'type': 'search', 'info': name_id}
    
    while True:
        timestart = []
        timeexit = []
        lines = you_split_time(my_data)

        if lines == 1:
            module.log_info('сервер не отвечает')
        elif lines == 2:
            msg = ('пропуск человека с номером', name_id, 'не пробит')
            module.log_info(msg)

        #статус наличия человека на предприятии (это лишнее, т.к. есть другие методы проверки)
        tek_status = int(lines[0])

        #получаем массивы прихода - ухода
        status_availability = came_left(lines, timestart, timeexit)
        
        if (status_availability == 1):      #если времени прихода нет
            msg = ('сотрудник',name_id, lines[1], lines[2],'еще не пришел')
            module.log_info(msg)
        elif (status_availability == 2):    #если времени ухода нет
            msg = ('сотрудник',name_id, lines[1], lines[2],'еще не ушел')
            module.log_info(msg)
        print('start =', timestart)
        print('exit =', timeexit)
        
        #получаем последнее время прихода, записываем его в Exel
        try:
            timeS = timestart[-1]
            #записываем время прихода на работу
            work_time.start_work(int(timeS[3:5]), int(timeS[0:2]), tekday, tekmonth, tekyear)
        except:
            None
        
        #получаем последнее время ухода, записываем его в Exel
        try:
            timeE = timeexit[-1]
            #записываем время ухода
            work_time.exit_work(int(timeE[3:5]), int(timeE[0:2]), tekday, tekmonth, tekyear)
        except:
            None
        
        print(timestart[-1] ,max_timE)
        
        try:
            #условие выключения компьютера
            if(timeexit[-1] > max_timE):
                print("Выключаюсь")
                shutdown_lib.signal_shutdown()
                break
            else:
                print("Не выключаюсь")
        except:
            None
        time.sleep(60)
    