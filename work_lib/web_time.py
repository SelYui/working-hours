# -*- coding: cp1251 -*-
'''
Модуль для получения рабочего времени с сайта марса
'''
import time, datetime, requests
from work_setting import dialog, module, adjacent_classes
from work_lib import work_time, shutdown_lib

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
    date = url_date()
    tekday = int(date[8:10])    #текущий день
    tekmonth = int(date[5:7])   #текущий месяц
    tekyear = int(date[0:4])    #текущий год
    print(date, tekday, tekmonth, tekyear)
    
    #получаем время после которого выключим компьютер
    max_timE = module.read_timeShut()
    #начальные условия
    name_id = module.read_number()
    my_data = {'type': 'search', 'info': name_id}
    while True:
        print(10)
        #также записываем текущее время в файл, что бы в случае сбоя записать его в Exel файл
        #получаем текущее время
        timeExit = datetime.datetime.now()
        #записываем текущее время в файл
        module.write_timeExit(timeExit.strftime("%d %m %Y %H:%M"))
        
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
        '''
        if (status_availability == 1):      #если времени прихода нет
            msg = 'сотрудник ' + lines[2] + lines[1] + ' еще не пришел'
            module.log_info(msg)
        elif (status_availability == 2):    #если времени ухода нет
            msg = 'сотрудник ' + lines[2] + lines[1] + ' еще не ушел'
            module.log_info(msg)
        
        print('start =', timestart)
        print('exit =', timeexit)
        '''
        #timeexit = ['18:40']
        
        #записываем массив времен прихода (если вдруг комп включили после нескольких заходов на марс)
        i = 0
        while(i < len(timestart)):
            print('cycl = ',i, len(timestart))
            #получаем последнее время прихода, записываем его в Exel
            try:
                timeS = timestart[i]
                #записываем время прихода на работу
                work_time.start_work(int(timeS[3:5]), int(timeS[0:2]), tekday, tekmonth, tekyear)
            except:
                msg = ('не удалось записать время прихода')
                module.log_info(msg)
            print(12)
            try:
                print(13)
                timeE = timeexit[i]
                #записываем время ухода
                work_time.exit_work(int(timeE[3:5]), int(timeE[0:2]), tekday, tekmonth, tekyear)
            except:
                print(14)
                msg = ('не удалось записать время ухода')
                module.log_info(msg)
            print(timestart, timeexit)
            print('end_cycl')
            i=i+1

        try:
            #условие выключения компьютера
            if(timeexit[-1] > max_timE):
                print("Выключаюсь")
                #выставляем признак завершения (для потока с таймером завершения)
                module.write_setting(0, 28)    #ставим признак штатного завершения
                return
            else:
                print("Не выключаюсь")
        except:
            None
        print(11)
        time.sleep(60)

#функция для получения даты с сайта
def url_date():
    rp = requests.post(url_mars)     #копируем сайт
    lines = rp.text                 
    num = lines.find(datemass)      #ищем строку с датой
    return lines[num+47:num+57]     #выделяем дату
    