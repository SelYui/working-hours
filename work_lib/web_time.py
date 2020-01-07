# -*- coding: cp1251 -*-
'''
������ ��� ��������� �������� ������� � ����� �����
'''
import time, requests
from work_setting import module
from work_lib import work_time, shutdown_lib
import work_lib

#����� ������ � ����� ������� �������� �������
symbol = ['<table border="1" data-count="','<th class="header" colspan="3">','<tr class="header">','</table>','<tr>','</tr>','<th>', '</th>','</td>','<td>','">',]
#������ ��� ������ ����
datemass = "����� �� ��������� - ���������� �� ����������� "
#�� ���� ����� �������� ���
url_mars = 'http://www.mars/asu/report/enterexit/'
#�� ���� ����� �������� ����� �������
url = 'getinfo.php'

#������� ��� ��������� ������ � ����� ������� � ����� �����
def you_split_time(my_data):
    #���������� ������ �� ��������� ������ ��� ������������� ���������
    try:
        rp = requests.post(url_mars+url, data = my_data)
    except:
        return 1
    lines = rp.text
    
    #������� �������� ������� �� ������
    for i in range(len(symbol)):
        lines = lines.replace(symbol[i],'')
    
    #�������� ������ ����
    lines = lines.split()
    
    #��������� ������� ��������
    if len(lines) < 7:
        return 2
    
    return lines

#�������� ������� ������� �������/����� (�������� �� ������)
def came_left(lines, timestart, timeexit):
    count = 7
    while count < len(lines):
        #����� �������
        try:
            start_tek = lines[count]
            timestart.append(start_tek[:5]) #���������� ������ ���� � ������
        except IndexError:
            return 1
    
        #����� ������� �����
        try:
            exit_tek = lines[count+1]
            timeexit.append(exit_tek[:5])   #���������� ������ ���� � ������
        except IndexError:
            return 2
    
        count = count+3
    
    return count

def web_main():
    #�������� ���� � �����
    rp = requests.post(url_mars)     #�������� ����
    lines = rp.text                 
    num = lines.find(datemass)      #���� ������ � �����
    date = lines[num+47:num+57]     #�������� ����
    tekday = int(date[8:10])
    tekmonth = int(date[5:7])
    tekyear = int(date[0:4])
    
    #�������� ����� ����� �������� �������� ���������
    max_timE = module.read_timeShut()
    
    #��������� �������
    name_id = module.read_number()
    my_data = {'type': 'search', 'info': name_id}
    
    while True:
        timestart = []
        timeexit = []
        lines = you_split_time(my_data)

        if lines == 1:
            module.log_info('������ �� ��������')
        elif lines == 2:
            msg = ('������� �������� � �������', name_id, '�� ������')
            module.log_info(msg)

        #������ ������� �������� �� ����������� (��� ������, �.�. ���� ������ ������ ��������)
        tek_status = int(lines[0])

        #�������� ������� ������� - �����
        status_availability = came_left(lines, timestart, timeexit)
        
        if (status_availability == 1):      #���� ������� ������� ���
            msg = ('���������',name_id, lines[1], lines[2],'��� �� ������')
            module.log_info(msg)
        elif (status_availability == 2):    #���� ������� ����� ���
            msg = ('���������',name_id, lines[1], lines[2],'��� �� ����')
            module.log_info(msg)
        print('start =', timestart)
        print('exit =', timeexit)
        
        #�������� ��������� ����� �������, ���������� ��� � Exel
        try:
            timeS = timestart[-1]
            #���������� ����� ������� �� ������
            work_time.start_work(int(timeS[3:5]), int(timeS[0:2]), tekday, tekmonth, tekyear)
        except:
            None
        
        #�������� ��������� ����� �����, ���������� ��� � Exel
        try:
            timeE = timeexit[-1]
            #���������� ����� �����
            work_time.exit_work(int(timeE[3:5]), int(timeE[0:2]), tekday, tekmonth, tekyear)
        except:
            None
        
        print(timestart[-1] ,max_timE)
        
        try:
            #������� ���������� ����������
            if(timeexit[-1] > max_timE):
                print("����������")
                shutdown_lib.signal_shutdown()
                break
            else:
                print("�� ����������")
        except:
            None
        time.sleep(60)
    