# -*- coding: cp1251 -*-
'''
������ ��� ������ � ������ �� ������������ ����� work_setting.txt
'''
import os, datetime

#������ � ��� ���� ����� + ���� ����
def log_info(msg):
    ''' ������ ������
    module.log_info("date: %s" % date)
    '''
    #�������� ����� ����
    log_time = datetime.datetime.now()
    #���������� ���� � ����
    f = open("work_setting\working_hour.log", "a")
    f.write(msg + "            " + str(log_time) + "\n")
    f.close()
    
#�������� ��� ����� Exel �� ������ ������������ �����
def read_name():
    #������ ����������� ��� ����
    WorkFile = '�������_�����.xls'
    
    #������ ���� ��������� � ���������� ������ ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines_name = f.readlines()
        try:
            WorkFile = lines_name[1]
        except:
            None
    #��� ������� �������� ������
    return WorkFile[:-1]

#������ � ����������� ���� ������ ����� �����
def write_name(new_name):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� ������ ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= 1:
        lines[1] = new_name + '\n'
    #��������� � ����� ������ ����� ����
    else:
        lines.append(new_name)
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#�������� ���� � ����� Exel �� ������ ������������ �����
def read_path():
    #������ ������� ������� ��� ������
    WorkPath = os.path.dirname(os.path.realpath(__file__)) + '\�������_�����.xls'
    
    #������ ���� ��������� � ���������� ������ ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            WorkPath = lines[4]
        except:
            None
    #��� ������� �������� ������
    return WorkPath[:-1]

#������ � ����������� ���� ������ ����
def write_path(new_path):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� ��������� ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= 4:
        lines[4] = new_path + '\n'
    #��������� � ����� ������ ����� ����
    else:
        lines.append(new_path)
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#��������� �� ����� ��������
def read_offset():
    #������ ������� ������� ��� ������
    WorkOffset = 0
    
    #������ ���� ��������� � ���������� ������ ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            WorkOffset = lines[10]
        except:
            None
    #��� ������� �������� ������
    return int(WorkOffset)
    
#���������� �������� � ����
def write_offset(new_offset):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� ������� ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= 10:
        lines[10] = str(new_offset) + '\n'
    #��������� � ����� ������ ����� ����
    else:
        lines.append(str(new_offset))
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#��������� �� ����� ���������� ����� �����
def read_reload():
    #��������
    WorkReload = 0
    
    #������ ���� ��������� � ���������� 10 ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            WorkReload = lines[13]
        except:
            None
    
    return int(WorkReload)
    
#���������� �������� � ����
def write_reload(new_reload):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� 10 ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= 13:
        lines[13] = str(new_reload) + '\n'
    #��������� � ����� ������ ����� ����
    else:
        lines.append(str(new_reload))
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()


#��������� �� ����� ��������� ��� ��� ���� ������ � �����
def read_check():
    #������ ������� ������� ��� ������
    CheckNum = 0 #�� ���������
    
    #������ ���� ��������� � ���������� 13 ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            CheckNum = lines[16]
        except:
            None
    #��� ������� �������� ������
    return int(CheckNum)
    
#���������� ���� ������ � ����� � ����
def write_checkt(new_check):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� ������� ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= 16:
        lines[16] = str(new_check) + '\n'
    #��������� � ����� ������ ����� ����
    else:
        lines.append(str(new_check))
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#��������� �� ����� ���� ����������
def read_number():
    #������ ������� ������� ��� ������
    CheckNum = 0 #�� ���������
    
    #������ ���� ��������� � ���������� 16 ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            CheckNum = lines[19]
        except:
            None
    #��� ������� �������� ������
    return CheckNum[:-1]
    
#���������� � ���� ����� ����������
def write_number(new_number):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� 16 ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= 19:
        lines[19] = str(new_number) + '\n'
    #��������� � ����� ������ ����� ����
    else:
        lines.append(str(new_number))
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()

#������ � ����� ������������ �����
def write_timeShut(timeShut):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� 19 ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= 22:
        lines[22] = str(timeShut) + '\n'
    #��������� � ����� ������ ����� ����
    else:
        lines.append(str(timeShut))
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()
    
#��������� �� ����� ��������� ������ (������ ����� � �����)
def read_timeShut():
    #�������� ����������
    timeShut = 0 #�� ���������
    
    #������ ���� ��������� � ���������� 19 ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            timeShut = lines[22]
        except:
            None
    #��� ������� �������� ������
    return timeShut[:-1]

#������ � ����������� ���� ������� ����������
def write_timeExit(timeExit):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� 22 ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= 25:
        lines[25] = str(timeExit) + '\n'
    #��������� � ����� ������ ����� ����
    else:
        lines.append(str(timeExit))
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()
    
#��������� �� ����� ��������� ������ (������ ����� � �����)
def read_timeExit():
    #�������� ����������
    timeExit = 0 #�� ���������
    
    #������ ���� ��������� � ���������� 22 ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            timeExit = lines[25]
        except:
            None
    #��� ������� �������� ������
    return timeExit[:-1]

#������������� ������ � ����������� ����
def write_setting(date, num_setting):
    #������ ���� ���������
    f = open('work_setting\work_setting.txt', 'r')
    lines = f.readlines()
    f.close()
    #�������� �������� ������ �� �����
    #��������� ��� � ����� ���� ����������� ������
    if len(lines)-1 >= num_setting:
        lines[num_setting] = str(date) + '\n'
    #��������� ����� ���������
    else:
        lines.append(str(date))
    #��������� ���� ������ ����� � ����
    save_f = open('work_setting\work_setting.txt', 'w')
    save_f.writelines(lines)
    save_f.close()
    
#��������� �� ����� �������� ������
def read_setting(num_setting):
    #�������� ����������
    date = 0 #�� ���������
    
    #������ ���� ��������� � ���������� 22 ������
    f = open('work_setting\work_setting.txt', 'r')
    with f:
        lines = f.readlines()
        try:
            date = lines[num_setting]
        except:
            log_info("�� ��������� ������: %s" % num_setting)
    #��� ������� �������� ������, ���������� ��� srt
    return date[:-1]

#������ �� ����� ����������
def read_help():
    #������ ���� ���������
    f = open('work_setting\work_help.txt', 'r')
    text = f.read()
    f.close()
    
    return str(text)

#�������� ������ exel �����
def new_timework_file(path):
    f = open(path, 'w')
    f.close()

#���������� � ����
def save_setting(new_path, mode):
    
    old_WorkPath = read_path()              #������ ������ �������� ����
    old_WorkName = read_name()              #������ ������ �������� ����� �����
    WorkPath = os.path.dirname(new_path)    #���� ����� � ������� ����� ����
    WorkName = os.path.basename(new_path)   #��� �����
    
    #����������
    if(mode == 'Repace'):
        #���� ��������� ������� ���������� ���������� ���� ���� � ����� ������
        if os.path.exists(WorkPath):        
            #���������� ����, ���� ���� ����� ��� ��� ����������
            os.replace((old_WorkPath + '/' + old_WorkName),(WorkPath + '/' + WorkName))
            save_warning = 2
            #������ � ����������� ���� ������ ����� �����
            write_name(WorkName)
            #������ � ����������� ���� ������ ����
            write_path(WorkPath)
    elif (mode == 'New'):
        save_warning = 1
        #������� ����
        new_file = os.open((WorkPath + '/' + WorkName),os.O_CREAT)
        os.close(new_file)
    #���������� � ���������� ������ �������� �� ������������ ����� ���� � ��� �����
    #� ������, ���� ������� �������: 0-��� ������, 1-������������ ����������, 2-��� ����� �� ����������)
    return read_path(), read_name(), save_warning
        
        