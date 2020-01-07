# -*- coding: cp1251 -*-
'''
������ ��� �������� �������� ������� � ������ ����������� � ����
'''
# -*- coding: cp1251 -*-
import datetime
import xlrd, xlwt
from xlutils.copy import copy

from work_setting import module

#������ �������
month_word = ('�� ������','�� �������','�� ����','�� ������','�� ���','�� ����','�� ����','�� ������','�� ��������','�� �������','�� ������','�� �������')
    
#������� ��� ������ ��� ��������� �����
def start_work(tekminute, tekhour, tekday, tekmonth, tekyear):
    '''
    tekyear = tekdateandtime.year   #������� ���
    tekmonth = tekdateandtime.month #������� �����
    tekday = tekdateandtime.day     #������� �����
    tekhour = tekdateandtime.hour   #������� ���
    tekminute = tekdateandtime.minute    #������� ������
    '''
    #�������� ���� � ����� � ��������
    wt_filename = module.read_path() + '/' + module.read_name()
    min_offset = module.read_offset()
    #��������� ��������� �������
    flg_dontdata = 0    #�������� ������� ������������ ����
    
    #�������� �������� �� �����
    if tekminute - min_offset >= 0:
        tekminute = tekminute - min_offset
    else:
        tekhour = tekhour - 1
        tekminute = 60 + (tekminute - min_offset)
    
    #��������� ��� Exel ����
    read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    write_book = copy(read_book)
    #��������� �� ���� �������� ����
    try:
        #���� ���� � ������� ����� ��� ����������
        sheet_index = read_book.sheet_names().index(str(tekyear))
        #�������� �������� ���� � ����� �����
        sheet = read_book.sheet_by_index(sheet_index)
        #����� �� ���� ������
        sheet_nrows = sheet.nrows       #������ � Exel
        sheet_ncols = sheet.ncols       #������� � Exel
            
    except ValueError:
        #���� ����� �������� ���� ���, ������� ���� ����
        sheet = write_book.add_sheet(str(tekyear))
        sheet_index = read_book.nsheets
        #�.�. �������� ������
        sheet_nrows = 0
        sheet_ncols = 0

    #���� ���� �� ������, ����� ������ � �����
    if (sheet_nrows and sheet_ncols) != 0:
        #�������� ��������� ����
        i = sheet_nrows-1
        while i > 0:
            lastdate = sheet.row_values(i)[0]
            #���� ������ ������, �� ���� ���� ����
            if lastdate == '':
                i=i-1
            else:
                dd = lastdate[0:2]  #����
                mm = lastdate[3:5]  #�����
                break
        
        #���� ����� ������
        if tekmonth == int(mm):
            #��������� �� �� ��� ���� �� ���� ������
            if tekday-1 == int(dd):
                #��������� �������, ������
                i = sheet_nrows
            #�� ������� ���� ��������� �� ������ ���
            elif tekday == int(dd):
                #���������� ������� ���� ������� � ������ (j) ��������
                if time_compare(sheet_nrows, sheet, 1, tekhour, tekminute) == 0:    #���� ����� ���� ��� ���������, �� ������ �� ������
                    return
                #��������� ����� ����������, ���� ����� 30 ���, �� �� ���������� ����� �������, ������� �� ���������
                if pc_reload(sheet.row_values(sheet_nrows-1)[2], tekhour, tekminute) == 0:
                    return
                    
                #�� ��������� ���� (������� ������������)
                flg_dontdata = 5826
                i = sheet_nrows
            #�������� ����� ������
            else:
                i = sheet_nrows+1
        #������� ����� �����
        else:
            i = sheet_nrows+2
            #����� �����
            write_book.get_sheet(sheet_index).write(i,0,month_word[tekmonth-1])
            i=i+1
    #���� ���� ������, �������� ��������� � ������
    else:
        i=0
        #����� �����
        write_book.get_sheet(sheet_index).write(i,0,month_word[tekmonth-1])
        i=i+1
    
    #��������� ������ �����
    if flg_dontdata != 5826:
        if tekday < 10 and tekmonth < 10:
            write_book.get_sheet(sheet_index).write(i,0,'0'+str(tekday)+'.0'+str(tekmonth)+'.'+str(tekyear))
        elif tekday < 10 and tekmonth > 10:
            write_book.get_sheet(sheet_index).write(i,0,'0'+str(tekday)+'.'+str(tekmonth)+'.'+str(tekyear))
        elif tekday > 10 and tekmonth < 10:
            write_book.get_sheet(sheet_index).write(i,0,''+str(tekday)+'.0'+str(tekmonth)+'.'+str(tekyear))
        else:
            write_book.get_sheet(sheet_index).write(i,0,str(tekday)+'.'+str(tekmonth)+'.'+str(tekyear))    
    #��������� ������ ��������
    if tekhour < 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,1,'0'+str(tekhour)+':0'+str(tekminute))
    elif tekhour < 10 and tekminute > 10:
        write_book.get_sheet(sheet_index).write(i,1,'0'+str(tekhour)+':'+str(tekminute))
    elif tekhour > 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,1,str(tekhour)+':0'+str(tekminute))
    else:
        write_book.get_sheet(sheet_index).write(i,1,str(tekhour)+':'+str(tekminute))

    #���� ������ ����� ������ 15, � ���������� ������ ���� �����, �� ��������� ������������ ����� � ������
    if tekday > 15 and int(dd) <= 15:
        #��������� ����� � ���� ������
        i = sheet_nrows-1
        #���� ������������ ������ � �����
        while sheet.row_values(i)[0] != month_word[tekmonth-1]:
            i=i-1
        #��������� ��� ������ ����� � ������
        #��������� �������� ������������ �����
        mount_sum = 0
        while i <= sheet_nrows-1:
            #���� ������ ������
            if sheet.row_values(i)[3] == '':
                i=i+1
            #�������� ���� ������������ � ���
            else:
                mount_sum = mount_sum + float(sheet.row_values(i)[3])
                i=i+1
        #��������� �� 3��� �����
        mount_sum = round(mount_sum,3)
        #��������� ����� ����� � ��������������� ������
        write_book.get_sheet(sheet_index).write(i,4,'('+str(mount_sum)+')')

    #��������� ������
    try:
        write_book.save(wt_filename)
    except Exception as e:
        module.log_info("�� ������� ��������� � Exel ����� �������. Exception: %s" % str(e))

#������� ��� ������ ��� ���������� �����     
def exit_work(tekminute, tekhour, tekday, tekmonth, tekyear):
    '''
    tekyear = tekdateandtime.year   #������� ���
    tekmonth = tekdateandtime.month #������� �����
    tekday = tekdateandtime.day     #������� �����
    tekhour = tekdateandtime.hour   #������� ���
    tekminute = tekdateandtime.minute    #������� ������
    '''

    #�������� ���� � ����� � ��������
    wt_filename = module.read_path() + '/' + module.read_name()
    min_offset = module.read_offset()
    
    #�������� �������� �� �����
    if tekminute + min_offset < 60:
        tekminute = tekminute + min_offset
    else:
        tekhour = tekhour + 1
        tekminute = (tekminute + min_offset) - 60
    
    #��������� ��� Exel ����
    read_book = xlrd.open_workbook(str(wt_filename), formatting_info=True)
    write_book = copy(read_book)
    
    #��������� �� ���� �������� ����
    try:
        #���� ���� � ������� ����� ��� ����������
        sheet_index = read_book.sheet_names().index(str(tekyear))
    except:
        #���� ����� �������� ���� ��� ������ ��������� start �� ���������
        return
    
    #�������� �������� ���� � ����� �����
    sheet = read_book.sheet_by_index(sheet_index)
    
    #�������� ��������� ����
    i = sheet.nrows-1
    while i > 0:
        lastdate = sheet.row_values(i)[0]
        #���� ������ ������, �� ���� ���� ����
        if lastdate == '':
            i=i-1
        else:
            dd = lastdate[0:2]  #����
            mm = lastdate[3:5]  #�����
            #� ���� ������ ����� ������ �������� ���������� ����� � ���
            time_date_index = i
            break
    
    #���� ����� ������
    if tekmonth == int(mm):
        #��������� �� �� ��� ���� ����� ��
        if tekday == int(dd):
            #��������� �������, ������ 
            i = sheet.nrows-1
        #���� ���-�� �� �������, �� start �� �������� ����� �� ��������� ������ ������
        else: return
    else: return
    
    #���������� ������� ���� ������� �� ������ (j) ��������
    if time_compare(sheet.nrows, sheet, 2, tekhour, tekminute) == 0:    #���� ����� ���� ��� ���������, �� ������ �� ������
        return

    #��������� ������ ��������
    if tekhour < 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,2,'0'+str(tekhour)+':0'+str(tekminute))
    elif tekhour < 10 and tekminute > 10:
        write_book.get_sheet(sheet_index).write(i,2,'0'+str(tekhour)+':'+str(tekminute))
    elif tekhour > 10 and tekminute < 10:
        write_book.get_sheet(sheet_index).write(i,2,str(tekhour)+':0'+str(tekminute))
    else:
        write_book.get_sheet(sheet_index).write(i,2,str(tekhour)+':'+str(tekminute))
    
    #��������� ���������� ������� ����� � ����������� ���
    i = sheet.nrows-1
    sum_timework = 0
    count_cycle = 0
    #���� �� ����� �� ������ ������ ���
    while i >= time_date_index:
        #�������� ����� �������
        time_start = sheet.row_values(i)[1]
        #�������� ���� � ������ �������
        hr_start = time_start[0:2]  #����
        min_start = time_start[3:5]  #������
        timestart = int(hr_start) + int(min_start)/60
        #�������� ���� � ������ �����
        if (int(sheet.ncols) < 3) or (sheet.row_values(i)[2] == '') or (i == time_date_index):
            hr_exit = tekhour  #����
            min_exit = tekminute  #������
        else:
            time_exit = sheet.row_values(i)[2]
            hr_exit = time_exit[0:2]  #����
            min_exit = time_exit[3:5]  #������
        timeexit = int(hr_exit) + int(min_exit)/60
        #��������
        timework = timeexit - timestart
        #�������� ����. ���� ���� ������ � ������� ����� ������ 4, �������� ���
        if (i == time_date_index) and timework > 4 and count_cycle == 0:
            timework = timework -1
        #����� ������� ����� � ���
        sum_timework = sum_timework + timework
        count_cycle = count_cycle+1
        i = i-1

    #��������� �� 3��� �����
    sum_timework = round(sum_timework,3)
    #��������� ������ �����
    write_book.get_sheet(sheet_index).write(time_date_index,3,str(sum_timework))
    #��������� ���������� ������� ����� � ������� ������
    i = sheet.nrows-1
    #���� ������������ ������ � �����
    while sheet.row_values(i)[0] != month_word[tekmonth-1]:
        i=i-1
    #���������� ������, ���� ������� �����
    index_sum = i
    #��������� ��� ������ ����� � ������
    i = time_date_index-1
    #��������� �������� ������������ ����� - ������ ��� ����������� ��������
    mount_sum = sum_timework
    while i > int(index_sum):
        #���� ������ ������
        if sheet.row_values(i)[3] == '':
            i=i-1
        else:
            #�������� ���� ������������ � ���
            mount_sum = mount_sum + float(sheet.row_values(i)[3])
            i=i-1
    #��������� �� 3��� �����
    mount_sum = round(mount_sum,3)
    #��������� ����� ����� � ��������������� ������
    write_book.get_sheet(sheet_index).write(index_sum,1,str(mount_sum))
    
    #��������� ������
    try:
        write_book.save(wt_filename)
    except Exception as e:
        module.log_info("�� ������� ��������� � Exel ����� �����. Exception: %s" % str(e))

#���� ���� ���������� �� �� ������� �� ��������� ����� ������� (���������� ��������� ��� �� ���������)
def pc_reload(timeexit, starthour, startminute):
    reload = module.read_reload()
    #���� ������ ������, ������� �� ���������, ������ �� ������ ����
    if timeexit == '':
        module.log_info("pc reload = 2")
        return 2    #����������� ������ - ���������� ������
    else:
        hour_e = timeexit[0:2]      #����
        minut_e = timeexit[3:5]     #������
    #������� ����� ����� � �������
    te = int(hour_e) + int(minut_e)/60
    ts = int(starthour) + int(startminute)/60
    #���� �������� �������� ������� � ����� ����� ��������, �� �� ���������� ����� ������� �������
    if ts - te < (reload/60):
        return 0    #������� �� �� ����� - �������
    else: return 1  #������� �� ����� - ���������� ������

#���������� ����� ������� � �������� � Exel, ���� �������, �� �� ����� ����������
def time_compare(sheet_nrows, sheet, j, tekhour, tekminute):
    #�������� ��������� ����� �������
    i = sheet_nrows-1
    while i > 0:
        lastdate = sheet.row_values(i)[j]
        #���� ������ ������, �� ���� ���� ����
        if lastdate == '':
            i=i-1
        else:
            shour = lastdate[0:2]  #����
            sminute = lastdate[3:5]  #�����
            break
    #���� ����� ������� � ��������� ���������� �������� - ������ �� ������. ������� �� ���������
    if (int(shour) == tekhour) and (int(sminute) == tekminute):
        return 0    #������� ���� ��������� � ��������� ����������
    else: return 1  #�� ���������
        
#�������� ��� ������ �� ��������� �� ������
def quit_app():
    #�������� ������� ���� � ����� �����
    tekdateandtimeExit = datetime.datetime.now()
    '''
    tekyear = tekdateandtimeExit.year   #������� ���
    tekmonth = tekdateandtimeExit.month #������� �����
    tekday = tekdateandtimeExit.day     #������� �����
    tekhour = tekdateandtimeExit.hour   #������� ���
    tekminute = tekdateandtimeExit.minute    #������� ������
    '''
    #���� ���������� ����� ���������� ���������� � ����
    module.write_timeExit(tekdateandtimeExit.strftime("%d %m %Y %H:%M"))
    
    #���������� ����� ���������� ����������
    #exit_work(tekminute, tekhour, tekday, tekmonth, tekyear)

def write_exit():
    #���������� � Exel ���� ����� ���������� ���������� ����������
    dtimE = module.read_timeExit()
    dtimE = dtimE.split()
    #���� ���������� ������ �� ������������
    if (dtimE != ''):
        timE = dtimE[-1]
        exit_work(int(timE[3:5]), int(timE[0:2]), int(dtimE[0]), int(dtimE[1]), int(dtimE[2]))
