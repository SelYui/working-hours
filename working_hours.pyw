# -*- coding: cp1251 -*-
'''
�������� ������ ���������� ��������
'''

from work_setting import dialog, module
from work_lib import work_time

if __name__ == '__main__':

    #���� �� �������� ��������� ����� ����� �� �����-�� ������� - ���������� ��� �� ������������ �����
    err_exit = module.read_setting(28)
    print(err_exit)
    if int(err_exit) != 0:
        print('������')
        work_time.write_exit()   
    #���������� ������� ���������� ����������
    module.write_setting(1, 28)                    #�� ������ �����������������!!!
    
    #��������� ���� ���������� �� dialog
    dialog.app_main()   
    
    