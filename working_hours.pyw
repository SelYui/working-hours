# -*- coding: cp1251 -*-
'''
�������� ������ ���������� ��������
'''

from work_setting import dialog, adjacent_classes

if __name__ == '__main__':
    
    #��������� � ��������� ������ ����� ShowWeb
    thread = adjacent_classes.ShowShutOrWeb()
    thread.start()
    
    #��������� ���� ���������� �� dialog
    dialog.app_main()
