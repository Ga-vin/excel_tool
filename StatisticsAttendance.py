# -*- coding: gb18030 -*-
'''
    Created on 2016-01-05

    @author: Gavin.Bai
    @note: Excel Tool To Statistics Attendance Data
    @version: v1.0
    @Modify:
    @License: (C)GPL
'''
## ----------------------------------------------------------------------
## Import package
import sys
import os
import time
import xlrd
import xlwt
import openpyxl
from tableview import TableView
## -----------------------------------------------------------------------------
## Constants for the whole project
tool_name = '����ͳ�ƹ��� '
tool_version = 'V1.0'
tool_author  = 'Gavin.Bai'
tool_copyright = '�汾����(C) Υ�߱ؾ�'

## Exception Definition
class ItemNotExistError(Exception):
    promt_information = ""
    
    def __init__(self, info):
        self.promt_information = info
    
    def getErrorString(self):
        return "<ItemNotExistError> : %s" % self.promt_information
    
class UpdateIndexError(Exception):
    promt_information = ""
    
    def __init__(self, info):
        self.promt_information = info
        
    def getErrorString(self):
        return "<UpdateIndexError> : %s" % self.promt_information

class RecordDateError(Exception):
    promt_information = ""
    
    def __init__(self, info):
        self.promt_information = info
    
    def getErrorString(self):
        return "<RecordDateError> : %s" % self.promt_information
## -----------------------------------------------------------------------------
## New Class definition
class StatisticData(TableView):
    '''
    New class which inherits from TableView, can use some features which are
    defined in base class, which also can extend its functions and features
    
    Person dict likes this:
      {
                          ���          id        : 0
                         ����           name      : None
                       ��ְʱ��    hire_date : None
                       ����             date      :
                       {
                           1��    date_1   : None
                           2��    date_2   : None
                           3��    date_3   : None
                           ........
                           31��    date_31 : None
                       }
                       �ٵ�����            late_times   : 0
                       ���ݴ���            off_time     : 0
                       �¼ٴ���            leave_time   : 0
                       ���ٴ���            sick_time    : 0
                       ���µ��ݽ��� last_rest_off : 0
                       �ֵ��ݽ���    now_rest_off  : 0
                       ��ע                 common        : 0
     }
          
    '''
    
    def __init__(self, file_name, sheet_name):
        '''
        Constructor function for the class
        '''
        super(StatisticData, self).__init__(file_name, sheet_name)
        self.person_list = []
        self.person_name_sets = set()

def printToolLogoHeader():
    '''
    Print the logo of the software
    '''
    print '*' * 60
    print ' ' * 75 + tool_name
    print ' ' * 23 + tool_version
    print ' ' * 21 + tool_author
    print ' ' * 70 + tool_copyright
    print ' ' * 15 + 'Start Time: ' + time.asctime(time.localtime())   

def printToolLogoEnd():
    '''
    Print the tail of logo for the software
    '''
    print ' ' *15 + 'End Time: ' + time.asctime(time.localtime())
    print '*' * 60     

def main():
    printToolLogoHeader()
    ## Add data flow
    pass

    printToolLogoEnd()

if __name__ == "__main__":
    main()