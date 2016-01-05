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
import tableview
## -----------------------------------------------------------------------------
## Constants for the whole project
tool_name = '考勤统计工具 '
tool_version = 'V1.0'
tool_author  = 'Gavin.Bai'
tool_copyright = '版本所有(C) 违者必究'

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