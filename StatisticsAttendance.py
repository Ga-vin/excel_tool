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
tool_name = '考勤统计工具 '
tool_version = 'V1.0'
tool_author  = 'Gavin.Bai'
tool_copyright = '版本所有(C) 违者必究'

## Exception Definition
class ItemNotExistError(Exception):
    promt_information = ""
    
    def __init__(self, info):
        super(ItemNotExistError, self).__init__()
        self.promt_information = info
    
    def getErrorString(self):
        return "<ItemNotExistError> : %s" % self.promt_information
    
class UpdateIndexError(Exception):
    promt_information = ""
    
    def __init__(self, info):
        super(UpdateIndexError, self).__init__()
        self.promt_information = info
        
    def getErrorString(self):
        return "<UpdateIndexError> : %s" % self.promt_information

class RecordDateError(Exception):
    promt_information = ""
    
    def __init__(self, info):
        super(RecordDateError, self).__init__()
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
                          序号          id        : 0
                         姓名           name      : None
                       入职时间    hire_date : None
                       年               year      : 0
                       月              month     : 0                
                       日期             date      :
                       {
                           1日    date_1   : None
                           2日    date_2   : None
                           3日    date_3   : None
                           ........
                           31日    date_31 : None
                       }
                       迟到次数            late_times   : 0
                       调休次数            off_time     : 0
                       事假次数            leave_time   : 0
                       病假次数            sick_time    : 0
                       上月调休结余 last_rest_off : 0
                       现调休结余    now_rest_off  : 0
                       备注                 common        : 0
     }
          
    '''
    
    def __init__(self, file_name, sheet_name):
        '''
        Constructor function for the class
        @file_name  : the file to be opened to read
        @sheet_name : the sheet which is in the file to be read information
        '''
        super(StatisticData, self).__init__(file_name, sheet_name)
        self.person_list = []
        self.person_name_sets = set()
        self.year = 0
        self.month = 0
        self.person_total = 0
        
    def addPersonNameToSets(self, person_name):
        '''
        Add specific name of person to sets, if the person name has been existed
        in the sets, False will be returned, else the person name will be added
        to the set, and True will be returned
        person_name : the name to be added to set
        '''
        ## assert (not person_name), "<addPersonNameToSets> : person_name should not be empty"
        if not person_name:
            print 'Person name is empty'
            return False
#         else:
#             print 'Person name is not empty'
#             print person_name
        
        if person_name in self.person_name_sets:
            return False
        else:
            self.person_name_sets.add(person_name)
            self.person_total += 1
            return True
        
    def getPersonNameSets(self):
        '''
        Return the set which contains the name sets of person
        '''
        return self.person_name_sets
    
    def getPersonTotal(self):
        '''
        Return the total person has been statistic
        '''
        return self.person_total
        
    def generatePersonObj(self):
        '''
        Generate person structure which describe the basic information
        about the person
        @return the person object has been created and assigned with 
        initial value
        '''
        
        person = dict()
        person['id']                = 0
        person['name']              = None
        person['hire_date']         = None
        person['year']              = 0
        person['month']             = 0
        person['date']              = dict()
        for i in range(1, 32):
            temp_dict               = dict()
            temp_dict['late']       = 0
            temp_dict['register']   = ""
            temp_dict['sign_out']   = ""
            temp_dict['overwork']   = 0
            temp_dict['leav_early'] = 0
            temp_dict['off']        = 0
            temp_dict['leave']      = 0
            temp_dict['sick']       = 0
            temp_dict['annual']     = 0
            temp_dict['outside']    = None
            temp_dict['statistic']  = ""
            person['date']['date_' + str(i)] = temp_dict
        person['late_time']         = 0.0
        person['off_time']          = 0.0
        person['leave_time']        = 0.0
        person['sick_time']         = 0.0
        person['annual_time']       = 0.0
        person['last_rest_off']     = 0.0
        person['now_rest_off']      = 0.0
        person['common']            = ""
        
        return person
    
    def displayPersonObj(self, person_dict, offset = 0):
        '''
        Print the specific information about the person which has been 
        passed in to the screen to debug
        @person_dict : the dict structure which contains the information of 
                       the person
        @offset      : the offset which is mainly to print space before
                       each line
        '''
        
        if not person_dict:
            print 'Person Dict is Empty'
        else:
            for kk in person_dict.keys():
                if type(person_dict[kk]) == type(dict()):
                    print ' ' * offset, '%-16s'%kk, ":"
                    self.displayPersonObj(person_dict[kk], offset + 4)
                else:
                    print ' ' * offset, '%-16s'%kk, ":", person_dict[kk]
        

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
    print "\n\n\n"  

def printToolLogoEnd():
    '''
    Print the tail of logo for the software
    '''
    print ' ' * 15 + 'End Time: ' + time.asctime(time.localtime())
    print '*' * 60     

def main():
    printToolLogoHeader()
    ## Add data flow
    Table = StatisticData("record_total.xlsx", 'specific')
#     person = Table.generatePersonObj()
#     Table.displayPersonObj(person)

    printToolLogoEnd()

if __name__ == "__main__":
    main()