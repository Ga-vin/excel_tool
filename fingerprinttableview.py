# -*- coding: gb18030 -*-
'''
    Created on 2015-09-19

    @author: Gavin.Bai
    @note: Excel Tool
    @version: v1.0
    @Modify:
    @License: (C)GPL
'''

## ----------------------------------------------------------------------
## Import package
import sys
import os
import xlrd
import xlwt
import openpyxl
import tableview

## -----------------------------------------------------------------------------
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
## Class Definition
class FingerprintTableView(tableview.TableView):
    '''
    From the fingerprint table, records of somebody can be check and get
    '''
    ## 查到的人员信息表
    person_lists = []
    person_name_sets = set()
    record_month = 1
    record_year  = 2015
    record_day   = 1
    
    def __init__(self, file_name, sheet_name, month = 1, year = 2015):
        super(FingerprintTableView, self).__init__(file_name, sheet_name)
        self.record_month = month
        self.record_year = year
    
    def mapNameListToIndexList(self, choice_list):
        '''
        Get the sheet title first, then according to the choice_list, get 
        index list of choice_list
        '''
        title = self.getHorizonTitle(0)
        temp_list = []
        for item in choice_list:
            try:
                temp_list.append(title.index(item))
            except ValueError, e:
                print '%s in <mapNameListToIndexList>' % e
            except ItemNotExistError, e:
                print e.getErrorString()
        return temp_list
    
    def setCurrentMonth(self, month):
        '''
        Set current record month value which is between 1 & 12
        '''
        if month > 0  and month < 13:
            self.record_month = month
        else:
            raise RecordDateError("Set current month which should be between 1 and 12")
        
    def getCurrentMonth(self):
        '''
        Return the record month
        '''
        return self.record_month
        
    def setCurrentYear(self, year):
        '''
        Set current record year value which should be than 1970
        '''
        if year > 1970:
            self.record_year = year
        else:
            raise RecordDateError("Set current year which should be than 1970")
    
    def getCurrentYear(self):
        '''
        Return the record year
        '''
        return self.record_year
    
    def formatDate(self, day):
        '''
        Format a date string with unicode style of xxxx-xx-xx
        '''
        return u'%d-%02d-%02d' % (self.getCurrentYear(), self.getCurrentMonth(), day)
    
    def getNextRowRecord(self, choice_list):
        '''
        According to the items in choice_list, check the next whole line, and return 
        specific columns in the line
        '''
        choice_index = self.mapNameListToIndexList(choice_list)
        #=======================================================================
        # print choice_index
        #=======================================================================
        temp = self.getNextLineRow()
        next_row_list = []
        for index in choice_index:
            next_row_list.append(temp[index])
        return next_row_list
    
    def getPersonIndexByName(self, name):
        '''
        Get the index in the person list according to the name
        '''
        index = 0
        for person in self.person_lists:
            if name == person['name']:
                break
            index += 1
        if index == len(self.person_lists):
            return -1
        else:
            return index
        
    def getPersonByIndex(self, index):
        '''
        Get the index's element 
        '''
        if index < 0 or index > len(self.person_lists):
            return None
        return self.person_lists[index]
    
    def getPersonList(self):
        '''
        Return the whole person list to user
        '''
        return self.person_lists
    
    def addPersonToList(self, new_dict):
        '''
        Add a new person record dict to the person list
        '''
        self.person_lists.append(new_dict)
    
    def isNameExist(self, name):
        '''
        Check whether the name exists in the person list
        '''
        temp = set()
        temp.add(name)
        return temp.issubset(self.person_name_sets)
    
    def isDateExist(self, name, da):
        '''
        Check whether the date exists in the person list
        '''
        item = self.getPersonByIndex(self.getPersonIndexByName(name))
        return item['record'].has_key(da)
    
    def addTotalRecord(self, record_list):
        '''
        Add all elements in the record_list to the person list
        '''
        empty_dict = {}
        empty_dict['name'] = record_list[0]
        empty_dict['record'] = {record_list[1] : None}
        empty_dict['record'][record_list[1]] = {'time' : record_list[2], 'absent' : None}
        self.addPersonToList(empty_dict)
        self.person_name_sets.add(record_list[0])
    
    def addDateTimeRecord(self, name, date, ti):
        '''
        Add only date & time to the person list
        '''
        index = self.getPersonIndexByName(name)
        #=======================================================================
        # self.person_lists[index]['record'] = {date : None}
        #=======================================================================
        self.person_lists[index]['record'][date] = {'time' : ti, 'absent' : None}
    
    def updatePersonFingerprintRecord(self, start_row, end_row, choice_list):
        '''
        According to the fingerprint table's record update the 
        person list information
        '''
        if start_row > end_row:
            raise UpdateIndexError("start index should be less than end index")
        if start_row < 0 or end_row > self.getRowNumbers():
            raise UpdateIndexError("end index should be less than total rows")
        
        start, end = start_row, end_row
        
        ## 保证在更新人员信息列表之前已经打开该excel文件
        if self.isReady():
            while True:
                ## 判断是否要结束更新操作
                if start > end:
                    break
                
                ## 获取一行记录信息
                temp_list = self.getNextRowRecord(choice_list)
                
                ## 如果姓名不在人员信息表中，添加整条记录
                if not self.isNameExist(temp_list[0]):
                    self.addTotalRecord(temp_list)
                else:
                    ## 如果日期不在人员信息表中，添加日期和时间
                    if not self.isDateExist(name = temp_list[0], da = temp_list[1]):
                        self.addDateTimeRecord(name = temp_list[0], date = temp_list[1], ti = temp_list[2])
                
                start += 1
                
    def updatePersonAbsentRecord(self, start_row, end_row, start_col, end_col):
        '''
        Update absent for each body of each day in this monty
        '''
        ## 索引异常处理
        if (start_row < 0) or (end_row < 0) or (start_row > end_row):
            raise UpdateIndexError("<updatePersonAbsentRecord> start row/end row should be right")
        if end_row > self.getRowNumbers():
            raise UpdateIndexError("<updatePersonAbsentRecord> end row should be less than total rows of the sheet")
        if (start_col < 0) or (end_col < 0) or (start_col > end_col):
            raise UpdateIndexError("<updatePersonAbsentRecord> start col/end col should be right")
        if end_col > self.getColNumbers():
            raise UpdateIndexError("<updatePersonAbsentRecord> end column should be less than total columns of the sheet")
        
        start, end = start_row, end_row
        while True:
            ## 判断读取信息结束
            if start > end:
                break
            
            ## 获取一行记录信息
            absent_record_list = self.getNextLineRow()
            if not self.isNameExist(absent_record_list[0]):
                print '<updatePersonAbsentRecord> %s does not exist' % absent_record_list[0]
            else:
                person_index = self.getPersonIndexByName(absent_record_list[0])
            
            start += 1
        
## -----------------------------------------------------------------------------
## Test Driver
def main():
    excel_file_obj = FingerprintTableView("attendance.xls", u"原始1")
    print 'The sheet has %d row, %d col' % (excel_file_obj.getRowNumbers(), excel_file_obj.getColNumbers())
    title = excel_file_obj.getHorizonTitle(1)
    print 'The title of the sheet 2 is'
    for item in excel_file_obj.getHorizonTitle(1):
        print item, 
    print

    ## 读取打卡记录
    choice_lists = [u'员工姓名', u'签到日期', u'签到时间']
     
    excel_file_obj.updatePersonFingerprintRecord(1, 100, choice_lists)
    person_list = excel_file_obj.getPersonList()
    for item in person_list:
        print item['name']
        for value in sorted(item['record'].keys()):
            print value, ' : ', item['record'][value]['time'], " : ", item['record'][value]['absent']
        print '=' * 40

    ## 读取请假信息
    excel_file_obj.resetCurrentRowIndex()
    if excel_file_obj.getSheetByName(u'原始2'):
        print 'Change to sheet2 Successfully'
        
    #===========================================================================
    # a = excel_file_obj.getNextRowRecord([u'员工姓名', u'签到日期', u'签到时间'])
    # b = excel_file_obj.getNextRowRecord([u'员工姓名', u'签到日期', u'签到时间'])
    # c = excel_file_obj.getNextRowRecord([u'员工姓名', u'签到日期', u'签到时间'])
    #  
    # for item in a:
    #     print item, '<', item[5:7], '>', '\t',   
    # print 
    # for item in b:
    #     print item, '\t',
    # print
    # for item in c:
    #     print item, '\t',
    # print
    #===========================================================================

if __name__ == "__main__":
    main()
