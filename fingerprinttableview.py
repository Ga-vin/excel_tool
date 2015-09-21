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

## -----------------------------------------------------------------------------
## Class Definition
class FingerprintTableView(tableview.TableView):
    '''
    From the fingerprint table, records of somebody can be check and get
    '''
    ## �鵽����Ա��Ϣ��
    person_lists = []
    person_name_sets = set()
    
    def __init__(self, file_name, sheet_name):
        super(FingerprintTableView, self).__init__(file_name, sheet_name)
    
    def mapNameListToIndexList(self, choice_list):
        '''
        Get the sheet title first, then according to the choice_list, get 
        index list of choice_list
        '''
        title = self.getHorizonTitle()
        temp_list = []
        for item in choice_list:
            try:
                temp_list.append(title.index(item))
            except ValueError, e:
                print '%s in <mapNameListToIndexList>' % e
            except ItemNotExistError, e:
                print e.getErrorString()
        return temp_list
    
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
        empty_dict['record'] = {record_list[1] : record_list[2]}
        self.addPersonToList(empty_dict)
    
    def addDateTimeRecord(self, name, date, ti):
        '''
        Add only date & time to the person list
        '''
        index = self.getPersonIndexByName(name)
        self.person_lists[index]['record'][date] = ti
    
    def updatePersonInformatioin(self, start_row, end_row, choice_list):
        '''
        According to the fingerprint table's record update the 
        person list information
        '''
        if start_row > end_row:
            raise UpdateIndexError("start index should be less than end index")
        if start_row < 0 or end_row > self.getRowNumbers():
            raise UpdateIndexError("end index should be less than total rows")
        
        start, end = start_row, end_row
        
        ## ��֤�ڸ�����Ա��Ϣ�б�֮ǰ�Ѿ��򿪸�excel�ļ�
        if self.isReady():
            while True:
                ## �ж��Ƿ�Ҫ�������²���
                if start > end:
                    break
                
                ## ��ȡһ�м�¼��Ϣ
                temp_list = self.getNextRowRecord(choice_list)
                
                ## �������������Ա��Ϣ���У����������¼
                if not self.isNameExist(temp_list[0]):
                    self.addTotalRecord(temp_list)
                else:
                    ## ������ڲ�����Ա��Ϣ���У�������ں�ʱ��
                    if not self.isDateExist(name = temp_list[0], da = temp_list[1]):
                        self.addDateTimeRecord(name = temp_list[0], date = temp_list[1], ti = temp_list[2])
                
                start += 1
        
## -----------------------------------------------------------------------------
## Test Driver
def main():
    excel_file_obj = FingerprintTableView("attendance.xls", u"ԭʼ1")
    print 'The sheet has %d row, %d col' % (excel_file_obj.getRowNumbers(), excel_file_obj.getColNumbers())
    print 'The horizontal title is ', 
    choice_lists = [u'Ա������', u'ǩ������', u'ǩ��ʱ��']
    
    excel_file_obj.updatePersonInformatioin(1, 10, choice_lists)
    person_list = excel_file_obj.getPersonList()
    for item in person_list:
        print item['name']
        for value in item['record'].keys():
            print value, ' : ', item['record'][value]
        print '=' * 40
    #===========================================================================
    # for item in excel_file_obj.getHorizonTitle():
    #     print item,'\t',
    # a = excel_file_obj.getNextRowRecord([u'Ա������', u'ǩ������', u'ǩ��ʱ��'])
    # b = excel_file_obj.getNextRowRecord([u'Ա������', u'ǩ������', u'ǩ��ʱ��'])
    # c = excel_file_obj.getNextRowRecord([u'Ա������', u'ǩ������', u'ǩ��ʱ��'])
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
