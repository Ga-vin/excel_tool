# -*- coding: gb18030 -*-
'''
    Created on 2015-09-18

    @author: Gavin.Bai
    @note: Excel Tool
    @version: v1.0
    @Modify:
    @License: (C)GPL
'''
## ----------------------------------------------------------------------
## Import package
import time
import sys
import openpyxl as pyxl
import xlrd
import xlwt

## ----------------------------------------------------------------------
## Define Exception
class OpenFileError(Exception):
    '''
    Deal some errors while open file failed
    '''
    promt_information = ""
    def __init__(self, promt_infor):
        print '[*] Error in <OpenFileError>: ', promt_infor
        self.promt_information = promt_infor

    def getErrorString(self):
        return self.promt_information

class GetHTitleError(Exception):
    '''
    Deal some errors while getting horizon title failed
    '''
    promt_informatioin = ""
    def __init__(self, promt_info):
        print '[*] Error in <GetHTitleError>: ', promt_info
        self.promt_information = promt_info

    def getErrorString(self):
        return self.promt_information

class GetVTitleError(Exception):
    '''
    Deal some errors while getting vertical title failed
    '''
    promt_information = ""
    def __init__(self, promt_info):
        print '[*] Error in <GetVTitleError>: ', promt_info
        self.promt_information = promt_info

    def getErrorString(self):
        return self.promt_information

class ItemNameError(Exception):
    '''
    Deal some errors while acquire item's names failed
    '''
    promt_information = ""
    def __init__(self, promt_info):
        print '[*] Error in <ItemNameError>: ', promt_info
        self.promt_information = promt_info

    def getErrorString(self):
        return self.promt_information

class ItemIndexError(Exception):
    '''
    Deal some errors while acquire item's index failed
    '''
    promt_information = ""
    def __init__(self, promt_info):
        print '[*] Error in <ItemIndexError>: ', promt_info

    def getErrorString(self):
        return self.promt_information
    
## -----------------------------------------------------------------------------
## Class Definition
class TableView(object):
    '''
    Class TableView is used to operate excel 2003/2007, even 2010 version.
    Open excel file and get information from the excel file.
    '''
    
    work_book  = None
    sheet      = None
    ready_flag = False
    file_name  = ""
    sheet_rows = 0
    sheet_cols = 0
    
    def __init__(self, file_name = "", sheet_name = ""):
        if not file_name:
            raise OpenFileError("Excel File is not specified!")
        ## 打开Excel文件
        try:
            self.work_book  = xlrd.open_workbook(file_name)
            self.file_name  = file_name
            self.ready_flag = True
        except (OpenFileError, IOError), e:
            print '[*] Open excel file <%s>  error' % file_name
            sys.exit() 
        ## 定位到对应的sheet，否则对应第一个sheet
        if self.ready_flag:
            if not sheet_name:
                self.sheet = self.work_book.sheet_by_index(0)
            else:
                self.sheet = self.work_book.sheet_by_name(sheet_name)
        ## 获取该sheet对应的总行数和总列数
        self.sheet_rows = self.sheet.nrows
        self.sheet_cols = self.sheet.ncols
        
    def getSheetByName(self, sheet_name):
        '''
        Change the sheet of the table which has been opened
        '''
        
        self.ready_flag = False
        if not self.work_book:
            raise OpenFileError("File does not open yet")
        
        if not sheet_name:
            print '[*] Sheet name should not be Empty'
            return False
        
        self.sheet = self.work_book.sheet_by_name(sheet_name)
        self.ready_flag = True
        self.sheet_rows = self.sheet.nrows
        self.sheet_cols = self.sheet.ncols
        
        return True
    
## -----------------------------------------------------------------------------
## Test Driver
def main():
    #===========================================================================
    # error = OpenFileError("excel_tool.py")
    # print error.getErrorString()
    # error = GetHTitleError("excel_tool.py")
    # print error.getErrorString()
    # error = GetVTitleError("excel_tool.py")
    # print error.getErrorString()
    # error = ItemNameError("excel_tool.py")
    # print error.getErrorString()
    # error = ItemIndexError("excel_tool.py")
    # print error.getErrorString()
    #===========================================================================
    excel_file = TableView('attendance.xlsx', u'原始1')
    
    
if __name__ == "__main__":
    main()
