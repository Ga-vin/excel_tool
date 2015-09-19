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
    current_row_index = 1
    current_col_index = 1
    
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
    
    def getSheetNames(self):
        '''
        Check the each sheet name by a list
        '''
        return self.work_book.sheet_names()
    
    def getRowNumbers(self):
        '''
        Get the numbers of rows of specific sheet
        '''
        return self.sheet_rows
    
    def getColNumbers(self):
        '''
        Get the number of cols of specific sheet
        '''
        return self.sheet_cols
    
    def isReady(self):
        '''
        Make sure the excel file has been opened, if it is, return True,
        else return False
        '''
        return self.ready_flag
    
    def getFileName(self):
        '''
        Return the name of opened excel file
        '''
        return self.file_name
    
    def getLineRow(self, num):
        '''
        Get the specific number row's value of the sheet
        If the row has nothing, return None
        '''
        if num > self.getRowNumbers():
            return None
        return self.sheet.row_values(num)
    
    def getLineCol(self, num):
        '''
        Get the specific number col's value of the sheet
        If the column has nothing, return None
        '''
        if num > self.getColNumbers():
            return None
        return self.sheet.col_values(num)
    
    def getCell(self, row, col):
        '''
        Get specific cell in the sheet of row & col numbers
        '''
        if row <= 0 and row > self.getRowNumbers():
            return None
        if col <= 0 and col > self.getColNumbers():
            return None
        return self.sheet.cell(row, col).value  
    
    def getHorizonTitle(self):
        '''
        Get the horizon title of the sheet
        '''
        return self.getLineRow(0)
    
    def getVerticalTitle(self):
        '''
        Get the vertical title of the sheet
        '''
        return self.getLineCol(0)
    
    def getNextLineRow(self):
        '''
        Get next line row of current sheet with increment by one
        '''
        if self.current_row_index > self.getRowNumbers():
            return None
        temp = self.getLineRow(self.current_row_index)
        self.current_row_index += 1
        return temp
    
    def getNextLineCol(self):
        '''
        Get next line column of current sheet with increment by one
        '''
        if self.current_col_index > self.getColNumbers():
            return None
        temp = self.getLineCol(self.current_col_index)
        self.current_col_index += 1
        return temp
    
    def setCurrentRowIndex(self, row):
        '''
        Set the current row index of the current sheet
        '''
        if row <= 0 and row > self.getRowNumbers():
            return None
        self.current_row_index = row
        return True
        
    def setCurrentColIndex(self, col):
        '''
        Set the current column index of the current sheet
        '''
        if col <= 0 and col > self.getColNumbers():
            return None
        self.current_col_index = col
        return True
    
    def resetCurrentRowIndex(self):
        '''
        Set the current row index to be zero, make it acquire data
        from first line row
        '''
        self.current_row_index = 1
        
    def resetCurrentColIndex(self):
        '''
        Set the current column index to be zero, make it acquire data
        from first line column
        '''
        self.current_col_index = 1

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
    excel_file = TableView('attendance.xlsx', u'原始2')
    for item in excel_file.getSheetNames():
        print item, 
    print '\nThe sheet 1 has %d rows' % excel_file.getRowNumbers()
    if u'原始1' in excel_file.getSheetNames():
        print 'yes'
    else:
        print 'no'
    print 'The sheet 1 has %d cols' % excel_file.getColNumbers()
    print 'The 1 line is ', excel_file.getLineRow(0)
    print 'The 2 line is ', excel_file.getLineRow(1)
    print 'The 3 line is ', excel_file.getLineRow(2)
    
    if excel_file.getSheetByName(u"原始1"):
        print '\nThe sheet 1 has %d rows' % excel_file.getRowNumbers()
        print 'The sheet 1 has %d cols' % excel_file.getColNumbers()
    
    print 'The (0, 0) cell is ', excel_file.getCell(0, 0)
    print 'The (0, 1) cell is ', excel_file.getCell(0, 1)
    print 'The (1, 0) cell is ', excel_file.getCell(1, 0)
    print 'The (1, 1) cell is ', excel_file.getCell(1, 1)
    excel_file.getHorizonTitle()
    
    
if __name__ == "__main__":
    main()
