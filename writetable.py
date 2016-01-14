# -*- coding: gb18030 -*-
'''
    Created on 2016-01-10

    @author: Gavin.Bai
    @note: Write table class
    @version: v1.0
    @Modify:
    @License: (C)GPL
'''
## ----------------------------------------------------------------------
## Import package 
import xlwt
import sys
import os
## ----------------------------------------------------------------------
## Define Exception
class OpenFileError(Exception):
    '''
    Deal some errors while open file failed
    '''
    promt_information = ""
    def __init__(self, promt_infor):
        super(OpenFileError, self).__init__()
        #print '[*] Error in <OpenFileError>: ', promt_infor
        self.promt_information = promt_infor

    def getErrorString(self):
        return self.promt_information
    
class SheetNameError(Exception):
    '''
    Deal some errors while acquire sheet name failed
    '''
    promt_information = ""
    def __init__(self, promt_info):
        super(SheetNameError, self).__init__()
        #print '[*] Error in <SheetNameError>: ', promt_info
        self.promt_information = promt_info

    def getErrorString(self):
        return self.promt_information

class WriteTable(object):
    '''
    The class type to write specific information to the sheet of file
    '''
    work_book  = None
    sheet_obj  = None
    sheet_name = ""
    file_name  = ""
    ready_flag = False
    
    def __init__(self, file_name, sheet_name = 'Sheet1'):
        '''
        Constructor for the class type
        @ file_name  : file to be written
        @ sheet_name : the sheet of the file to be filled
        '''
        super(WriteTable, self).__init__()
        self.work_book = xlwt.Workbook()
        if not file_name: 
            raise OpenFileError("<Constructor> : File name should not be empty")
        self.file_name  = file_name
        self.sheet_name = sheet_name
        self.sheet_obj  = self.work_book.add_sheet(self.sheet_name, cell_overwrite_ok = True)
        self.ready_flag = True
    
    def setVHeader(self, header):
        '''
        Set vertical header for the file to be written
        @ header : for the first column of the table
        '''
        ## 设置字体样式
        style       = xlwt.XFStyle()
        font        = xlwt.Font()
        font.name   = u'微软雅黑'
        font.bold   = True
        style.font  = font
        tittle_style = xlwt.easyxf('font: height 400, name Arial Black, colour_index blue, bold on; align: wrap on, vert centre, horiz center;'      "borders: top double, bottom double, left double, right double;")
        index = 0
        for item in header:
            self.sheet_obj.write(index, 0, item, tittle_style)
        
    def setHHeader(self, header):
        '''
        Set horizontal header for the file to be written
        @ header : for the first row of the table
        '''
        ## 设置字体样式
        style       = xlwt.XFStyle()
        font        = xlwt.Font()
        font.bold   = True
        font.name   = u'微软雅黑'
        font.height = 400
        ## 设置居中
        align = xlwt.Alignment()
        align.horz = xlwt.Alignment.HORZ_CENTER
        align.vert = xlwt.Alignment.VERT_CENTER
        
        style.alignment = align
        style.font  = font
        style.borders = self.setBorders()
        index = 0
        for item in header:
            self.sheet_obj.write(0, index, item, style)
            self.sheet_obj.col(index).width = 256 * 20
            index += 1
        self.sheet_obj.col(index-1).width = 256 * 65
            
    @staticmethod
    def setBorders():
        border = xlwt.Borders()
        border.left   = 1
        border.right  = 1
        border.top    = 1
        border.bottom = 1
        return border
    
    def isTableReady(self):
        '''
        Check whether the table which is used to be written has been ready
        @ True will be returned if it is ready, False will be returned
        '''
        return self.ready_flag
    
    def writeToFile(self):
        '''
        When write all information to the file, save operation is needed to be 
        done
        '''
        self.work_book.save(self.file_name)
        
    def setTableCellHeight(self, row_index, height):
        self.sheet_obj.row(0).heigh_mismatch = True
        self.sheet_obj.row(0).width = height
        
    def setValue(self, row, col, value):
        '''
        Write value to the specific row and col
        @ row   : the specific row to be written
        @ col   : the specific column to be written
        @ value : the value to be written
        '''
        self.sheet_obj.write(row, col, value)
        
    def setValueWithStyle(self, row, col, value, style):
        '''
        Write value to the specific row and col with style
        @ row   : the specific row to be written
        @ col   : the specific column to be written
        @ value : the value to be written
        @ style : the style to set table
        '''
        self.sheet_obj.write(row, col, value, style)
        
    def addSheetByName(self, sheet_name):
        '''
        Add sheet to the existed file
        @ sheet_name : the specific name of sheet to add to the file
        '''
        if not sheet_name:
            raise SheetNameError("<addSheetByName> : Sheet name should not be empty")
        
        if self.work_book:
            new_sheet_obj = self.work_book.add_sheet(sheet_name, cell_overwrite_ok = True)
            return new_sheet_obj
        else:
            return None
        
def main():
    write_file = WriteTable("test.xls", "Test")
    header = [u'姓名', u'年龄', u'住址', u'性别', u'毕业学校']
    write_file.setHHeader(header)
    write_file.setValue(1, 0, u'白亮')
    write_file.setValue(1, 1, 28)
    write_file.setValue(1, 2, u'上海')
    write_file.setValue(1, 3, u'男')
    write_file.setValue(1, 4, u'西安交通大学')
    write_file.writeToFile()

if __name__ == "__main__":
    main()