# -*- coding: gb18030 -*-
'''
    Created on 2016-01-08

    @author: Gavin.Bai
    @note: Main Entry Function For Excel Tool To Statistics Attendance Data
    @version: v1.0
    @Modify:
    @License: (C)GPL
'''
## ----------------------------------------------------------------------
from datetype import Date
import StatisticsAttendance
from StatisticsAttendance import StatisticData
import xlrd
import sys

class TableObjectEmptyError(Exception):
    '''
    Exception class which is to deal with when table object that passed in 
    the function is None
    '''
    promt_information = "<TableObjectEmptyError> : "
    def __init__(self, infor):
        super(TableObjectEmptyError, self).__init__()
        self.promt_information += infor
    def getErrorString(self):
        return self.promt_information
    
class TableHeaderEmptyError(Exception):
    '''
    Exception class which is to deal with when the parameter table header is 
    empty
    '''
    promt_information = "<TableHeaderEmptyError> : "
    def __init__(self, infor):
        super(TableHeaderEmptyError, self).__init__()
        self.promt_information = infor
    def getErrorString(self):
        return self.promt_information
    
class TableDataEndError(Exception):
    '''
    Exception class which is to deal with when reading from table to the end
    '''
    promt_information = "<TableDataEndError> : "
    def __init__(self, infor):
        super(TableDataEndError, self).__init__()
        self.promt_information += infor
    def getErrorString(self):
        return self.promt_information

class TableHeaderLengthError(Exception):
    '''
    Exception class which is to deal with when the length of table header is not
    same with one line which reads from the table
    '''
    promt_information = "<TableHeaderLengthError> : "
    def __init__(self, infor):
        super(TableHeaderLengthError, self).__init__()
        self.promt_information += infor
    def getErrorString(self):
        return self.promt_information

def getSpecificLineData(obj, table_header):
    '''
    Fill each item to the person information according to the table horizontal
    header
    @ Return a dictionary
    '''
    ## 检查传入参数是否为空
    if not obj:
        raise TableObjectEmptyError("<getSpecificLineData> : Table object is None") 
    
    ## 检查传入的表格标题头是否为空
    if not table_header:
        raise TableHeaderEmptyError("<getSpecificLineData> : Table header is None")
    
    temp_obj = dict()
    data = obj.getNextLineRow()
    if not object:
        ## 读到表格的最后一行
        raise TableDataEndError("<getSpecificLineData> : End of Table object")
    elif len(table_header) != len(data):
        ## 标题长度和数据长度不一致
        raise TableHeaderLengthError("<getSpecificLineData> : Header length is same")
    else:
        for key, item in zip(table_header, data):
            if not item:
                item = 0
            temp_obj[key] = item
    
    return temp_obj

def addPersonName(table_obj, person_id, person_dict):
    '''
    Add the name of person to name sets
    @ table_obj   : Table object has been opened
    @ person_id   : the counts of total person have been added to the set
    @ person_dict : Get person name from the person_dict
    '''
    if not table_obj:
        print "[*] Table object is None"
        return None
    
    if not person_dict:
        print "[*] Person dictionary is None"
        return None
    
    if table_obj.addPersonNameToSets(person_dict[u"姓名"]):
        person_id += 1
        person_obj = table_obj.generatePersonObj()
        if not person_obj:
            print "[*] Create person object failed"
            return None
        person_obj["name"] = person_dict[u"姓名"]
        person_obj["id"]   = person_id
    else:
        person_obj = None
    
    return (person_id, person_obj)

def addPersonDate(date_string, split_char = '/'):
    '''
    Split the input parameter with specific character, and convert to a Date 
    class object
    @ If some error happens, None will be returned
    '''
    if not date_string:
        print '[*] Date string is empty'
        return None
    
    try:
        raw_string = date_string.split(split_char)
        date_obj = Date(int(raw_string[0]), int(raw_string[1]), int(raw_string[2]))
    except ValueError, e:
        print '[*] ERROR: ', e
        return None
    return date_obj

def addPersonLateMinutes(sign_time, fix_sign_time):
    '''
    Calculate the minutes for the person has been late
    @ sign_time     : his first sign time
    @ fix_sign_time : fix last sign time
    '''
    tuple_list          = ['year', 'month', 'day', 'hour', 'minute', 'second']
    ## 获得签到时间元组
    sign_time_tuple     = xlrd.xldate_as_tuple(sign_time, 0)
    ## 获得规定签到时间元组
    fix_sign_time_tuple = xlrd.xldate_as_tuple(fix_sign_time, 0)
    
    sign_minutes = sign_time_tuple[tuple_list.index('hour')]*60 \
        + sign_time_tuple[tuple_list.index('minute')]
    fix_sign_minutes = fix_sign_time_tuple[tuple_list.index('hour')]*60 \
        + fix_sign_time_tuple[tuple_list.index('minute')]
    total_minutes = sign_minutes - fix_sign_minutes
    if total_minutes >= 0:
        return total_minutes
    else:
        return 0
    
def addPersonLeaveEarlyMinutes(sign_time, fix_sign_time):
    '''
    Calculate the minutes for the person has been leaved early
    @ sign_time     : his leave time
    @ fix_sign_time : fix leave time   
    '''
    tuple_list          = ['year', 'month', 'day', 'hour', 'minute', 'second']
    ## 获得签退时间
    sign_time_tuple     = xlrd.xldate_as_tuple(sign_time, 0)
    ## 获得规定签退时间
    fix_sign_time_tuple = xlrd.xldate_as_tuple(fix_sign_time, 0) 
    
    sign_minutes = sign_time_tuple[tuple_list.index('hour')]*60\
        + sign_time_tuple[tuple_list.index('minute')]
    fix_sign_minutes = fix_sign_time_tuple[tuple_list.index('hour')]*60\
        + fix_sign_time_tuple[tuple_list.index('minute')]
    total_minutes = sign_minutes - fix_sign_minutes
    return total_minutes

def addPersonOvertimeNotWorkday(come_time, leave_time):
    '''
    Calculate the minutes for the person has been on overwork in unworkday
    @ come_time  : his coming time
    @ leave_time : his leave time
    '''
    tuple_list = ['year', 'month', 'day', 'hour', 'minute', 'second']
    ## 获得签到时间
    come_time_tuple  = xlrd.xldate_as_tuple(come_time, 0)
    ## 获得离开时间
    leave_time_tuple = xlrd.xldate_as_tuple(leave_time, 0)
    come_minutes  = come_time_tuple[tuple_list.index('hour')]*60 + come_time_tuple[tuple_list.index('minute')]
    leave_minutes = leave_time_tuple[tuple_list.index('hour')]*60 + leave_time_tuple[tuple_list.index('minute')]
    total_minutes = leave_minutes - come_minutes
    return total_minutes

## 主入口程序
def main():
    read_table_name       = "record_total.xlsx"
    read_table_sheet_name = "specific"
    person_total_list     = dict()
    
    ## 便于区分添加的编号
    person_id = 0
    start_from_row = 1
    
    ## 打印标题
    StatisticsAttendance.printToolLogoHeader()
    
    ## 打开读取表格
    read_table = StatisticData(read_table_name, read_table_sheet_name)
    ## 提取表格头
    read_table_header = read_table.getHorizonTitle()
    if not read_table_header:
        print '<main> : Data Table Header is Empty'
        ## 如何为空表，则结束
        sys.exit()
        
    ## 设置读取数据开始行，便于直接读取下一行
    read_table.setCurrentRowIndex(start_from_row)
    ## 开始读取统计数据
    last_person_id = 1
    temp = 0
    while True:
        try:
            is_new_person = False
            ## 读取一行的数据并转化为一个字典
            person_dict = getSpecificLineData(read_table, read_table_header)
            
            if not person_dict[u'姓名']:
                print '[*] End of the table'
                break
            
            ## 添加姓名到集合中
            person_id, person_obj_temp = addPersonName(read_table, person_id, person_dict)
            if person_obj_temp:
                ## 添加到人员列表中
                is_new_person = True
                person_total_list[person_id] = person_obj_temp
            
            date_string = person_dict[u'日期']
            date_temp   = addPersonDate(date_string)
            if not date_temp:
                print "[*] Date convert fails. System will be terminated"
                sys.exit()
            ## 如果是新人，则添加年和月
            if is_new_person:
                is_neww_person = False
                person_total_list[person_id]['year']  = date_temp.getYear()
                person_total_list[person_id]['month'] = date_temp.getMonth()
            
            ## 获取具体的工作日
            work_day = date_temp.getDay()
            if not date_temp.isWeekDay():
                ## 非工作日，算加班
                overtime = addPersonOvertimeNotWorkday(person_dict[u'签到时间'], person_dict[u'签退时间'])
                person_total_list[person_id]['date']['date_'+str(work_day)]['late'] = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['leav_early'] = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['off']        = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['leave']      = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['sick']       = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['annual']     = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['outside']    = overtime
            else:
                late_minutes = addPersonLateMinutes(person_dict[u'签到时间'], person_dict[u'规定上班时间'])
                early_overtime_minutes = addPersonLeaveEarlyMinutes(person_dict[u'签退时间'], person_dict[u'规定下班时间'])
                ## 迟到时间
                person_total_list[person_id]['date']['date_'+str(work_day)]['late'] = late_minutes
                
                if early_overtime_minutes > 0:
                    ## 加班的情况
                    person_total_list[person_id]['date']['date_'+str(work_day)]['overwork']   = early_overtime_minutes
                    person_total_list[person_id]['date']['date_'+str(work_day)]['leav_early'] = 0
                else:
                    ## 早退的情况
                    person_total_list[person_id]['date']['date_'+str(work_day)]['overwork'] = 0
                    person_total_list[person_id]['date']['date_'+str(work_day)]['leav_early'] = (early_overtime_minutes * (-1))
                
                ## 事假时间
                person_total_list[person_id]['date']['date_'+str(work_day)]['leave'] = int(person_dict[u'事假时间'])
                ## 病假时间
                person_total_list[person_id]['date']['date_'+str(work_day)]['sick'] = int(person_dict[u'病假时间'])
                ## 调休/年假时间
                absent_hour = float(person_dict[u'带薪假时间   （调休或其他）'])
                absent_reason = person_dict[u'带薪假别']
                if u'调休' == absent_reason:
                    person_total_list[person_id]['date']['date_'+str(work_day)]['off']    = absent_hour
                    person_total_list[person_id]['date']['date_'+str(work_day)]['annual'] = 0
                elif u'年假' == absent_reason:
                    person_total_list[person_id]['date']['date_'+str(work_day)]['off']    = 0
                    person_total_list[person_id]['date']['date_'+str(work_day)]['annual'] = absent_reason
                ## 外勤情况
                person_total_list[person_id]['date']['date_'+str(work_day)]['outside'] = person_dict[u'外勤及其他异常说明']    
                
            if person_id > last_person_id:
                last_person_id = person_id
        except (TableObjectEmptyError, TableHeaderEmptyError,\
                TableDataEndError, TableHeaderLengthError) as e:
            print e.getErrorString()
        ## For debug
#         print work_day, '~'*60
#         read_table.displayPersonObj(person_total_list[person_id])
#         temp += 1
#         if temp > 22:
#             break
    print u'共有%d人'%person_id
    ## 打印工具结尾
    StatisticsAttendance.printToolLogoEnd()

if __name__ == "__main__":
    main()