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
import xlwt
import sys
from writetable import WriteTable
from writetable import OpenFileError

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
    
    ## 无签到时间
    if 6 == sign_time_tuple.count(0):
        return -1
    
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
    
    ## 无签退时间
    if 6 == sign_time_tuple.count(0):
        return -1
    
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

def doPersonStatisticsData(person_list):
    '''
    Statistics all kinds of information about the person, such as off, late,
    absent, sick-absent, and so on. Fill the statistics information to the 
    specific date field.
    At last, the person_list which has been upgrated will be returned
    @ person_list : contains all person information which is dictionary
    '''
    ## 复制一份拷贝，以防修改先前的对象
    import copy
    person_list_cp = copy.deepcopy(person_list)
    length_person_list = len(person_list_cp.keys())
    for key in range(1, length_person_list+1):
        person = person_list_cp[key]
        ## 针对每一个人统计
        work_month = person['month']
        work_year  = person['year']
        date       = Date(work_year, work_month, 1)
        ## 获取当前月有多少天
        days_has   = date.hasDays()
        person_off_hours    = 0.0
        person_late_minutes = {'below_10':0, 'between_10_30':0, 'up_30':""}
        person_leav_early_minutes = 0.0
        person_leave_hours  = 0.0
        person_sick_hours   = 0.0
        person_annual_hours = 0.0
        for day in range(1, days_has + 1):
            ## 调休时间统计
            off = person['date']['date_'+str(day)]['off']
            if  float(off) != 0.0:
                person_off_hours += float(off)
                person['date']['date_'+str(day)]['statistic'] += u'调'+ str(float(off))
            ## 事假时间统计
            leave = person['date']['date_'+str(day)]['leave']
            if float(leave) != 0.0:
                person_leave_hours += float(leave)
                person['date']['date_'+str(day)]['statistic'] += u'事' + str(float(leave))
            ## 年假时间统计
            annual = person['date']['date_'+str(day)]['annual']
            if float(annual) != 0.0:
                person_annual_hours += float(annual)
                person['date']['date_'+str(day)]['statistic'] += u'年' + str(float(annual))
            ## 病假时间统计
            sick = person['date']['date_'+str(day)]['sick']
            if float(sick) != 0.0:
                person_sick_hours += float(sick)
                person['date']['date_'+str(day)]['statistic'] += u'病' + str(float(sick))
            ## 早退时间统计
            leav_early = person['date']['date_'+str(day)]['leav_early']
            if -1 == leav_early:
                person['date']['date_'+str(day)]['statistic'] += '[' + u'无签退' + ']' + ' '
            elif (leav_early*(-1)) > 0.0:
                person_leav_early_minutes += float(leav_early*(-1))
                person['date']['date_'+str(day)]['statistic'] += u'早退' + str(float(leav_early*(-1)))
#                 if leav_early > 500:
#                     person['common'] += u'\u2660' + str(day)
            ## 加班时间统计
            overwork = person['date']['date_'+str(day)]['overwork']
            if overwork != 0.0:
                person['date']['date_'+str(day)]['statistic'] += u'加' + str(float(overwork))
            ## 迟到时间统计
            late = person['date']['date_'+str(day)]['late']
            if late > 0.0:
                if late > 0.0 and late <= 10.0:
                    person_late_minutes['below_10'] += 1
                elif late > 10.0 and late < 30.0:
                    person_late_minutes['between_10_30'] += 1
                else:
                    person_late_minutes['up_30'] += u'\u2605' + str(day) + " "
            elif late == -1:
                person['date']['date_'+str(day)]['statistic'] += '[' + u'无签到' + ']' + ' '
            ## 外勤统计       
            if (person['date']['date_'+str(day)]['outside']) and (Date(work_year, work_month, day).isWeekDay()):
                person['common'] += '[' + str(day) + ']' + person['date']['date_'+str(day)]['outside'] + ' '
                            
            ## 无任何异常情况时显示为正常
            if off <= 0.0 and leave <= 0.0 and \
                annual <= 0.0 and sick <= 0.0 and leav_early == 0.0 and\
                overwork <= 0.0:
                person['date']['date_'+str(day)]['statistic'] = u'正常'
        ## 月度情况汇总
        person['off_time']    = person_off_hours
        person['leave_time']  = person_leave_hours
        person['sick_time']   = person_sick_hours
        person['annual_time'] = person_annual_hours  
        if person_late_minutes['below_10'] >= 3:
            person_late_minutes['below_10'] -= 3
        person['late_time']   = person_late_minutes['below_10'] + person_late_minutes['between_10_30']
        person['common']      += person_late_minutes['up_30'] + ' '   
        
    return person_list_cp   

def generateTableHeader(year, month):
    '''
    Generate table header for the file to be written
    @ year  : the year to be statistic
    @ month : the month to be statistic
    '''
    header = [u'序号', u'姓名', u'入职日期', u'月份']
    for day in range(1, Date(year, month, 1).hasDays()+1):
        date = Date(year, month, day)
        if date.isWeekDay():
            header.append(day)
    rest = [u'迟到次数', u'调休', u'事假', u'病假', u'年假', u'截止上月调休结余', u'现结余', u'备注']
    for item in rest:
        header.append(item) 
    
    return header

def setTableBorders():
    border = xlwt.Borders()
    border.top    = 1
    border.bottom = 1
    border.left   = 1
    border.right  = 1
    
    return border

def setTableStyle():
    style = xlwt.XFStyle()
    style.borders = setTableBorders()
    
    ## 设置居中
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER
    align.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = align
    
    return style

def writePersonDataTable(file_name, person_list, sheet_name = 'Sheet1'):
    '''
    Write each person information in the person list to sheet name which has
    been created in the specific file
    @ file_name   : the file will be created to save information
    @ person_list : it contains specific information about each person in the company
    @ sheet_name  : which sheet will be written to in the file, which has default
                    name, it is 'Sheet1'
    '''
    try:
        write_table = WriteTable(file_name, sheet_name)
    except OpenFileError, e:
        print e.getErrorString()
    
    if not person_list:
        raise TableObjectEmptyError("<writePersonDataTable> : Person list is empty")
    
    ## 逐人逐项写入文件
    if write_table.isTableReady():
        ## 写入标题
        header = generateTableHeader(person_list[1]['year'], person_list[2]['month'])
        write_table.setHHeader(header)
        
        style = setTableStyle()
        ## 写入信息内容
        for key in sorted(person_list.keys()):
            ## key可以代表行号, col_index表示列数
            col_index = 0
            ## 写序号
            write_table.setValueWithStyle(key, col_index, person_list[key]['id'], style)
            ## 写姓名
            write_table.setValueWithStyle(key, col_index+1, person_list[key]['name'], style)
            ## 写入职日期
            pass
            write_table.setValueWithStyle(key, col_index+2, "", style)
            ## col_index + 2
            ## 写月份
            write_table.setValueWithStyle(key, col_index+3, unicode(person_list[key]['month'])+u'月', style)
            ##col
            col_index += 4
            ## 写入工作日统计情况
            for day in header[header.index(1) : header.index(u'迟到次数')]:
                day_date = 'date_' + str(day)
                write_table.setValueWithStyle(key, col_index, person_list[key]['date'][day_date]['statistic'], style)
                col_index += 1
            ## 写入迟到次数
            write_table.setValueWithStyle(key, col_index, person_list[key]['late_time'], style)
            ## 写入调休次数
            write_table.setValueWithStyle(key, col_index+1, person_list[key]['off_time'], style)
            ## 写入事假时间
            write_table.setValueWithStyle(key, col_index+2, person_list[key]['leave_time'], style)
            ## 写入病假时间
            write_table.setValueWithStyle(key, col_index+3, person_list[key]['sick_time'], style)
            ## 写入年假时间
            write_table.setValueWithStyle(key, col_index+4, person_list[key]['annual_time'], style)
            ## 写入截止上月调休结余
            pass
            write_table.setValueWithStyle(key, col_index+5, "", style)
            ## 写入现结余
            pass
            write_table.setValueWithStyle(key, col_index+6, "", style)
            ## 写入备注
            write_table.setValueWithStyle(key, col_index+7, person_list[key]['common'], style)
            
            write_table.setTableCellHeight(key, 1000)
        ## 同步写入文件
        write_table.writeToFile()

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
                person_total_list[person_id]['date']['date_'+str(work_day)]['leav_early']  = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['off']         = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['leave']       = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['sick']        = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['annual']      = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['overwork']    = overtime
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
                    person_total_list[person_id]['date']['date_'+str(work_day)]['overwork']   = 0
                    person_total_list[person_id]['date']['date_'+str(work_day)]['leav_early'] = early_overtime_minutes## (early_overtime_minutes * (-1))
                
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
                    person_total_list[person_id]['date']['date_'+str(work_day)]['annual'] = absent_hour
                ## 外勤情况
                person_total_list[person_id]['date']['date_'+str(work_day)]['outside'] = person_dict[u'外勤及其他异常说明']
                
                ## 上个月调休结余时间（后续需要再补充）
                pass
                ## End 上个月调休结余时间    
                
            if person_id > last_person_id:
                last_person_id = person_id
        except (TableObjectEmptyError, TableHeaderEmptyError,\
                TableDataEndError, TableHeaderLengthError) as e:
            print e.getErrorString()

    ## 得到汇总后的人员情况字典
    person_lastest_list = doPersonStatisticsData(person_total_list)
    ## 将汇总情况写入文件
    writePersonDataTable(u'测试文件.xls', person_lastest_list)
    ## read_table.displayPersonObj(person_lastest_list[1])
    
    ## 打印工具结尾
    StatisticsAttendance.printToolLogoEnd()

if __name__ == "__main__":
    main()