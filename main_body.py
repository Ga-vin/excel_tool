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
    ## ��鴫������Ƿ�Ϊ��
    if not obj:
        raise TableObjectEmptyError("<getSpecificLineData> : Table object is None") 
    
    ## ��鴫��ı�����ͷ�Ƿ�Ϊ��
    if not table_header:
        raise TableHeaderEmptyError("<getSpecificLineData> : Table header is None")
    
    temp_obj = dict()
    data = obj.getNextLineRow()
    if not object:
        ## �����������һ��
        raise TableDataEndError("<getSpecificLineData> : End of Table object")
    elif len(table_header) != len(data):
        ## ���ⳤ�Ⱥ����ݳ��Ȳ�һ��
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
    
    if table_obj.addPersonNameToSets(person_dict[u"����"]):
        person_id += 1
        person_obj = table_obj.generatePersonObj()
        if not person_obj:
            print "[*] Create person object failed"
            return None
        person_obj["name"] = person_dict[u"����"]
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
    ## ���ǩ��ʱ��Ԫ��
    sign_time_tuple     = xlrd.xldate_as_tuple(sign_time, 0)
    ## ��ù涨ǩ��ʱ��Ԫ��
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
    ## ���ǩ��ʱ��
    sign_time_tuple     = xlrd.xldate_as_tuple(sign_time, 0)
    ## ��ù涨ǩ��ʱ��
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
    ## ���ǩ��ʱ��
    come_time_tuple  = xlrd.xldate_as_tuple(come_time, 0)
    ## ����뿪ʱ��
    leave_time_tuple = xlrd.xldate_as_tuple(leave_time, 0)
    come_minutes  = come_time_tuple[tuple_list.index('hour')]*60 + come_time_tuple[tuple_list.index('minute')]
    leave_minutes = leave_time_tuple[tuple_list.index('hour')]*60 + leave_time_tuple[tuple_list.index('minute')]
    total_minutes = leave_minutes - come_minutes
    return total_minutes

## ����ڳ���
def main():
    read_table_name       = "record_total.xlsx"
    read_table_sheet_name = "specific"
    person_total_list     = dict()
    
    ## ����������ӵı��
    person_id = 0
    start_from_row = 1
    
    ## ��ӡ����
    StatisticsAttendance.printToolLogoHeader()
    
    ## �򿪶�ȡ���
    read_table = StatisticData(read_table_name, read_table_sheet_name)
    ## ��ȡ���ͷ
    read_table_header = read_table.getHorizonTitle()
    if not read_table_header:
        print '<main> : Data Table Header is Empty'
        ## ���Ϊ�ձ������
        sys.exit()
        
    ## ���ö�ȡ���ݿ�ʼ�У�����ֱ�Ӷ�ȡ��һ��
    read_table.setCurrentRowIndex(start_from_row)
    ## ��ʼ��ȡͳ������
    last_person_id = 1
    temp = 0
    while True:
        try:
            is_new_person = False
            ## ��ȡһ�е����ݲ�ת��Ϊһ���ֵ�
            person_dict = getSpecificLineData(read_table, read_table_header)
            
            if not person_dict[u'����']:
                print '[*] End of the table'
                break
            
            ## ���������������
            person_id, person_obj_temp = addPersonName(read_table, person_id, person_dict)
            if person_obj_temp:
                ## ��ӵ���Ա�б���
                is_new_person = True
                person_total_list[person_id] = person_obj_temp
            
            date_string = person_dict[u'����']
            date_temp   = addPersonDate(date_string)
            if not date_temp:
                print "[*] Date convert fails. System will be terminated"
                sys.exit()
            ## ��������ˣ�����������
            if is_new_person:
                is_neww_person = False
                person_total_list[person_id]['year']  = date_temp.getYear()
                person_total_list[person_id]['month'] = date_temp.getMonth()
            
            ## ��ȡ����Ĺ�����
            work_day = date_temp.getDay()
            if not date_temp.isWeekDay():
                ## �ǹ����գ���Ӱ�
                overtime = addPersonOvertimeNotWorkday(person_dict[u'ǩ��ʱ��'], person_dict[u'ǩ��ʱ��'])
                person_total_list[person_id]['date']['date_'+str(work_day)]['late'] = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['leav_early'] = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['off']        = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['leave']      = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['sick']       = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['annual']     = 0
                person_total_list[person_id]['date']['date_'+str(work_day)]['outside']    = overtime
            else:
                late_minutes = addPersonLateMinutes(person_dict[u'ǩ��ʱ��'], person_dict[u'�涨�ϰ�ʱ��'])
                early_overtime_minutes = addPersonLeaveEarlyMinutes(person_dict[u'ǩ��ʱ��'], person_dict[u'�涨�°�ʱ��'])
                ## �ٵ�ʱ��
                person_total_list[person_id]['date']['date_'+str(work_day)]['late'] = late_minutes
                
                if early_overtime_minutes > 0:
                    ## �Ӱ�����
                    person_total_list[person_id]['date']['date_'+str(work_day)]['overwork']   = early_overtime_minutes
                    person_total_list[person_id]['date']['date_'+str(work_day)]['leav_early'] = 0
                else:
                    ## ���˵����
                    person_total_list[person_id]['date']['date_'+str(work_day)]['overwork'] = 0
                    person_total_list[person_id]['date']['date_'+str(work_day)]['leav_early'] = (early_overtime_minutes * (-1))
                
                ## �¼�ʱ��
                person_total_list[person_id]['date']['date_'+str(work_day)]['leave'] = int(person_dict[u'�¼�ʱ��'])
                ## ����ʱ��
                person_total_list[person_id]['date']['date_'+str(work_day)]['sick'] = int(person_dict[u'����ʱ��'])
                ## ����/���ʱ��
                absent_hour = float(person_dict[u'��н��ʱ��   �����ݻ�������'])
                absent_reason = person_dict[u'��н�ٱ�']
                if u'����' == absent_reason:
                    person_total_list[person_id]['date']['date_'+str(work_day)]['off']    = absent_hour
                    person_total_list[person_id]['date']['date_'+str(work_day)]['annual'] = 0
                elif u'���' == absent_reason:
                    person_total_list[person_id]['date']['date_'+str(work_day)]['off']    = 0
                    person_total_list[person_id]['date']['date_'+str(work_day)]['annual'] = absent_reason
                ## �������
                person_total_list[person_id]['date']['date_'+str(work_day)]['outside'] = person_dict[u'���ڼ������쳣˵��']    
                
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
    print u'����%d��'%person_id
    ## ��ӡ���߽�β
    StatisticsAttendance.printToolLogoEnd()

if __name__ == "__main__":
    main()