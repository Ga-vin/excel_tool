# -*- coding: gb18030 -*-
'''
    Created on 2016-01-08

    @author: Gavin.Bai
    @note: Date type which is to deal date in the tool
    @version: v1.0
    @Modify:
    @License: (C)GPL
'''
## ----------------------------------------------------------------------
## ----------------------------------------------------------------------
## Package
import datetime
import sys

class DateInvalidError(Exception):
    '''
    Error for deal with invalid date object
    '''
    promt_information = ""
    def __init__(self, info):
        super(DateInvalidError, self).__init__()
        self.promt_information = info
    
    def getErrorString(self):
        return "<DateInvalidError: <%s>" % self.promt_information
    
class Date(object):
    '''
    Date class type to handle all kinds of operations for date
    @year   : year
    @month  : month
    @day    : day 
    '''
    __year  = 0
    __month = 0
    __day   = 0
    def __init__(self, year, month, day):
        '''
        Constructor for the Date class type
        '''
        self.setYear(year)
        self.setMonth(month)
        self.setDay(day)
        
    def __repr__(self):
        '''
        Format the instance express
        '''
        return "%d-%d-%d" % (self.getYear(), self.getMonth(), self.getDay())
        
    def setYear(self, year):
        '''
        Set the year according to the parameter passed in 
        @year : which is to be set 
        '''
        if year < 0 or year > 3000:
            raise DateInvalidError("Year: %d"%year)
        self.__year = year
        
    def getYear(self):
        '''
        Get the current year
        '''
        return self.__year
    
    def setMonth(self, month):
        '''
        Set the month according to the parameter passed in which should be 
        between 1 and 12
        @month : which is to be set
        '''
        if month < 1 or month > 12:
            raise DateInvalidError("Month between 1-12")
        self.__month = month
        
    def getMonth(self):
        '''
        Get current month
        '''
        return self.__month
        
    def setDay(self, day):
        '''
        Set the day according to the parameter passed in which should be valid
        acompany with the month
        @day  : which is to be set
        '''
        if self.getMonth() in [1, 3, 5, 7, 8, 10, 12]:
            if day > 0 and day < 32:
                self.__day = day
            else:
                raise DateInvalidError("Day in 1, 3, 5, 7, 8, 10, 12 month days are at most 31")
        elif self.getMonth() in [4, 6, 9, 11]:
            if day > 0 and day < 31:
                self.__day = day
            else:
                raise DateInvalidError("Day in 4, 6, 9, 11 month days are at most 30")
        else:
            if day > 29 or day < 0:
                raise DateInvalidError("Day in 2 month days are at most 29")
            elif 29 == day:
                if self.isLeapYear():
                    self.__day = day
                else:
                    raise DateInvalidError("Day in 2 month, and the year is not leap year, days should be at most 28")
            else:
                self.__day = day
    
    def getDay(self):
        '''
        Get current day
        '''
        return self.__day
        
    def isLeapYear(self):
        '''
        Check whether the specific year is leap year
        @Return True if it is, False else not
        '''
        if self.getYear()%400 == 0 or \
        (self.getYear()%4 == 0 and self.getYear()%100 != 0):
            return True
        else:
            return False
        
    def getWeekday(self):
        '''
        Get current weekday according to the year, month and day
        with number between 1 and 7
        '''
        import time
        time_str = "%d-%d-%d" % (self.getYear(), self.getMonth(), self.getDay())
        time_orign = time.strptime(time_str, "%Y-%m-%d")
        return time_orign.tm_wday + 1
    
    def getWeekdayFormat(self):
        '''
        Get current weekday according to the year, month and day
        with unicode string
        '''
        week_list = [u"", u"星期一", u"星期二", u"星期三", u"星期四", u"星期五", u"星期六", u"星期日"]
        return week_list[self.getWeekday()]
    
    def isWeekDay(self):
        '''
        Check whether current day is weekday which is between 1 and 5
        @Return True if it is, False else not
        '''
        week = self.getWeekday()
        if week > 0 and week < 6:
            return True
        else:
            return False

    def hasDays(self):
        '''
        Return the days of the specific month has
        @ month : specific month to check
        '''
        if self.getMonth() in [1, 3, 5, 7, 8, 10, 12]:
            return int(31)
        elif self.getMonth() in [4, 6, 9, 11]:
            return int(30)
        else:
            if self.isLeapYear():
                return int(29)
            else:
                return int(28)
        
def main():
    print 'Normal date object: '
    today = Date(2015, 01, 8)
    print today
    
    print 'Unnormal date object: 2015-01-32'
    try:
        a = Date(2016, 3, 32)
        print a 
    except DateInvalidError, e:
        print e.getErrorString()
    
    try:
        bb = Date(2016, 1, 11)
        print bb
        print bb.getWeekday()
        print bb.getWeekdayFormat()
        if bb.isWeekDay():
            print 'Work'
        else:
            print 'No work'
    except DateInvalidError, e:
        print e.getErrorString()

if __name__ == "__main__":
    main()
    