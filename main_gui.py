# -*- coding: gb18030 -*-
'''
    Created on 2016-01-11

    @author: Gavin.Bai
    @note: Main GUI interface for the tool to statistic excel information
    @version: v1.0
    @Modify:
    @License: (C)GPL
'''
## ----------------------------------------------------------------------
import wx
import time
import sys
import os
import main_body

class Frame(wx.Frame):
    img_list = ['jiaoxue.jpg', 'jiaoxue2.jpg', 'jiaoxue3.jpg', 'jiaoxue4.jpg']
    read_file_path     = ""
    read_file_name     = ""
    r_write_file_name  = ""
    r_write_sheet_name = ""
    
    def __init__(self, parent, id, title):
        super(Frame, self).__init__(parent, id, title,
                                    size = wx.Size(960, 350))
        self.__panel = wx.Panel(self, -1)
        self.addImage()
        
        label_1 = wx.StaticText(self.__panel, -1, u'ԭʼ���ݱ��λ��', 
                                pos = wx.Point(100, self.img_1.GetHeight() + 30),
                                size = wx.Size(150, 30),
                                style = wx.RAISED_BORDER|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL)
        
        self.orgin_data = wx.TextCtrl(self.__panel, -1,
                                      pos = wx.Point(label_1.GetPosition().x+label_1.GetSize().width+20, self.img_1.GetHeight() + 30),
                                      size = wx.Size(400, 30),
                                      style = wx.TE_READONLY)
        self.find_file  = wx.Button(self.__panel, -1, u'���ļ�...',
                                    pos = wx.Point(self.orgin_data.GetPosition().x+self.orgin_data.GetSize().width+20, self.img_1.GetHeight() + 30),
                                    size = wx.Size(150, 30))
        label_2 = wx.StaticText(self.__panel, -1, u'�����ļ���',
                                pos = wx.Point(100, label_1.GetPosition().y+label_1.GetSize().GetHeight()+30),
                                size = wx.Size(150, 30),
                                style = wx.RAISED_BORDER|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL)
        self.write_file_name = wx.TextCtrl(self.__panel, -1,
                                           pos = wx.Point(label_2.GetPosition().x+label_2.GetSize().width+20, label_2.GetPosition().y),
                                           size = wx.Size(200, 30))
        label_3 = wx.StaticText(self.__panel, -1, u'Sheet����',
                                pos = wx.Point(self.write_file_name.GetPosition().x+self.write_file_name.GetSize().width+20, label_2.GetPosition().y),
                                size = wx.Size(150, 30),
                                style = wx.RAISED_BORDER|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL)
        self.write_sheet_name = wx.TextCtrl(self.__panel, -1,
                                            pos = wx.Point(label_3.GetPosition().x+label_3.GetSize().width+20, label_2.GetPosition().y),
                                            size = wx.Size(180, 30))
        self.btn_ok = wx.Button(self.__panel, -1, u'ִ��ͳ��(&R)',
                                pos = wx.Point(label_2.GetPosition().x+100, label_2.GetPosition().y+label_2.GetSize().GetHeight()+30),
                                size = wx.Size(150, 30))
        self.btn_ok.Enable(False)
        self.btn_reset = wx.Button(self.__panel, -1, u'����(&C',
                                   pos = wx.Point(self.btn_ok.GetPosition().x+self.btn_ok.GetSize().width+30, self.btn_ok.GetPosition().y),
                                   size = wx.Size(150, 30))
        self.btn_quit  = wx.Button(self.__panel, -1, u'�˳�(&X)',
                                   pos = wx.Point(self.btn_reset.GetPosition().x+self.btn_reset.GetSize().width+30, self.btn_reset.GetPosition().y),
                                   size = wx.Size(150, 30))
        
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.btn_quit)
        self.Bind(wx.EVT_BUTTON, self.OnChooseFile, self.find_file)
        self.Bind(wx.EVT_BUTTON, self.OnDoReset, self.btn_reset)
        self.Bind(wx.EVT_BUTTON, self.OnDoStatistic, self.btn_ok)
        ## ������ʱ��
        ##self.timer_change_bmp = wx.Timer(self)
        ##self.Bind(wx.EVT_TIMER, self.OnTimerEvent, self.timer_change_bmp)
        ##self.timer_change_bmp.Start(2000)
        
    def getImgNextOrder(self):
        img_list_cp = self.img_list[1:]
        img_list_cp.append(self.img_list[0])
        self.img_list = img_list_cp
        return img_list_cp
    
    def addImage(self):
        self.img_1 = wx.Image("jiaoxue.jpg", wx.BITMAP_TYPE_JPEG).ConvertToBitmap()
        self.img_2 = wx.Image("jiaoxue2.jpg", wx.BITMAP_TYPE_JPEG).ConvertToBitmap()
        self.img_3 = wx.Image("jiaoxue3.jpg", wx.BITMAP_TYPE_JPEG).ConvertToBitmap()
        self.img_4 = wx.Image("jiaoxue4.jpg", wx.BITMAP_TYPE_JPEG).ConvertToBitmap()
        self.bmp1 = wx.StaticBitmap(parent = self.__panel, bitmap = self.img_1)
        self.bmp2 = wx.StaticBitmap(parent = self.__panel, bitmap = self.img_2, pos = (self.img_1.GetWidth(), 0))
        self.bmp3 = wx.StaticBitmap(parent = self.__panel, bitmap = self.img_3, pos = (self.img_1.GetWidth()*2, 0))
        self.bmp4 = wx.StaticBitmap(parent = self.__panel, bitmap = self.img_4, pos = (self.img_1.GetWidth()*3, 0))
        
    def OnTimerEvent(self, evt):
        self.img_list = self.getImgNextOrder()
        self.img_1 = wx.Image(self.img_list[0], wx.BITMAP_TYPE_JPEG).ConvertToBitmap()
        self.img_2 = wx.Image(self.img_list[1], wx.BITMAP_TYPE_JPEG).ConvertToBitmap()
        self.img_3 = wx.Image(self.img_list[2], wx.BITMAP_TYPE_JPEG).ConvertToBitmap()
        self.img_4 = wx.Image(self.img_list[3], wx.BITMAP_TYPE_JPEG).ConvertToBitmap()
        self.bmp1 = wx.StaticBitmap(parent = self, bitmap = self.img_1)
        self.bmp2 = wx.StaticBitmap(parent = self, bitmap = self.img_2, pos = (self.img_1.GetWidth(), 0))
        self.bmp3 = wx.StaticBitmap(parent = self, bitmap = self.img_3, pos = (self.img_1.GetWidth()*2, 0))
        self.bmp4 = wx.StaticBitmap(parent = self, bitmap = self.img_4, pos = (self.img_1.GetWidth()*3, 0))
    
    def OnExit(self, evt):
        '''
                        �˳�����
        '''
        sys.exit()
    
    def OnChooseFile(self, evt):
        '''
                        ѡ��Ҫͳ�Ƶ�ԭ����excel�ļ�
        '''
        dialog = wx.FileDialog(None, u'��ѡ��ԭʼ�����ļ�', wx.EmptyString, 
                               wx.EmptyString, "*.xls;*.xlsx", 
                               style = wx.OPEN|wx.HIDE_READONLY)
        if wx.ID_OK == dialog.ShowModal():
            self.read_file_path = dialog.GetPath()
            if not self.read_file_path:
                wx.MessageDialog(None, u'��û����ѡ���ļ�', u'����').ShowModal()
            else:
                self.orgin_data.SetValue(self.read_file_path)
                if self.read_file_path:
                    self.read_file_name = self.read_file_path.split('\\')[-1]
                self.btn_ok.Enable(True)
    
    def OnDoStatistic(self, evt):
        '''
                        ִ��ͳ�ƹ���
        '''
        if self.write_file_name.IsEmpty():
            wx.MessageDialog(None, u'д���ļ�������Ϊ�գ�����Ҫ���ļ���', u'����').ShowModal()
            return 
        else:
            raw_name = self.write_file_name.GetValue()
            if '.' in raw_name:
                self.r_write_file_name = raw_name.split('.')[0]
            else:
                self.r_write_file_name = raw_name
        
        if self.write_sheet_name.IsEmpty():
            wx.MessageDialog(None, u'д���ļ�������Ϊ��', u'����').ShowModal()
            return 
        else:
            self.r_write_sheet_name = self.write_sheet_name.GetValue()
        
        if self.r_write_file_name and self.r_write_sheet_name:
            main_body.main(self.r_write_file_name, self.r_write_sheet_name)
            self.btn_ok.Enable(False)
    
    def OnDoReset(self, evt):
        '''
                        ����Ѿ���д��д���ļ����ͱ����ƣ��Լ�ѡ����ļ�
        '''
        self.write_file_name.Clear()
        self.write_sheet_name.Clear()
        self.orgin_data.Clear()
        
class App(wx.App):
    
    def OnInit(self):
        frame = Frame(parent = None, id = -1, title = u'����ͳ�ƹ��� ----�Ϻ�����ҽ�ƿƼ����޿Ƽ���˾���²�')
        frame.Show()
        self.SetTopWindow(frame)
        return True
    
def main():
    app = App()
    app.MainLoop()
    
if __name__ == "__main__":
    main()
        