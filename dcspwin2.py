#coding=utf-8

import wx
import wx.lib.filebrowsebutton as fb
import pandas as pd
import os
from settings2 import Settings
import dcfunctions as dcf
from rundata2 import RunData
          
class DcspWin(wx.App):        
    def OnInit(self):
        self.mset = Settings()
        self.write1 = pd.ExcelWriter(self.mset.rlt)
        self.frame = wx.Frame(None, -1, title="Auto HW to Services", size=(1300,800), style=wx.DEFAULT_FRAME_STYLE)
        
        self.statusbar = self.frame.CreateStatusBar(2, wx.STB_SIZEGRIP)
        self.statusbar.SetStatusWidths([-2, -3])
        self.statusbar.SetStatusText("DCSP Work on HW", 0)
        self.statusbar.SetStatusText("Welcome To Dell DCSP Team!", 1)
        
        IMG_BASE_DIR = os.path.join(os.getcwd(), 'img')
        iconfilename1 = os.path.join(IMG_BASE_DIR, 'david.ico')
        icon1 = wx.Icon(iconfilename1, wx.BITMAP_TYPE_ICO)
        self.frame.SetIcon(icon1)
        self.frame.Show(True)
        self.SetTopWindow(self.frame)
        self.frame.SetMaxSize((1300,800))
        self.frame.CentreOnScreen()
        self.panel = wx.Panel(self.frame)
        
        st1 = wx.StaticText(self.panel, -1, "Open EMC HW Funnel File: ", size=(240,-1), style=wx.ALIGN_RIGHT)
#        st2 = wx.StaticText(self.panel, -1, "Open David's Funnel List File: ", size=(240,-1), style=wx.ALIGN_RIGHT)
        st3 = wx.StaticText(self.panel, -1, "Select result file path directory: ", size=(240,-1), style=wx.ALIGN_RIGHT)
        fb1 = fb.FileBrowseButton(self.panel, -1, size=(1000,-1), labelText=" ", startDirectory=self.mset.src1, changeCallback=self.OnFB1)
#        fb2 = fb.FileBrowseButton(self.panel, -1, size=(1000,-1), labelText=" ", startDirectory=self.mset.src2, changeCallback=self.OnFB2)
        fb3 = fb.DirBrowseButton(self.panel, -1, size=(1000,-1), labelText=" ", startDirectory=self.mset.rltdir, changeCallback=self.OnFB3)
        
        slabel = wx.StaticText(self.panel, -1, "Select current working week:\t", size=(240,-1), style=wx.ALIGN_RIGHT)
        self.sel1 = wx.Choice(self.panel, -1, choices=self.mset.selList)
        self.sel1.Bind(wx.EVT_CHOICE, self.EvtChoice)
        sdcsplabel = wx.StaticText(self.panel, -1, "Select DCSP: ", size=(240,-1), style=wx.ALIGN_RIGHT)
        self.sdcsp = wx.Choice(self.panel, -1, choices=self.mset.dcsp_names_all)
        
        self.fb4 = wx.TextCtrl(self.panel, -1, style=wx.TE_MULTILINE | wx.TE_READONLY, size=(1250,510))
        self.mset.logs = dcf.mergestrs(dcf.show_welcome0_2())
        self.fb4.AppendText(self.mset.logs)
    
        self.shortlabel = wx.StaticText(self.panel, -1, "<---Click to generate funnel ... ", size=(600,-1))
        shortfunnel = wx.Button(self.panel, -1, "Start", size=(240,-1))
        shortfunnel.Bind(wx.EVT_BUTTON, self.ShortFunnel)
        
        fb5 = wx.Button(self.panel, -1, "Save log to a file (No Click, No Save) ...",)
        fb5.Bind(wx.EVT_BUTTON, self.OnFB5)
        
        fb6 = wx.Button(self.panel, -1, "Quit", size=(240,-1))
        fb6.Bind(wx.EVT_BUTTON, self.OnExitApp)
        
        self.logshow = wx.TextCtrl(self.panel, -1, size=(600,-1), style=wx.TE_READONLY)
        self.logshow.AppendText('You haven\'t set log file name!')
        
        box = wx.GridBagSizer(3, 0)
        box.Add(st1, pos=(0,0))
        box.Add(fb1, pos=(0,1), span=(1,3))
 #       box.Add(st2, pos=(1,0))
 #       box.Add(fb2, pos=(1,1), span=(1,3))
        box.Add(st3, pos=(2,0))
        box.Add(fb3, pos=(2,1), span=(1,3))
        box.Add(slabel, pos=(3,0))
        box.Add(self.sel1, pos=(3,1))
        box.Add(sdcsplabel, pos =(3,2))
        box.Add(self.sdcsp, pos=(3,3))
        box.Add(self.fb4, pos=(4,0), span=(1,4))
        box.Add(shortfunnel, pos=(5,0))
        box.Add(self.shortlabel, pos=(5,1), span=(1,3))
        box.Add(fb6, pos=(6,0))        
        box.Add(fb5, pos=(6,1))
        box.Add(self.logshow, pos=(6,2), span=(1,2))
       
        self.panel.SetSizer(box)
        self.panel.Fit()
        return True

    def OnExitApp(self, evt):
        if self.mset.completed >= 1:
            self.frame.Destroy()
            return True
        else:
            dlg = wx.MessageDialog(self.panel, message="Your work hasn't been completed!", caption="Warning")
            if dlg.ShowModal() == wx.ID_OK:
                dlg.Destroy()
                return False

    def OnFB1(self, evt):
        self.mset.src1 = evt.GetString()
        self.mset.src1_s = 1
        return True
    
      
    def OnFB3(self, evt):
        self.mset.tmp1 = evt.GetString()
        self.mset.rlt = os.path.join(self.mset.tmp1, self.mset.crtwk + ' - EMC Result.xlsx')
        self.write1 = pd.ExcelWriter(self.mset.rlt)
        self.mset.rlt_s = 1 
        return True
    
    def OnFB5(self, evt):
        if self.mset.completed >= 1 and self.mset.logf_s == -1:
            dlg = wx.FileDialog(self.panel, message="Save log file as ...", defaultDir=os.getcwd(), defaultFile="", wildcard=self.mset.wildcardlog, style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
            dlg.SetFilterIndex(0)
            if dlg.ShowModal() == wx.ID_OK:
                self.mset.logf_s = 1
                self.mset.logf = dlg.GetPath()
                self.fb4.SaveFile(filename=self.mset.logf)
                self.logshow.Clear()
                self.logshow.AppendText(self.mset.logf)
        else:
            dlg = wx.MessageDialog(self.panel, message="Please wait for the program running finished!", caption="Warning")
            if dlg.ShowModal() == wx.ID_OK:
                dlg.Destroy()

    def EvtChoice(self, evt):
        self.mset.crtwk = self.sel1.GetStringSelection()
        tmp1 = self.mset.crtwk
        self.mset.wtd_work = tmp1[-2:]
        self.mset.qtr_working1 = tmp1[0:-4]
        if self.mset.rlt_s == 1:
            self.mset.rlt = os.path.join(self.mset.tmp1, tmp1 + ' - EMC Result.xlsx')
            self.write1 = pd.ExcelWriter(self.mset.rlt)
        self.mset.selList_s = 1
        return True
    
    def ShortFunnel(self, evt):
        tmpbool = (self.mset.src1_s == 1) and (self.mset.rlt_s == 1) and (self.mset.selList_s == 1)
        if tmpbool:
            self.mset.rdy = 1
            self.mset.src1_s = -1
            self.mset.rlt_s = -1
            self.mset.selList_s = -1
            self.shortlabel.SetLabel('The result file name is: ' + self.mset.rlt)
            run1 = RunData(self.mset, self.write1)
            run1.run_data(self, self.mset)
            self.mset.rdy = -1
            self.logshow.Clear()
            self.logshow.AppendText('You haven\'t set log file name!')
            self.sel1.SetSelction = -1
            self.mset.logf_s = -1
            return True
        else:
            dlg = wx.MessageDialog(self.panel, message="Your haven\'t completed your preparation!", caption="Warning")
            if dlg.ShowModal() == wx.ID_OK:
                dlg.Destroy()
                return False

if __name__ == '__main__':
    dcf.prt_head()
    app = DcspWin()
    app.MainLoop()