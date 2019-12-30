#coding=utf-8

import os
import datetime

class Settings():
    def __init__(self):
        BASE_DIR = os.getcwd()
        self.src1 = os.path.join(BASE_DIR, 'data')
        self.src1_s = -1
        self.src2 = self.src1
        self.src2_s = -1
        self.rltdir = self.src1
        self.crtwk = 'FY20Q4WK08'
        self.crtwk_s = -1
        self.rlt = os.path.join(self.rltdir,  self.crtwk + ' - Result.xlsx')
        self.rlt_s = -1
        
        self.rdy = -1
        self.wtd_work = '07'
        self.completed = 0
        self.tmp1 = ''
        self.logf = os.path.join(BASE_DIR, 'funnellog' + str(datetime.datetime.now()) + '.txt')
        self.logs = "" #将日志输出在大窗口中
        self.logf_s = -1
                
        self.tmpdata = os.path.join(self.rltdir, self.crtwk + ' - tmpdata.xlsx')
        
        self.wildcard = "Excel File (*.xlsx)|*.xlsx|"     \
                   "All files (*.*)|*.*"
        self.wildcardlog = "TXT File (*.txt)|*.txt|"     \
                   "All files (*.*)|*.*"
        
        self.selList = ['FY20Q4WK07', 'FY20Q4WK08', 'FY20Q4WK09', 'FY20Q4WK10', 'FY20Q4WK11', 'FY20Q4WK12', 'FY20Q4WK13']
        self.selList_s = -1
        
        self.src1_sheet = 'HW SFDC Funnel'
        self.src2_sheet = 'Funnel List'
        
        self.hw_to_zgx_list = [13, 13, 31, 28, 32, 27, 1, 70, 70, 9, 46, 51, 64, 63, 63, 15, 15, 15, 15, 15, 15, 47, 68, 15, 15, 15, 15, 15]
        self.win_rate_all = ['Lost, Cancelled - 0%', 'Win - 100%', 'Commit - 90%', 'Discover - 10%', 'Qualify - 30%', 'Propose - 60%', 'Order Submitted - 99%', 'Plan - 1%']
        self.win_rate_check_useful = ['Commit - 90%', 'Order Submitted - 99%', 'Propose - 60%', 'Discover - 10%', 'Qualify - 30%']
        self.win_rate_check_drop = ['Lost, Cancelled - 0%', 'Win - 100%', 'Plan - 1%']
        self.win_rate_check_middle = ['Discover - 10%', 'Qualify - 30%']
        self.sales_forecast_all = ['Best Case', 'Closed', 'Pipeline', 'Commit', 'Omitted']
        self.sales_forecast_useful = ['Best Case', 'Pipeline', 'Commit']
        self.for_drop = ['Dled', 'Lost']
        self.dcsp_names_all = ['Chang, David', 'Kong, DeYuan', 'You, Judith', 'Hua, Xuming', 'Tong, Coco']
        self.dcsp_name_checking = ['Chang, David']
        self.dcsp_ntname_checking = ['David_Chang2']
        self.qtr_all = ['FY20Q04', 'FY21Q01', 'FY21Q02', 'FY21Q03', 'FY21Q04']
        self.qtr_working = ['FY20Q04']
        self.qtr_working1 = 'FY20Q4'
        self.dlwk_header = self.qtr_working[0] + 'WK'
        
        self.svc_types = ['CI Deployment', 'Product Attached', 'Residency']