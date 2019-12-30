#coding=utf-8
#这是主程序，可以运行 - Vesion 2.0

import pandas as pd
import math
import dcfunctions as dcf
import xlrd

class RunData():
    def __init__(self, mset, write1):
        self.write1 = write1
        self.fordel = []

    def run_data(self, app, mset):
        mset.logs = dcf.mergestrs(dcf.show_welcome_2(mset.src1, mset.rlt))
        app.fb4.AppendText(mset.logs)
        mset.logs = "####Read source EMC HW Funnel into DataFrame Object:\n"
        app.fb4.AppendText(mset.logs)
        din1 = pd.read_excel(mset.src1, sheet_name=mset.src1_sheet)
        lines_hw_funnel = din1.shape[0]
        columns_hw_funnel = din1.shape[1]
        mset.logs = dcf.mergestrs(dcf.type_src_2(lines_hw_funnel, columns_hw_funnel))
        app.fb4.AppendText(mset.logs)
    #################################################
        mset.logs = dcf.mergestrs(dcf.print_start_search(mset.src1_sheet))
        app.fb4.AppendText(mset.logs)
        sheetname1 = mset.src1_sheet 
        dcf.sheetbackup(self.write1, din1, mset.src1_sheet)
    #################################################
        app.fb4.AppendText("####开始处理第一轮来自EMC HW Funnel的数据：\n")
        mset.logs = dcf.mergestrs(dcf.print_start_search(sheetname1))
        app.fb4.AppendText(mset.logs)
        app.fb4.AppendText("******** 找出EMC HW中的有效列，生成dout2，并修改 ********\n")
         
        tocols0 = []
        tocols = []
        for n in din1.columns: tocols0.append(n)
        for p in mset.hw_to_zgx_list: tocols.append(tocols0[p])
        dout2 = pd.DataFrame(din1, columns=tocols)
         
        lentmp = dout2.shape[0]
        for q in range(lentmp):
            dout2.iloc[q,6] = mset.qtr_working1 + 'WK{:02d}'.format(int(dout2.iloc[q,6]) - 202039)
            if int(dout2.iloc[q,9]) <=1 and int(dout2.iloc[q,10]) <=1 and int(dout2.iloc[q,11]) <=1:
                dout2.iloc[q,9] = math.ceil(dout2.iloc[q,7]/10000)
            else:
                dout2.iloc[q,9] = int(dout2.iloc[q,9]/1000)
            for qt1 in [7, 8, 10, 11]: dout2.iloc[q,qt1] = int(dout2.iloc[q,qt1]/1000)
            dout2.iloc[q,5] = dout2.iloc[q,5].split()[-1]
        dout2.sort_values(by=['Prod Att Rev', 'Forecast_Status'], ascending=False, inplace=True)
        app.fb4.AppendText("####没有删除动作，只是改了一下WK和将没有PA的，用CR 10%估计值填上。\n")    
        sheetname2 = 'sheet_dout2'    
        dcf.sheetlast(self.write1, dout2, sheetname2)
        mset.logs = dcf.mergestrs(dcf.print_end_search(sheetname1))
        app.fb4.AppendText(mset.logs)
    #################################################
        app.fb4.AppendText("####再对生成的dout2进行整理，相加去重，生成dout3\n")
        mset.logs = dcf.mergestrs(dcf.print_start_search(sheetname2))
        app.fb4.AppendText(mset.logs)
            
        for r in range(lentmp):
            if r not in self.fordel:
                for s in range(r+1, lentmp):
                    if dout2.iloc[r,1] == dout2.iloc[s,1] and dout2.iloc[r,10] <= 1 \
                            and dout2.iloc[r,11] <= 1 and dout2.iloc[s,10] <= 1 \
                            and dout2.iloc[s,11] <= 1 and (s not in self.fordel):
                        self.fordel.append(s)
                        for tmpxx in range(7, 12): dout2.iloc[r,tmpxx] += dout2.iloc[s,tmpxx]
     
        app.fb4.AppendText("These lines will be summed firstly! \n") 
        mset.logs = dcf.mergestrs(dcf.prt_fordel(self.fordel, sheetname2))
        app.fb4.AppendText(mset.logs)
        mset.logs = dcf.mergestrs(dcf.print_end_search(sheetname2))
        app.fb4.AppendText(mset.logs)    
        
        mset.logs = dcf.mergestrs(dcf.print_start_drop(len(self.fordel), sheetname2))
        app.fb4.AppendText(mset.logs)
         
        sheetname3 = 'sheet_out3'
        dout3 = dcf.droplns(self.write1, dout2, self.fordel, sheetname3)
         
        mset.logs = dcf.mergestrs(dcf.print_end_drop(len(self.fordel), sheetname2))
        app.fb4.AppendText(mset.logs)
    
        app.fb4.AppendText("####对dout3的数据开始分析……\n")           
    #################################################
        lentmp = dout3.shape[0]
        tmplogs = {'All':'Total Funnel with 100% won deals PA, Res and CI numbers are:\t\tPA--$',
                   '100%':'Total 100% won deals PA, Res and CI numbers are:\t\t\tPA--$',
                   '90%':'Total 90% win rate deals PA, Res and CI numbers are:\t\t\tPA--$',
                   '60%':'Total 60% win rate deals PA, Res and CI numbers are:\t\t\tPA--$',
                   '30%':'Total 30% win rate deals PA, Res and CI numbers are:\t\t\tPA--$',
                   '<30%':'Total less than 30% win rate deals PA, Res and CI numbers are:\t\tPA--$',
                   }
        tmpdata = [[0 for i in range(3)] for i in range(6)] #3 columns and 6 lines
#       tmpdata = [[0, 0, 0],[0, 0, 0],[0, 0, 0],[0, 0, 0],[0, 0, 0],[0, 0, 0]]                
        
        key1 = 0
        for key in tmplogs.keys():
            for x in range(lentmp):
                if key == 'All':
                    for xxx in range(9,12): tmpdata[0][xxx-9] += int(dout3.iloc[x,xxx])
                elif key == '100%' and dout3.iloc[x,5] == '100%':
                    for xxx in range(9,12): tmpdata[1][xxx-9] += int(dout3.iloc[x,xxx])
                elif key == '90%' and dout3.iloc[x,5] == '90%':
                    for xxx in range(9,12): tmpdata[2][xxx-9] += int(dout3.iloc[x,xxx])
                elif key == '60%' and dout3.iloc[x,5] == '60%':
                    for xxx in range(9,12): tmpdata[3][xxx-9] += int(dout3.iloc[x,xxx])
                elif key == '30%' and dout3.iloc[x,5] == '30%':
                    for xxx in range(9,12): tmpdata[4][xxx-9] += int(dout3.iloc[x,xxx])    
                elif key == '<30%' and int(dout3.iloc[x,5][:-1]) < 30:
                    for xxx in range(9,12): tmpdata[5][xxx-9] += int(dout3.iloc[x,xxx])
                else: continue        
            mset.logs = tmplogs[key] + format(tmpdata[key1][0],'>4d') + 'K,\t\tRes--$' + format(tmpdata[key1][1], '>4d') + 'K,\t\tCI--$' + format(tmpdata[key1][2], '>4d') + 'K\n'       
            app.fb4.AppendText(mset.logs)
            key1 +=1
        
        self.write1.close()
        mset.logs = dcf.mergestrs(dcf.show_end(mset.rlt))
        app.fb4.AppendText(mset.logs)
        mset.completed += 1
