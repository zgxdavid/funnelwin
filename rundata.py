#coding=utf-8
#这是主程序，可以运行 - Vesion 1.0

import pandas as pd
import math
import dcfunctions as dcf
import xlrd

class RunData():
    def __init__(self, mset, write1, write2):
        self.write1 = write1
        self.write2 = write2
        self.fordel = []

    def run_data(self, app, mset):
        mset.logs = dcf.mergestrs(dcf.show_welcome(mset.src1, mset.src2, mset.rlt))
        app.fb4.AppendText(mset.logs)
        mset.logs = "####Read source HW Funnel & ZGX Funnel into two DataFrame Objects\n"
        app.fb4.AppendText(mset.logs)
        din1 = pd.read_excel(mset.src1, sheet_name=mset.src1_sheet)
        din2 = pd.read_excel(mset.src2, sheet_name=mset.src2_sheet)
        lines_hw_funnel = din1.shape[0]
        columns_hw_funnel = din1.shape[1]
        lines_zgx_funnel = din2.shape[0]
        columns_zgx_funnel = din2.shape[1]
        mset.logs = dcf.mergestrs(dcf.type_src(lines_hw_funnel, columns_hw_funnel, lines_zgx_funnel, columns_zgx_funnel))
        app.fb4.AppendText(mset.logs)
    #################################################
        mset.logs = dcf.mergestrs(dcf.print_start_search(mset.src1_sheet))
        app.fb4.AppendText(mset.logs)
        
        for i in range(lines_zgx_funnel):
            if i not in self.fordel:
                if (din2.iloc[i,1] not in mset.dcsp_ntname_checking) or (din2.iloc[i,11] in mset.for_drop):
                    self.fordel.append(i)
                else:                
                    for j in range(lines_hw_funnel):
                        if (str(din2.iloc[i,3]).strip() == str(din1.iloc[j,28]).strip()) and (str(din1.iloc[j,51]).strip() in mset.win_rate_check_drop):
                            self.fordel.append(i)
    
        mset.logs = dcf.mergestrs(dcf.prt_fordel(self.fordel, mset.src2_sheet))
        app.fb4.AppendText(mset.logs)
        mset.logs = dcf.mergestrs(dcf.print_end_search(mset.src2_sheet))
        app.fb4.AppendText(mset.logs)
    #################################################        
        mset.logs = dcf.mergestrs(dcf.print_start_drop(len(self.fordel), mset.src2_sheet))
        app.fb4.AppendText(mset.logs)
         
        sheetname0 = 'sheet_dout0'
        dcf.sheetbackup(self.write2, din2, mset.src2_sheet)
        dout0 = dcf.droplns(self.write1, din2, self.fordel, sheetname0)
        dcf.sheetbackup(self.write2, dout0, sheetname0)
           
        mset.logs = dcf.mergestrs(dcf.print_end_drop(len(self.fordel), mset.src2_sheet))
        app.fb4.AppendText(mset.logs)
    #################################################
        app.fb4.AppendText("####开始整理原始硬件的Funnel，先将别人的，Rev为0的，非本Q的，Won or Lost的单子，以及Sales认为不靠谱的单子\n")      
        mset.logs = dcf.mergestrs(dcf.print_start_search(mset.src1_sheet))
        app.fb4.AppendText(mset.logs)
            
        self.fordel.clear()
        for l in range(lines_hw_funnel):
            if str(din1.iloc[l,13]).strip() not in mset.dcsp_name_checking:
                self.fordel.append(l)
            elif str(din1.iloc[l,47]).strip() not in mset.qtr_working:
                self.fordel.append(l)
            elif str(din1.iloc[l,51]).strip() not in mset.win_rate_check_useful:
                self.fordel.append(l)
            elif str(din1.iloc[l,53]).strip() not in mset.sales_forecast_useful:
                self.fordel.append(l)
            elif float(din1.iloc[l,64]) < 1.00:
                self.fordel.append(l)
            else:
                continue
             
        mset.logs = dcf.mergestrs(dcf.prt_fordel(self.fordel, mset.src1_sheet))
        app.fb4.AppendText(mset.logs)
        mset.logs = dcf.mergestrs(dcf.print_end_search(mset.src1_sheet))
        app.fb4.AppendText(mset.logs)
    #################################################
        mset.logs = dcf.mergestrs(dcf.print_start_drop(len(self.fordel), mset.src1_sheet))
        app.fb4.AppendText(mset.logs)
     
        sheetname1 = 'sheet_dout1'
    #   dcf.sheetbackup(self.write2, din1, mset.src1_sheet) #这个文件太大，写的时候会将文件锁死，后面打不开文件，就先不备份了
        dout1 = dcf.droplns(self.write1, din1, self.fordel, sheetname1)
        dcf.sheetbackup(self.write2, dout1, sheetname1)
             
        mset.logs = dcf.mergestrs(dcf.print_end_drop(len(self.fordel), mset.src1_sheet))
        app.fb4.AppendText(mset.logs)
    #################################################
        app.fb4.AppendText("####开始处理第一轮来自于HW SFDC Funnel的dout1\n")
        app.fb4.AppendText("####将dout1中和ZGX Funnel List相对应的列进行数据修改，将会生成dout2没有删除的动作\n")
        mset.logs = dcf.mergestrs(dcf.print_start_search(sheetname1))
        app.fb4.AppendText(mset.logs)
        app.fb4.AppendText("******** 找出dout1中的有效列，生成dout2，并修改 ********\n")
         
        tocols0 = []
        tocols = []
        for n in dout1.columns: tocols0.append(n)
        for p in mset.hw_to_zgx_list: tocols.append(tocols0[p])
     
        dout2 = pd.DataFrame(dout1, columns=tocols)
        dout2.columns = dout0.columns
         
        lentmp = dout2.shape[0]
        for q in range(lentmp):
            dout2.iloc[q,0] = 'DCSP'
            dout2.iloc[q,1] = mset.dcsp_ntname_checking[0]
            dout2.iloc[q,21] = mset.qtr_working1
            qtmp1 = str(dout2.iloc[q,6]).split()
            if len(qtmp1) == 4: qtmp2 = qtmp1[0] + ' ' + qtmp1[1] + ' ' + qtmp1[2]
            elif len(qtmp1) == 3: qtmp2 = qtmp1[0] + ' ' + qtmp1[1]
            else: qtmp2 = 'NA'
            dout2.iloc[q,6] = qtmp2
            dout2.iloc[q,10] = mset.qtr_working1 + 'WK{:02d}'.format(int(dout2.iloc[q,10]) - 202039)
            qtmp1 = str(dout2.iloc[q,11]).split()
            dout2.iloc[q,11] = str(qtmp1[-1])
            if str(dout2.iloc[q,13]).strip() not in mset.svc_types:
                dout2.iloc[q,12] = math.ceil(dout2.iloc[q,12]/10000)
                dout2.iloc[q,13] = 'Product Attached'
                dout2.iloc[q,14] = 'New Services'
            else: dout2.iloc[q,12] = math.ceil(dout2.iloc[q,12]/1000)
        app.fb4.AppendText("####没有删除动作，只是将dout1生成的和dout0一样列数和内容的视图dout2写回\n")    
        sheetname2 = 'sheet_dout2'    
        dcf.sheetbackup(self.write2, dout2, sheetname2)
        dcf.sheetlast(self.write1, dout2, sheetname2)
         
        mset.logs = dcf.mergestrs(dcf.print_end_search(sheetname1))
        app.fb4.AppendText(mset.logs)
    #################################################
        app.fb4.AppendText("####再对生成的dout2 from dout1进行整理，相加去重，生成dout3\n")
        mset.logs = dcf.mergestrs(dcf.print_start_search(sheetname2))
        app.fb4.AppendText(mset.logs)
            
        self.fordel.clear()
        for r in range(lentmp):
            if r not in self.fordel:
                for s in range(r+1, lentmp):
                    if dout2.iloc[r,3] == dout2.iloc[s,3] and dout2.iloc[r,13] == dout2.iloc[s,13] and (s not in self.fordel):
                        self.fordel.append(s)
                        dout2.iloc[r,12] += dout2.iloc[s,12]
     
        app.fb4.AppendText("These lines will be summed firstly! \n") 
        mset.logs = dcf.mergestrs(dcf.prt_fordel(self.fordel, sheetname2))
        app.fb4.AppendText(mset.logs)
        mset.logs = dcf.mergestrs(dcf.print_end_search(sheetname2))
        app.fb4.AppendText(mset.logs)
    #################################################
        mset.logs = dcf.mergestrs(dcf.print_start_drop(len(self.fordel), sheetname2))
        app.fb4.AppendText(mset.logs)
         
        sheetname3 = 'sheet_out3'
        douttmp2 = pd.read_excel(mset.rlt, sheet_name=sheetname2)
        dout3 = dcf.droplns(self.write1, douttmp2, self.fordel, sheetname3)
        dcf.sheetbackup(self.write2, dout3, sheetname3)
         
        mset.logs = dcf.mergestrs(dcf.print_end_drop(len(self.fordel), sheetname2))
        app.fb4.AppendText(mset.logs)
    #################################################    
        app.fb4.AppendText("####对dout3中小于30K服务的单子进行删除，生成dout4\n")   
        mset.logs = dcf.mergestrs(dcf.print_start_search(sheetname3))
        app.fb4.AppendText(mset.logs)
         
        sheetname4 = 'sheet_out4'
        self.fordel.clear()
        lentmp = dout3.shape[0]
        for y in range(lentmp):
            if int(dout3.iloc[y,12]) < 30 and str(dout3.iloc[y,13]).strip() == 'Product Attached':
                self.fordel.append(y)
         
        mset.logs = dcf.mergestrs(dcf.prt_fordel(self.fordel, sheetname3))
        app.fb4.AppendText(mset.logs)
        mset.logs = dcf.mergestrs(dcf.print_end_search(sheetname3))
        app.fb4.AppendText(mset.logs)
    #################################################
        app.fb4.AppendText("####开始对dout3中小于30K服务单子删除，生成dout4\n")
        mset.logs = dcf.mergestrs(dcf.print_start_drop(len(self.fordel), sheetname3))
        app.fb4.AppendText(mset.logs)
         
        douttmp3 = pd.read_excel(mset.rlt, sheet_name=sheetname3)
        dout4 = dcf.droplns(self.write1, douttmp3, self.fordel, sheetname4)
        dcf.sheetbackup(self.write2, dout4, sheetname4)
         
        mset.logs = dcf.mergestrs(dcf.print_end_drop(len(self.fordel), sheetname3))
        app.fb4.AppendText(mset.logs)
    #################################################    
        app.fb4.AppendText("####开始将dout4添加到初步清理过的ZGX Funnel List的dout0后面，生成dout5\n")
        dout5 = dout0.append(dout4, ignore_index=True)
        sheetname5 = 'sheet_out5'
        dcf.sheetbackup(self.write2, dout5, sheetname5)    
        dcf.sheetlast(self.write1, dout5, sheetname5)
    #################################################
        app.fb4.AppendText("####对dout5再做一次去重，生成dout6\n")
        mset.logs = dcf.mergestrs(dcf.print_start_search(sheetname5))
        app.fb4.AppendText(mset.logs)
        
        sheetname6 = 'sheet_out6'
        self.fordel.clear()
        douttmp5 = pd.read_excel(mset.rlt, sheet_name=sheetname5)
        lentmp = douttmp5.shape[0]
        for z1 in range(lentmp):
            if z1 not in self.fordel:
                for z2 in range(z1+1, lentmp):
                    if douttmp5.iloc[z2,2] == douttmp5.iloc[z1,2] and douttmp5.iloc[z2,3] == douttmp5.iloc[z1,3]:
                        if str(douttmp5.iloc[z1,15]).strip() != '' and str(douttmp5.iloc[z1,13]).strip() == str(douttmp5.iloc[z2,13]).strip() and str(douttmp5.iloc[z1,13]).strip() == 'Product Attached':
                            douttmp5.iloc[z1,12] = max(douttmp5.iloc[z1,12], douttmp5.iloc[z2,12]) 
                            dlwktmp1 = str(douttmp5.iloc[z1,10])
                            dlwktmp2 = str(douttmp5.iloc[z2,10])
                            if dlwktmp1[0:5] == mset.qtr_working1 and dlwktmp2[0:5] == mset.qtr_working1:
                                douttmp5.iloc[z1,10] = mset.qtr_working1 + 'WK{:02d}'.format(max(int(dlwktmp1[-2:]), int(dlwktmp2[-2:])))
                            winratetmp1 = str(douttmp5.iloc[z1,11])
                            winratetmp2 = str(douttmp5.iloc[z2,11])
                            if winratetmp1[-1] == '%':
                                winratetmp1 = winratetmp1[0:-1]
                                winratetmp1update = float(winratetmp1)/100
                            elif winratetmp1 != 'Dled' or winratetmp1 != 'Lost':
                                winratetmp1update = float(winratetmp1)
                            else:
                                winratetmp1update = 0.0
                            if winratetmp2[-1] == '%':
                                winratetmp2 = winratetmp2[0:-1]
                                winratetmp2update = float(winratetmp2)/100
                            elif winratetmp2 != 'Dled' or winratetmp2 != 'Lost':
                                winratetmp2update = float(winratetmp2)
                            else:
                                winratetmp2update = 0.0
                            douttmp5.iloc[z1,11] = dcf.numtorate(max(winratetmp1update, winratetmp2update))
                            self.fordel.append(z2)
                        if str(douttmp5.iloc[z2,13]) == str(douttmp5.iloc[z1,13]) and str(douttmp5.iloc[z2,13]) == 'CI Deployment':
                            self.fordel.append(z2)
                        if str(douttmp5.iloc[z2,13]) == str(douttmp5.iloc[z1,13]) and str(douttmp5.iloc[z2,13]) == 'Residency' and str(douttmp5.iloc[z2,14]) == 'Residency':
                            self.fordel.append(z2)
        self.fordel.sort()
        mset.logs = dcf.mergestrs(dcf.prt_fordel(self.fordel, sheetname5))
        app.fb4.AppendText(mset.logs)    
        mset.logs = dcf.mergestrs(dcf.print_end_search(sheetname5))
        app.fb4.AppendText(mset.logs)
    #################################################
        app.fb4.AppendText("####开始对dout5中的数据整理和删除，生成dout6\n")
        mset.logs = dcf.mergestrs(dcf.print_start_drop(len(self.fordel), sheetname5))
        app.fb4.AppendText(mset.logs)
        dout6 = dcf.droplns(self.write1, douttmp5, self.fordel, sheetname6)
        dcf.sheetbackup(self.write2, dout6, sheetname6)
        mset.logs = dcf.mergestrs(dcf.print_end_drop(len(self.fordel), sheetname5))
        app.fb4.AppendText(mset.logs)
    #################################################
        app.fb4.AppendText("####开始对dout6的数据进行修整\n")
        sheetname7 = 'sheet_out7'
        self.fordel.clear()
        douttmp6 = pd.read_excel(mset.rlt, sheet_name=sheetname6)
        lentmp = douttmp6.shape[0]
        for ss1 in range(lentmp):
            douttmp6.iloc[ss1,12] = int(round(float(douttmp6.iloc[ss1,12])))
            if '%' not in str(douttmp6.iloc[ss1,11]):
                douttmp6.iloc[ss1,11] = dcf.numtorate(float(douttmp6.iloc[ss1,11]))
            if str(douttmp6.iloc[ss1,14]).strip() == 'Product Attached' or str(douttmp6.iloc[ss1,14]).strip() == 'Residency':
                douttmp6.iloc[ss1,14] = 'New Services'
            if str(douttmp6.iloc[ss1,15]).strip() == 'nan':
                douttmp6.iloc[ss1,15] = '确认项目机会'
            if str(douttmp6.iloc[ss1,16]).strip() == 'nan':
                douttmp6.iloc[ss1,16] = 'No'
            if str(douttmp6.iloc[ss1,18]).strip() == 'nan':
                douttmp6.iloc[ss1,18] = 'Align with AE/ASE/SSE'
            if str(douttmp6.iloc[ss1,19]).strip() == 'nan':
                douttmp6.iloc[ss1,19] = 'HW SFDC Funnel'
            douttmp6.iloc[ss1,20] = int(dcf.ratetonum(str(douttmp6.iloc[ss1,11]))*float(douttmp6.iloc[ss1,12]))
            for ss2 in range(ss1+1, lentmp):
                if douttmp6.iloc[ss1,2] == douttmp6.iloc[ss2,2] and str(douttmp6.iloc[ss2,7]).strip() == 'nan':
                    douttmp6.iloc[ss2,7] = douttmp6.iloc[ss1,7]
                    douttmp6.iloc[ss2,8] = douttmp6.iloc[ss1,8]
        app.fb4.AppendText("####对dout6的数据修整结束，修改好的dout6存到dout7中\n")           
        dcf.sheetlast(self.write1, douttmp6, sheetname7)
        dcf.sheetbackup(self.write2, douttmp6, sheetname7)               
    #################################################    
        self.write1.close()
        self.write2.close()
        mset.logs = dcf.mergestrs(dcf.show_end(mset.rlt))
        app.fb4.AppendText(mset.logs)
        mset.completed += 1
