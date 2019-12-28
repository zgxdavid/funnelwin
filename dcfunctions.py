#coding=utf-8

import wx
import sys
#import os

def prt_head():
    # 输出Python和wxPython的版本，调试用#############################################
    print("Python %s" % sys.version)
    print("wx.version: %s" % wx.version())
    # print("OS current pid is: %s" % os.getpid()); input("Press Enter to go ahead...")
    # ##########################################################################

def show_welcome0():
    tstr1 = "========++++++++++++++++++++++++++++++++++++++++++++++++========"
    tstr2 = "******** 張國祥  David Chang, David.Chang2@Dell.com ********"
    tstr3 = "******** HW SFDC Funnel/Funnel List生成工具 - Version 1.0 *********"
    tstr = [tstr1, tstr2, tstr3, tstr1]
    return tstr

def show_welcome(src1='', src2='', rlt=''):
    tstr4 = "\tSource file of HW SFDC Funnel from Yu Honghong:\n\t\'" + src1 + "\'"
    tstr5 = "\tSource file of Funnel List from David Chang: \n\t\'" + src2 + "\'"
    tstr6 = "\tGenerated Funnel List in:\n\t\'" + rlt + "\'"
    tstr = [tstr4, tstr5, tstr6]
    return tstr

def print_start_search(sheetname=''):
    tstr1 = "******** Start searching " + sheetname + " ********"
    tstr = [tstr1]
    return tstr

def print_end_search(sheetname=''):
    tstr1 = "******** Search " + sheetname + " end ********"
    tstr = [tstr1]
    return tstr

def print_start_drop(lines=0, sheetname=''):
    tstr1 = "******** Start dropping " + str(lines) + " lines from " + sheetname + " ********"
    tstr = [tstr1]
    return tstr

def print_end_drop(lines=0, sheetname=''):
    tstr1 = "******** End drop " + str(lines) + " lines from " + sheetname + " ********"
    tstr = [tstr1]
    return tstr

def show_end(rlt=''):
    tstr1 = "========++++++++++++++++++++++++++++++++++++++++++++++++========"
    tstr2 = "******** 程序执行结束，已经生成最终文件，放在 \'Funnel List\'表中："
    tstr3 = "******** 文件路径：\' " + rlt + "\'"
    tstr = [tstr1, tstr2, tstr3, tstr1]
    return tstr

def type_src(lines_hw=0, cols_hw=0, lines_zgx=0, cols_zgx=0):
    tstr1 = "\tSource \'HW SFDC Funnel\' lines from Yu Honghong are: " + str(lines_hw)
    tstr2 = "\tSource \'HW SFDC Funnel\' columns from Yu Honghong are: " + str(cols_hw)
    tstr3 = "\tSource lines of \'Funnel List\' from David Chang are: " + str(lines_zgx)
    tstr4 = "\tSource columns of \'Funnel List\' from David Chang are: " + str(cols_zgx)
    tstr = [tstr1, tstr2, tstr3, tstr4]
    return tstr

def prt_fordel(fordel=[], sheetname=''):
    tmp1 = len(fordel)
    tstr1 = "Total " + str(tmp1) + " lines in the table " + sheetname + " will be deleted: "
    tstr = [tstr1]
    tstrtmp =''
    for k in range(tmp1):
        if (k+1) % 10 != 0:
            tstrtmp += str(fordel[k] + 2) + '\t'
        else:
            tstr.append(tstrtmp)
            tstrtmp = ''
    return tstr

def sheetbackup(write2, df, sheetname='tmp1'):
    if sheetname != 'tmp1':
        df.to_excel(write2, sheet_name=sheetname, index=None)
        write2.save()
        return True
    else: return False        
    
def sheetlast(write1, df, sheetname='tmp1'):
    if sheetname != 'tmp1':
        df.to_excel(write1, sheet_name=sheetname, index=None)
        write1.save()
        return True
    else: return False
    
def droplns(write1, df, fordel=[], sheetname='unknown'):
    if sheetname != 'unknown':
        dfout1 = df.drop(labels = [v for v in fordel])
        dfout1.to_excel(write1, sheet_name=sheetname, index=None)
        write1.save()
        return dfout1
    else:
        dfout1 = df
        return dfout1

def mergestrs(strs=[]):
    if len(strs) == 0:
        tmpstr = ''
        return tmpstr
    else:
        tmpstr = ''
        for x in range(len(strs)):
            tmpstr += strs[x] + '\n'
        return tmpstr

def numtorate(number=0.0):
    return '{:.0f}%'.format(number*100)

def ratetonum(rate='90%'):
    return float(rate[0:-1])/100.00