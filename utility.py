# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 22:57:39 2020

@author: kia69
"""
import re
import xlwt as xt
from xlwt import default_style
from datetime import timedelta

workdays = ['fill','星期一','星期二','星期三','星期四','星期五']

'''
单元格格式
'''
bold = "font:bold True;"
centre = "alignment: horz centre, vert centre;"
border = "borders: left thin, right thin, top thin, bottom thin;"
font_color = "font: color red;"
background_fill = "pattern: pattern solid, fore_color ice_blue;"
class F():
    
    a = xt.easyxf(bold+centre+border)           #字体加粗，上下居中，外侧框线
    b = xt.easyxf(centre+border)           #上下居中，外侧框线
    c = xt.easyxf(font_color+border+centre)          #标红，上下居中，外侧框线
    d = xt.easyxf(bold+background_fill+centre+border)  #字体加粗，背景色填充，上下居中，外侧框线
    e = xt.easyxf(border)             #外侧框线 
    f = xt.easyxf(border)        #外侧框线
    bold_percent = xt.easyxf(bold+centre+border,num_format_str='0.00%')       
    normal_percent = xt.easyxf(centre+border,num_format_str='0.00%')
    two_decimal = xt.easyxf(centre+border,num_format_str='0.00')
    four_decimal = xt.easyxf(centre+border,num_format_str='0.0000')
    default = default_style

class T():
    A = 0
    B = 1
    C = 2
    D = 3
    E = 4
    F = 5
    G = 6
    H = 7
    I = 8
    J = 9
    K = 10
    L = 11
    M = 12
    N = 13
    O = 14
    P = 15
    Q = 16
    R = 17
    S = 18
    T = 19
    U = 20

def checkConfig(filels,config):      
    configlist = config.sections()
    notFound = set()
    for item in filels:
        if item.endswith('xls') or item.endswith('csv'):
            si = item.find('-') + 1
            if(item[si:-4] not in configlist):
                notFound.add(item[si:-4])
    assert len(notFound) == 0, '找不到以下券商的配置格式：'+str(notFound)

def prev_weekday(adate):
    adate -= timedelta(days=1)
    while adate.weekday() > 4: # Mon-Fri are 0-4
        adate -= timedelta(days=1)
    return adate

def isfloat(st):
    try:
        float(st)
        return True
    except:
        return False
    
def get_line(lines,row):
    line = lines[row]
    for n in range(10):
        line = re.sub(r'\t\t','\t \t',line)
    line = line.replace('\n','')
    line = line.replace('\t',',')
    line = line.replace('\'','')
    line = line.replace('\"','')
    line = line.replace('=','')
    line = re.sub(',,+',',',line)
    line = line.strip(',')
    return line.split(',')

def dongfangzhengquan_transaction(lines):
    n_line = [get_line(lines,i) for i in range(len(lines))]
    n_dict = [dict(zip(n_line[0],line)) for line in n_line[1:] if '合计' not in line]
    result = []
    result.append('证券代码,证券名称,成交数量,成交价格,成交金额,买卖标志')
    for item in n_dict:
        temp = [item['证券代码'],item['证券名称']]
        amount = int(item['买入数量']) - int(item['卖出数量'])
        if amount > 0:
            temp.extend([amount,item['买入均价'],float(item['买入成交金额'])-float(item['卖出成交金额']),'买入'])
        elif amount < 0:
            temp.extend([-amount,item['卖出均价'],float(item['卖出成交金额'])-float(item['买入成交金额']),'卖出'])
        else:
            continue
        temp = [str(xx) for xx in temp]
        result.append(','.join(temp))
    return result