# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 18:37:37 2020

@author: xinjian.cao
"""
import configparser
import pandas as pd
from utility import get_line,dongfangzhengquan_transaction,checkConfig,prev_weekday
from factory import sheetFactory
import os
from datetime import datetime
from report import report

def helper(productName,args,config,alldata):
    revoke = hold = transaction = []
    date = None
    for item in args:
        si = item.find('-') + 1
        if date:
            assert item[:8] == date, print('{}有两个不同日期的文件'.format(productName))
        if item.endswith('csv'):
            csv_data = pd.read_csv(item,'r')
            lines = [csv_data.iloc[i][0] for i in range(csv_data.shape[0])]
        else:
            with open(item,'r') as file:
                lines = file.readlines()
        if '撤单' in item:
            revoke.append((item[si:-4],lines))
        if '持仓' in item:
            hold.append((item[si:-4],lines))
        if '成交' in item:
            if(item[si:-4] == '东方证券'):
                transaction.append((item[si:-4],dongfangzhengquan_transaction(lines)))
                continue
            transaction.append((item[si:-4],lines))
        date = item[:8]
        
    report = report(productName,date,config)
    report.read_holdings(hold)
    report.read_transactions(transaction)
    report.read_revokes(revoke)
    producer = sheetFactory(report)
    producer.cal_params()
    row = producer.write_title(0)
    row = producer.write_header(row)
    row = producer.write_holdings(row)
    row = producer.write_transactions(row)
    producer.write_cash(row)
    producer.save()
        
def main():
    config = configparser.ConfigParser()
    filename = '配置文件\格式.ini'
    config.read(filename,encoding='utf-8')
    filels = os.listdir()
    checkConfig(filels,config)
    
    preLs = {}
    for item in filels:
        if item.endswith('xls'):
            si = item.find('-')-2
            preLs[item[8:si]] = []
          #  copyfile(item,item.replace('xls','txt'))
        elif item.endswith('csv'):
            si = item.find('-')-2
            preLs[item[8:si]] = []
            
    filels = os.listdir()
    for item in filels:
        if item.endswith('xls') or item.endswith('csv'):
            si = item.find('-') - 2
            preLs[item[8:si]].append(item)
            date = item[:8]
    
    yesterday = prev_weekday(datetime.strptime(date,'%Y%m%d'))
    yesterday = datetime.strftime(yesterday,'%Y%m%d')
    alldata = pd.read_csv('配置文件\{}产品数据.csv'.format(yesterday),encoding='gb18030')
    alldata = alldata.set_index('产品')
    finished = 0
    for key in preLs.keys():
        alldata = helper(key,preLs[key],alldata,config)
        finished += 1
    
    print('成功生成交易汇总：{}/{}'.format(finished,len(preLs.keys())))
    for item in filels:
        if item.endswith('txt'):
            os.remove(item)
    
    alldata.to_csv(r'配置文件\{}产品数据.csv'.format(date),encoding='gb18030')
    
if __name__ == '__main__':
    main()