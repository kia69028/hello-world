# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 20:43:17 2020

@author: xinjian.cao
"""
import re,json,os,datetime
from utility import get_line,isfloat,F,T,xt,workdays,default_style

class report:
    def __init__(self,name,date,config):
        self.name = name
        self.date_str = date
        self.date_dt = datetime.datetime.strptime(date,'%Y%m%d')
        self.config = config
        self.totalAsset = 0
        self.totalAvailable = 0
        self.prevTotalAsset = 0
        self.prevMonthNet = 0
        self.prevYearNet = 0
        self.shares = 0
        self.note = ' '
        self.totalIn = 0
        self.holdingList = []
        self.transactionList = [] #买卖成交分开
        self.mergedTransactionList = [] #合并所有交易（日内回转合并在一起）
        self.money_fund = []
        self.cash_management = []
        self.workBook = xt.Workbook(encoding='utf-8')
        self.workSpace = self.workBook.add_sheet(name,cell_overwrite_ok=True)
        self.filepath = '{}交易汇总\\'.format(date)
        self.workSpace.set_portrait(0)
        self.currentName = self.filepath+'{}每日交易汇总{}-自动.xls'.format(self.name,self.date_str)
        if not os.path.exists(self.filepath):
            os.mkdir(self.filepath)
    
    
    def readYesterday(self,alldata):
        row = alldata.loc[self.name]
        self.prevTotalAsset = row['资产']
        self.shares = row['份额']
        self.totalIn = row['总入金']
        self.prevMonthNet = row['上月末净值']
        self.prevYearNet = row['2019年末净值']
        self.note = row['备注']
    
    def read_holdings(self,raw_holding):
        config = self.config
        holdings_rejected = config.get('通用配置','不计入持仓').split(',')
        gushouTransform = config.get('通用配置','转债股手转换').split(',')
        dc = {}
        for brokerName,lines in raw_holding:
            
            #读取资产和可用资金
            row,col = config.get(brokerName,'总资产').split(',')
            self.totalAsset += float(get_line(lines,int(row) - 1)[ord(col) - ord('A')])
            row,col = config.get(brokerName,'可用资金').split(',')
            self.totalAvailable += float(get_line(lines,int(row)-1)[ord(col) - ord('A')])
            
            #开始读取持仓，并将相同的产品不同券商下的持仓相加在一起
            s_Row = int(config.get(brokerName,'持仓起始行数')) - 1
            headings = get_line(lines,s_Row)
            format_adj = json.loads(config.get(brokerName,'持仓格式'))
            global bond,money_fund
            for item in format_adj.keys():
                pos = headings.index(item)
                headings[pos] = format_adj[item]
            index = headings.index('证券代码')
            index_b = headings.index('证券名称')
            for i in range(s_Row+1,len(lines)):
                row = get_line(lines,i)
                row = [float(row[x]) if ((x != index and x != index_b) and isfloat(row[x])) else row[x] for x in range(len(row))]
                rowDict = dict(zip(headings,row))
                if(re.match('\d{6}',rowDict['证券代码']) is None or any(elem in rowDict['证券名称'] for elem in holdings_rejected)):
                    continue
                if(rowDict['证券代码'].startswith('511')):
                    self.money_fund.append(rowDict)
                    continue
                if brokerName in gushouTransform and rowDict['证券代码'].startswith('11'):
                    rowDict['证券数量'] = rowDict['证券数量'] * 10
                rowDict['动用资金'] = rowDict['证券数量'] * rowDict['成本价']
                item = rowDict
                if item['证券名称'] not in dc.keys():
                    dc[item['证券名称']] = item
                else:
                    dc[item['证券名称']]['证券数量'] += item['证券数量']
                    dc[item['证券名称']]['动用资金'] += item['动用资金']
                    dc[item['证券名称']]['最新市值'] += item['最新市值']                
        
        #计算合并后的每个股票持仓的平均持仓成本和平均当前价
        for rowDict in dc.values():
            rowDict['成本价'] = 0 if rowDict['证券数量'] == 0 else rowDict['动用资金']/rowDict['证券数量']
            rowDict['当前价'] = 0 if rowDict['证券数量'] == 0 else rowDict['最新市值']/rowDict['证券数量']
            rowDict['盈亏比例'] = (0 if rowDict['动用资金'] == 0 else rowDict['浮动盈亏']/(rowDict['动用资金']))
            self.holdingList.append(rowDict)
    
    def read_transactions(self,raw_transaction):
        config = self.config
        weituo = config.get('通用配置','委托成交记录')
        transaction_rejected = config.get('通用配置','不计入交易').split(',')
        gushouTransform = config.get('通用配置','转债股手转换').split(',')
        merged_dc = {}  #交易汇总，日内回转合并在一起，比如 某只股票 当日卖出5000，又买入3000，则汇总为卖出2000股
        templist = []
        for brokerName,lines in raw_transaction:
            s_Row = int(config.get(brokerName,'交易起始行数')) - 1
            try:
                executor = json.loads(config.get('交易员',self.name))
            except Exception as e:
                executor = {"下达人":"","执行人":""}
            headings = get_line(lines, s_Row)
            format_adj = json.loads(config.get(brokerName,'交易格式'))
            for item in format_adj.keys():
                pos = headings.index(item)
                headings[pos] = format_adj[item]
            index = headings.index('证券代码')
            for i in range(s_Row+1,len(lines)):
                row = get_line(lines,i)
                if brokerName in weituo:
                    if any('撤单' in elem for elem in row):
                        continue
                    if not any('成交' in elem for elem in row):
                        continue
                    
                row = [float(row[x]) if (x != index and isfloat(row[x])) else row[x] for x in range(len(row))]
                rowDict = dict(zip(headings,row))
                if any(elem in rowDict['证券名称'] for elem in transaction_rejected):
                    continue
                if(rowDict['证券名称'].startswith("GC")):
                    self.cash_management.append(rowDict)
                    continue
                if brokerName in gushouTransform and rowDict['证券代码'].startswith('11'):
                    rowDict['成交数量'] = rowDict['成交数量'] * 10
                rowDict['下达人'] = executor['下达人']
                rowDict['执行人'] = rowDict['执行人']
                templist.append(rowDict)
            
                item = rowDict
                sign = 1 if ('买入' in item['买卖标志']) else -1
                if item['证券名称'] not in merged_dc.keys():
                    item['证券名称']['成交数量'] = sign*item['成交数量']
                    merged_dc[item['证券名称']] = item
                else:
                    merged_dc[item['证券名称']]['成交数量'] += sign*item['成交数量']
        
        for direction in ['买入','卖出']:
            ls = [x for x in templist if (direction in x['买卖标志'])]
            dc = {}
            for item in ls:
                if item['证券名称'] not in dc.keys():
                    dc[item['证券名称']] = item
                else:
                    dc[item['证券名称']]['成交数量'] += item['成交数量']
                    dc[item['证券名称']]['成交金额'] += item['成交数量']*item['成交价格']            
            for item in dc.values():
                item['成交价格'] = item['成交金额'] / item['成交数量']
                item['买买标志'] = direction
                self.transactionList.append(item)
                
        for rowDict in merged_dc.values():
            if rowDict['成交数量'] == 0:
                continue
            rowDict['买卖标志'] = '买入' if rowDict['成交数量'] > 0 else '卖出'
            rowDict['成交数量'] = abs(rowDict['成交数量'])
            self.mergedTransactionList.append(rowDict)    
        
    def read_revoke(self,raw_revoke):
        for brokerName,lines in raw_revoke:
            headings = get_line(lines,0)
            format_adj = json.loads(self.config.get(brokerName,'撤单格式'))
            for item in format_adj.keys():
                pos = headings.index(item)
                headings[pos] = format_adj[item]
            for i in range(1,len(lines)):
                row = get_line(lines,i)
                row = [float(item) if isfloat(item) else item for item in row]
                rowDict = dict(zip(headings,row))
                if '买入' in rowDict['买卖标志']:
                    self.totalAvailable += (rowDict['委托数量']-rowDict['成交数量'])*rowDict['委托价格']
