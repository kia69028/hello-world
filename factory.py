# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 20:15:01 2020

@author: xinjian.cao
"""

from producer import baseProducer
from utility import isfloat,F,T
import pandas as pd

stock_name_list = ['主题','锐进','长嬴','指数增强']
bond_name_list = ['红包','红宝石','可转债']
neutral_name_list = ['合益','稳健']


class stockProducer(baseProducer):
    def __init__(self,report):
        super(stockProducer,self).__init__(report)
    
    def write_transactions(self,startRow):
        if self.report.name in ['主题2号','主题3号']:
            return super(stockProducer,self).write_transaction(startRow)
        else:
            self.localRowNum = startRow
            self.rowNumDict['transaction'] = startRow
            self.write_merge(0, 11, '三、今日交易', F.d)
            self.write_merge(0, 11, '(1)证券交易', F.e)
            self.write_transaction_direction()
            self.rowNumDict['transaction_end'] = self.localRowNum
            return self.localRowNum
    
    def cal_position(self):
        self.securityPos = self.holdingSum['holds'] / self.totalAsset
        
    def write_header(self,startRow):
        rowResult = super(stockProducer,self).write_header(startRow)
        if ('指数增强' in self.report.name):
            self.write_merge(0,2,'',F.b,self.rowNumDict['header']+5)
            self.write_merge(3,5,'',self.rowNumDict['header']+5)
            self.write_merge(0,2,'',self.rowNumDict['header']+6)
            self.write_merge(3,5,'',self.rowNumDict['header']+6)  
        return rowResult


class bondProducer(baseProducer):
    def __init__(self,report):
        super(bondProducer,self).__init__(report)
    
    def cal_position(self):
        self.bondPos = self.holdingSum['holds'] / self.totalAsset
        
    def write_header(self,startRow):
        rowResult = super(bondProducer,self).write_header(startRow)
        self.write_merge(0,2,'可购买可转债资金（万）80%仓位',F.b,self.rowNumDict['header']+4)
        val = (0.8 - self.holdingSum['used']/self.totalAsset)*self.totalAsset/10000 if self.net > 1 else (0.8 - self.holdingSum['holds']/self.totalAsset)*self.totalAsset/10000
        self.write_merge(3,5,val,F.two_decimal,self.rowNumDict['header']+4)
        self.write_merge(0,2,'理财资金（万）',F.b,self.rowNumDict['header']+5)
        self.write_merge(3,5,0,self.rowNumDict['header']+5)
        return rowResult
    
    
class neutralProducer(baseProducer):
    def __init__(self,report):
        super(neutralProducer,self).__init__(report)
        self.exposure = 0 
        self.futurelist = []
        self.futureSum = {'数量':0,'保证金':0,'占比':0,'浮动盈亏':0,'盈亏比例':0}
        self.futures = None
        self.count = 0
        self.read_futures()
        
    def read_futures(self):
        for x in neutral_name_list:
            name = x if x in self.report.name else 'None'        
        futures = pd.read_excel(self.report.date_str+name+'股指期货.xlsx')
        self.futures = futures
        for i in range(len(futures)):
            if isfloat(futures.iloc[i,1]):
                break
            if futures.iloc[i,1].startswith("IF") or futures.iloc[i,1].startswith("IC"):
                self.futurelist.append({'证券代码':futures.iloc[i,1],'证券名称':futures.iloc[i,2],'开仓均价':futures.iloc[i,T.G],'数量':futures.iloc[i,T.E],
                                   '方向':futures.iloc[i,T.F],'保证金':futures.iloc[i,T.K],'占比':futures.iloc[i,T.K]/futures.iloc[0,0],
                                   '收盘价':futures.iloc[i,T.I],'浮动盈亏':futures.iloc[i,T.J],'盈亏比例':futures.iloc[i,T.J]/futures.iloc[0,0]})
                self.count += 1
        self.future_hold = futures.iloc[0,0] / 10000
        self.future_diff = futures.iloc[count+3,9]
        
    def write_holdings(self,startRow):
        self.localRowNum = startRow
        self.rowNumDict['holding'] = startRow
        self.write_merge(0,11,'二、盘后持仓概览',F.d)
        self._write_future_holding()
        self.write_merge(0,11,'(2)股票持仓',F.d)
        self.write_stock_holding()
        self.rowNumDict['holding_end'] = self.localRowNum
        return self.localRowNum
    
    def write_header(self,startRow):
        rowResult = super(bondProducer,self).write_header(startRow)
        self.write_merge(6,8,'仓位',F.b,self.rowNumDict['header']+1)
        self.write_merge(9,11,(self.holdingSum['holds'] + self.futures.iloc[0,0]) / self.totalAsset, F.bold_percent,self.rowNumDict['header']+1)
        self.write_merge(0,2,'期货权益（万）',F.b,self.rowNumDict['header']+5)
        self.write_merge(3,5,self.futures.iloc[0,0]/10000,F.two_decimal,self.rowNumDict['header']+5)
        self.write_merge(0,2,'基差（现货-期货）',F.b,self.rowNumDict['header']+6)
        self.write_merge(3,5,self.futures.iloc[self.count+3,9],F.two_decimal,self.rowNumDict['header']+6)        
        return rowResult
        
    def _write_future_holding(self):
        mergelist = [(0,),(1,),(2,),(3,),(4,),(5,),(6,),(7,),(8,),(9,10),(11,)]
        self.write_merge(0, 11, '(1)股指期货', F.d)
        self.write_merge_row(mergelist, 
                                 ['序号','证券代码','证券名称','开仓均价','数量','方向','保证金','占比','收盘价','浮动盈亏（期货账户）','盈亏比例'], 
                                 [F.b for x in range(11)])
        for i,item in enumerate(self.futurelist,1):
            self.write_merge_row(mergelist,[i,item['证券代码'],item['证券名称'],item['开仓均价'],item['数量'],item['方向'],item['保证金'],item['占比'],item['收盘价'],
                                            item['浮动盈亏'],item['盈亏比例']],
                                 [F.b,F.b,F.b,F.b,F.b,F.b,F.b,F.normal_percent,F.b,F.c,F.normal_percent])
        
        self.write_merge_row([(0,3),(4,),(5,),(6,),(7,),(8,),(9,10),(11,)],
                                 ['合计',self.futureSum['数量'],'',self.futureSum['保证金'],self.futureSum['占比'],'',self.futureSum['浮动盈亏'],self.futureSum['盈亏比例']],
                                 [F.a,F.b,F.b,F.b,F.bold_percent,F.b,F.c,F.normal_percent])
    
    def cal_totalAsset(self):
        super(bondProducer,self).cal_totalAsset()
        self.totalAsset += self.futures.iloc[0,0]
        if '合益' in self.report.name:
            datediff = (datetime.today().date() - dtt.date(2020,4,16)).days
            for n in range(len(self.futures)):
                if self.futures.iloc[n,0] == '入金':
                    break
            n += 1        
            current_hold = self.futures.iloc[n,1]
            self.totalAsset += current_hold-175664.31-self.mf['holds']-4000*datediff+(1.535-self.futures.iloc[self.count+2,2])*mf['holds']/1.535
    
    def cal_totalAvailable(self):
        super(bondProducer,self).cal_totalAvailable()
        if '合益' in self.report.name:
            self.totalAvailable -= self.mf['holds']
        
    def cal_position(self):
        self.securityPos = self.holdingSum['holds'] / self.totalAsset
        for item in self.futurelist:
            for key in self.futureSum.keys():
                self.futureSum[key] += item[key]
    
            
def sheetFactory(report):
    if any(elem in report.name for elem in stock_name_list):
        return stockProducer(report)
    if any(elem in report.name for elem in bond_name_list):
        return bondProducer(report)
    if any(elem in report.name for elem in neutral_name_list):
        return neutralProducer(report)
    
    