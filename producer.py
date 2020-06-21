# -*- coding: utf-8 -*-
"""
Created on Mon Jun 22 01:52:59 2020

@author: kia69
"""


from report import report
from utility import F,T,workdays,default_style

class baseProducer:
    def __init__(self,report):
        self.report = report
        self.localRowNum = 0
        self.rowNumDict = {}
        self.totalAsset = 0
        self.totalAvailable = 0
        self.securityPos = 0
        self.bondPos = 0
        self.transactionSum = {'动用':0,'释放':0}
        self.workSpace = report.workSpace
        self.holdingSum = {'used':0,'holds':0,'profit':0,'profit_pct':0}
        
    def write_merge_row(self,mergelist,contentlist,stylelist,rowNum = None):
        rowNum = self.localRowNum if rowNum is None else rowNum
        assert (len(mergelist) == len(contentlist) and len(contentlist) == len(stylelist)), 'length of mergelist/contentlist/stylelist does not match!'
        for i in range(len(mergelist)):
            if(len(mergelist[i]) < 2):
                self.workSpace.write(rowNum,mergelist[i][0],contentlist[i],stylelist[i])
            else:
                self.workSpace.write_merge(rowNum,rowNum,mergelist[i][0],mergelist[i][1],contentlist[i],stylelist[i])
        self.localRowNum += 1
    
    def write_merge(self,beg,end,content,style = default_style,rowNum = None):
        rowNum = self.localRowNum if rowNum is None else rowNum
        self.workSpace.write_merge(rowNum,rowNum,beg,end,style)
        self.localRowNum += 1
        
    def write_title(self,startRow):
        self.rowNumDict['title'] = startRow
        self.localRowNum = startRow
        self.write_merge_row([(0,2),(3,5),(6,7),(8,9),(10,11)], 
                ['名称',self.report.name+'私募基金','日期',self.report.date_dt.strftime("%Y/%m/%d"),
                 workdays[int(self.report.date_dt.strftime('%w'))]], 
                [F.a for x in range(5)])
        return self.localRowNum
    
    def write_header(self,startRow):
        mergelist = [(0,2),(3,5),(6,8),(9,11)]
        half_mergelist = [(6,8),(9,11)]
        self.localRowNum = startRow 
        self.rowNumDict['header'] = startRow
        self.write_merge(0,11,'一、盘后账户情况',F.d)
        self.write_merge_row(mergelist,['总资产',self.totalAsset,'总仓位（权益+可转债）',self.securityPos+self.bondPos],[F.b,F.b,F.b,F.normal_percent])
        self.write_merge_row(mergelist,['可用资金（含逆回购',self.totalAvailable,'权益类仓位',self.securityPos],[F.b,F.b,F.b,F.normal_percent])
        self.write_merge_row(mergelist,['总盈亏',self.totalAsset - 10000*self.totalIn,'可转债仓位',self.bondPos], [F.b,F.b,F.b,F.normal_percent])
        self.write_merge_row(half_mergelist,['净值',self.totalAsset/self.report.shares],[F.b,F.four_decimal])
        self.write_merge_row(half_mergelist,['当日盈亏',self.totalAsset - self.report.prevTotalAsset],[F.b,F.b])
        self.write_merge_row(half_mergelist,['当日盈亏比例',(self.totalAsset - self.report.prevTotalAsset) / self.report.prevTotalAsset],[F.b,F.normal_percent])
        self.write_merge_row(mergelist,['份额数',self.shares,'月涨跌幅',(self.totalAsset/self.report.shares) / self.report.prevMonthNet - 1],[F.b,F.c,F.b,F.normal_percent])
        self.write_merge_row(mergelist,['总入金（万）',self.totalIn,'年涨跌幅', (self.totalAsset/self.report.shares) / self.report.prevYearNet - 1],[F.b,F.c,F.b,F.normal_percent])
        return self.localRowNum
    
    def write_holdings(self,startRow):
        mergelist = [(0,),(1,),(2,),(3,),(4,),(5,),(6,),(7,),(8,),(9,),(10,),(11,)]
        self.localRowNum = startRow
        self.rowNumDict['holding'] = startRow
        self.write_merge(0,11,'二、盘后持仓概览',F.d)
        for i,item in enumerate(self.report.holdingList,1):
            self.write_merge_row(mergelist,[i,item['证券代码'],item['证券名称'],item['成本价'],
                                            item['证券数量'],item['当前价'],item['动用资金'],item['动用资金']/self.totalAsset,
                                            item['最新市值'],item['最新市值']/self.totalAsset,item['浮动盈亏'],item['盈亏比例']],
                                 [F.b,F.b,F.b,F.b,F.b,F.b,F.b,F.normal_percent,F.b,F.normal_percent,F.b,F.normal_percent])
            
        self.write_merge_row([(0,5),(6,),(7,),(8,),(9,),(10,),(11,)],['合计',self.holdingSum['used'],self.holdingSum['used']/self.self.holdingSumAsset,
                                                                      self.holdingSum['holds']/self.self.holdingSumAsset,self.holdingSum['profit'],self.holdingSum['profit_pct']],
                                                                      [F.a,F.a,F.normal_percent,F.a,F.normal_percent,F.b,F.bold_percent])
        return self.localRowNum
    
    def write_transaction(self,startRow):
        self.localRowNum = startRow
        self.rowNumDict['transaction'] = startRow
        self.write_merge(0, 11, '三、今日交易', F.d)
        self._write_transaction_merged()
        self._write_transaction_direction()
        return self.localRowNum
    
    def _write_transaction_merged(self):
        self.write_merge(0, 11, '交易指令',F.e)
        self.write_merge_row([(0,),(1,3),(4,5),(6,7),(8,9),(10,11)], 
                             [' ','指令内容','指令下达时间','下达人','执行人','完成情况'], 
                             [F.e,F.b,F.b,F.b,F.b,F.b])
        if len(self.report.mergedTransactionList) == 0:
            self.write_merge_row([(0,),(1,3),(4,5),(6,7),(8,9),(10,11)], 
                        [1,' ',' ',' ',' ',' '], 
                        [F.b,F.b,F.b,F.b,F.b,F.b])
        else:
            for i,item in enumerate(self.report.mergedTransactionList,1):
                self.write_merge_row([(0,),(1,3),(4,5),(6,7),(8,9),(10,11)], 
                                [item['买卖标志']+item['证券名称']+str(item['成交数量'])+'股','9:30:00',item['下达人'],item['执行人'],'完成'], 
                                [F.b,F.b,F.b,F.b,F.b,F.b])
        
    def _write_transaction_direction(self):
        self.write_merge(0, 11, '证券交易',F.e)
        for direction,action in [('买入','动用'),('卖出','释放')]:
            ls = [x for x in self.report.transactionList if direction == x['买卖标志']]
            self.write_merge_row([(0,1),(2,),(3,4),(5,),(6,7),(8,),(9,),(10,11)],
                    ['序号','证券代码','证券名称','交易方向','交易数量','成交均价','收盘价',action+'资金'],
                    [F.b,F.b,F.b,F.b,F.b,F.b,F.b,F.b])
            if len(ls) == 0:
                self.write_merge_row([(0,1),(2,),(3,4),(5,),(6,7),(8,),(9,),(10,11)],
                [1,' ',' ',' ',' ',' ',' ',' '],
                [F.a,F.a,F.a,F.a,F.a,F.a,F.a,F.a])
            else:
                for i,item in enumerate(ls,1):
                    self.write_merge_row([(0,1),(2,),(3,4),(5,),(6,7),(8,),(9,),(10,11)],
                    [i,item['证券代码'],item['证券名称'],item['买卖标志'],item['成交数量'],item['成交价格'],'',item['成交金额']],
                    [F.b,F.b,F.b,F.b,F.b,F.b,F.b,F.b])
            self.write_merge_row([(0,9),(10,11)],['合计{}资金'.format(action),self.transactionSum[action]],[F.a,F.b])
        self.write_merge_row([(0,9),(10,11)],['买卖轧差',self.transactionSum['释放'] - self.transactionSum['动用']],[F.a,F.b])
    
    def write_cash(self,startRow):
        self.localRowNum = startRow
        self.rowNumDict['cash'] = startRow
        
        
    def params_cal(self):
        report = self.report
        #进行一些总体的计算
        