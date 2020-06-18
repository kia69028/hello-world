# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 20:43:17 2020

@author: xinjian.cao
"""
import re,json,os,datetime
from utility import get_line,isfloat,F,T,xt,workdays

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
        self.securityPosition = 0
        self.bondPosition = 0
        self.holdingList = []
        self.transactionList = [] #买卖成交分开
        self.mergedTransactionList = [] #合并日内回转交易
        self.money_fund = []
        self.cash_management = []
        self.rowNum = 0
        self.workBook = xt.Workbook(encoding='utf-8')
        self.workSpace = slef.workBook.add_sheet(productName,cell_overwrite_ok=True)
        self.filepath = '{}交易汇总\\'.format(date)
        self.workSpace.set_portrait(0)
        self.currentName = self.filepath+'{}每日交易汇总{}-自动.xls'.format(self.name,self.date_str)
        if not os.path.exists(self.filepath):
            os.mkdir(self.filepath)
    
    def write_merge_row(self,mergelist,contentlist,stylelist):
        assert (len(mergelist) == len(contentlist) and len(contentlist) == len(stylelist)), 'length of mergelist/contentlist/stylelist does not match!'
        for i in range(len(mergelist)):
            if(len(mergelist[i]) < 2):
                self.workSpace.write(self.rowNum,mergelist[i][0],contentlist[i],stylelist[i])
            else:
                self.workSpace.write_merge(self.rowNum,self.rowNum,mergelist[i][0],mergelist[i][1],contentlist[i],stylelist[i])
        self.rowNum += 1
        
    
    def write_header(self):
        self.write_merge_row([(0,2),(3,5),(6,7),(8,9),(10,11)], 
                ['名称',self.name+'私募基金','日期',self.date_dt.strftime("%Y/%m/%d"),
                 workdays[int(self.date_dt.strftime('%w'))]], 
                [F.a,F.a,F.a,F.a,F.a])
    
    def write_overview(self):
        self.write_merge_row([(0,11)],['一、盘后账户情况'],[F.d])
        self.write_merge_row([(0,2),(3,5),(6,8),(9,11)],['总资产',self.totalAsset,'总仓位（权益+可转债）',])
    workSpace.write_merge(2,2,0,2,'总资产',F.b)
    workSpace.write_merge(2,2,3,5,totalAsset,F.b)
    workSpace.write_merge(2,2,6,8,'总仓位（权益+可转债）',F.b)
    workSpace.write_merge(2,2,9,11,xt.Formula('J4+J5'),F.normal_percent)
    
    workSpace.write_merge(3,3,0,2,'可用资金（含逆回购）',F.b)
    
    workSpace.write_merge(4,4,0,2,'总盈亏',F.b)
    workSpace.write_merge(4,4,3,5,xt.Formula('D3-D10*10000'),F.b)
              
    workSpace.write_merge(5,5,0,2,'可加仓资金（80%仓位）(万)',F.b)
    workSpace.write_merge(5,5,3,5,xt.Formula('MIN(D4,(0.8-D7)*D3)/10000'),F.two_decimal)      
    
    workSpace.write_merge(5,5,6,8,'净值',F.b)
    workSpace.write_merge(5,5,9,11,xt.Formula('D3/D9'),F.four_decimal)
    
    workSpace.write_merge(6,6,6,8,'当日盈亏',F.b)
    workSpace.write_merge(6,6,9,11,xt.Formula('D3-{}'.format(lastDay['上日总资产'])),F.b)

    workSpace.write_merge(7,7,0,2,'是否超限',F.b)
    workSpace.write_merge(7,7,3,5,' ',F.b)
    workSpace.write_merge(7,7,6,8,'当日盈亏比例',F.b)
    workSpace.write_merge(7,7,9,11,xt.Formula('(D3-{})/{}'.format(lastDay['上日总资产'],lastDay['上日总资产'])),F.normal_percent)  
    
    workSpace.write_merge(8,8,0,2,'份额数',F.b)
    workSpace.write_merge(8,8,3,5,lastDay['份额数'],F.c)
    workSpace.write_merge(8,8,6,8,'月涨跌幅',F.b)
    workSpace.write_merge(8,8,9,11,xt.Formula('J6/{}-1'.format(lastDay['上月末净值'])),F.normal_percent) 
    
    workSpace.write_merge(9,9,0,2,'总入金（万）',F.b)
    workSpace.write_merge(9,9,3,5,lastDay['总入金'],F.c)
    workSpace.write_merge(9,9,6,8,'年涨跌幅',F.b)
    workSpace.write_merge(9,9,9,11,xt.Formula('J6/{}-1'.format(lastDay['去年净值'])),F.normal_percent)  
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
            totalRevoke = 0
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
                self.transactionList.append(item)
                
        for rowDict in merged_dc.values():
            if rowDict['成交数量'] == 0：:
                continue
            rowDict['买卖标志'] = '买入' if rowDict['成交数量'] > 0 else '卖出'
            rowDict['成交数量'] = abs(rowDict['成交数量'])
            self.mergedTransactionList.append(rowDict)    
        
    def read_revoke(self,raw_revoke):
        for brokerName,lines in raw_revoke:
            headings = get_line(lines,0)
            format_adj = json.loads(config.get(brokerName,'撤单格式'))
            for item in format_adj.keys():
                pos = headings.index(item)
                headings[pos] = format_adj[item]
            for i in range(1,len(lines)):
                row = get_line(lines,i)
                row = [float(item) if isfloat(item) else item for item in row]
                rowDict = dict(zip(headings,row))
                if '买入' in rowDict['买卖标志']:
                    self.totalAvailable += (rowDict['委托数量']-rowDict['成交数量'])*rowDict['委托价格']
    def write_
    def write_transactionOrder(workSpace,rowNum,transactionLs,executor):
        dc = {}
        for item in transactionLs:
            sign = 1 if ('买入' in item['买卖标志']) else -1
            if(item['证券名称'] not in dc.keys()):
                dc[item['证券名称']] = sign*item['成交数量']
            else:
                dc[item['证券名称']] += sign*item['成交数量']
        
        '''
        if len(dc) == 0:
            rowNum = write_merge_row(workSpace, rowNum, [(0,),(1,3),(4,5),(6,7),(8,9),(10,11)], 
                            [1,' ',' ',' ',' ',' '], 
                            [F.b,F.b,F.b,F.b,F.b,F.b])
        '''
        
        for i,item in enumerate(dc,1):
            workSpace.write(rowNum,0,i,F.b)
            direction = '买入' if (dc[item] > 0) else '卖出'
            if dc[item] == 0:
                continue
            xiadaren = executor['下达人'] if executor else ' '
            zhixingren = executor['执行人'] if executor else ' '
            rowNum = write_merge_row(workSpace, rowNum, [(1,3),(4,5),(6,7),(8,9),(10,11)], 
                            [direction+item+str(abs(int(dc[item])))+'股','9:30:00',xiadaren,zhixingren,'完成'], 
                            [F.b,F.b,F.b,F.b,F.b])
        return rowNum
    
    def write_securityTrans(workSpace,rowNum,transactionLs,direction):
        write_merge_row(workSpace,rowNum,[(0,1),(2,),(3,4),(5,),(6,7),(8,),(9,)],
                        ['序号','证券代码','证券名称','交易方向','交易数量','成交均价','收盘价'],
                        [F.b,F.b,F.b,F.b,F.b,F.b,F.b])
    
        if (direction == "买入"):
            workSpace.write_merge(rowNum,rowNum,10,11,'动用资金',F.b)
        else:
            workSpace.write_merge(rowNum,rowNum,10,11,'释放资金',F.b)
        rowNum += 1
        
        ls = [x for x in transactionLs if (direction in x['买卖标志'])]
        _sum = 0
        
        dc = {}
        for item in ls:
            if item['证券名称'] not in dc.keys():
                dc[item['证券名称']] = item
            else:
                dc[item['证券名称']]['成交数量'] += item['成交数量']
                dc[item['证券名称']]['成交金额'] += item['成交数量']*item['成交价格']
                
        if (len(dc) == 0):
           rowNum = write_merge_row(workSpace,rowNum,[(0,1),(2,),(3,4),(5,),(6,7),(8,),(9,),(10,11)],
                    [1,' ',' ',' ',' ',' ',' ',' '],
                    [F.a,F.a,F.a,F.a,F.a,F.a,F.a,F.a])
           
        for i,item in enumerate(list(dc.values()),1):
            item['买卖标志'] = direction
            item['成交价格'] = item['成交金额'] / item['成交数量']
            rowNum = write_merge_row(workSpace, rowNum, [(0,1),(2,),(3,4),(5,),(6,7),(8,),(9,),(10,11)], 
                                     [i,item['证券代码'],item['证券名称'],item['买卖标志'],
                                      item['成交数量'],item['成交价格'],' ',item['成交金额']], 
                                     [F.b,F.b,F.b,F.b,F.b,F.b,F.b,F.b])
            _sum += item['成交金额']
        
        if (direction == '买入'):
            workSpace.write_merge(rowNum,rowNum,T.A,T.J,'合计动用资金',F.a)
            workSpace.write_merge(rowNum,rowNum,10,11,_sum,F.b)
        else:
            workSpace.write_merge(rowNum,rowNum,T.A,T.J,'合计释放资金',F.a)
            workSpace.write_merge(rowNum,rowNum,10,11,_sum,F.b)
        rowNum += 1
        
        return (rowNum,_sum)