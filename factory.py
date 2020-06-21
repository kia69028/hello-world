# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 20:15:01 2020

@author: xinjian.cao
"""


    
class stockReport:
    def __init__(self,report):
        super(stockReport,self).__init__(report)
    
    
    
    
            


def sheetFactory(name,date,config):
    if any(elem in name for elem in ['主题','锐进','长嬴','指数增强']):
        return stockReport(name,date,config)
    if any(elem in name for elem in ['红包','红宝石','可转债']):
        return bondReport(name,date,config)
    if any(elem in name for elem in ['中性']):
        return neutralReportt(name,date,config)
    
    