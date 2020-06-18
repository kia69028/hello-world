# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 20:15:01 2020

@author: xinjian.cao
"""

from workSpace import workSpace

class stock_product(workSpace):
    
            


def sheetFactory(name,config):
    if any(elem in name for elem in ['主题','锐进','长嬴','指数增强']):
        return stock_product(name,config)
    if any(elem in name for elem in ['红包','红宝石','可转债']):
        return bond_product(name,config)
    if any(elem in name for elem in ['中性']):
        return neutral_product(name,config)
    
    