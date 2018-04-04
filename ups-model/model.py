# -*- coding: utf-8 -*-
"""
Created on Tue Apr  3 21:08:04 2018

@author: Tzu-Ying, Pear, Yi
Project name: UPS Optimization Model for Network Design
"""

import pulp
import math
import openpyxl


data = openpyxl.load_workbook('filename',read_only=True, data_only=True)
sheet = data['Sheet1']

cost_truck = {}
title = list(sheet.rows[0])



for row in sheet.rows:
    temp = []
    if row != title:
        for cell in row:
            temp.append(cell.value)
    ori = temp[0]
    temp.pop(0)
    cost_truck[ori] = {}
    for i in title:
        cost_truck[ori][i] = temp[i]


data2 = openpyxl.load_workbook('filename', read_only=True, data_only=True)
sheet2 = data2['Sheet1']

delivery_day = {}
title = list(sheet2.row[0])


for row in sheet2.rows:
    temp = []
    if row != title:
        for cell in row:
            temp.append(cell.value)
    ori = temp[0]
    temp.pop(0)
    delivery_day[ori] = {}
    for i in title:
        delivery_day[ori][i] = temp[i]