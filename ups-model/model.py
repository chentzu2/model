# -*- coding: utf-8 -*-
"""
Created on Tue Apr  3 21:08:04 2018

@author: Tzu-Ying, Pear, Yi
Project name: UPS Optimization Model for Network Design
"""

import pulp
import math
import openpyxl

data = openpyxl.load_workbook('list_zipcode.xlsx',read_only=True, data_only=True)
sheet = data['Sheet1']
title = []
for row in sheet.rows:
    
    for cell in row:
        title.append(cell.value)


data = openpyxl.load_workbook('cost_matrix_truck.xlsx',read_only=True, data_only=True)
sheet = data['cost']

cost_truck = {}

#inputting cost for the truck {i1:{j1:C[i1,j1,m],j2:C[i1,j2,m]...}, i2:{j1:C[i2,j1,m], j2:C[i2,j2,m]...}....} (m=T)
# structure would be like {ori:{dest: ***, dest2:***} ori2: {dest:...}}
for row in sheet.rows:
    temp = []
    for cell in row:
        temp.append(cell.value)
    ori = temp[0]
    temp.pop(0)
    cost_truck[ori] = {}
    for i in range(len(title)):
        cost_truck[ori][title[i]] = []
        cost_truck[ori][title[i]].append(temp[i])



data2 = openpyxl.load_workbook('cost_matrix_deliveryday.xlsx', read_only=True, data_only=True)
sheet2 = data2['day']

delivery_day = {}

#inputting require days for each ori to dest (truck transport)
# structure would be like {ori:{dest: ***, dest2:***} ori2: {dest:...}}

for row in sheet2.rows:
    temp = []

    for cell in row:
        temp.append(cell.value)
    ori = temp[0]
    temp.pop(0)
    delivery_day[ori] = {}
    for i in range(len(title)):
        delivery_day[ori][title[i]] = []
        delivery_day[ori][title[i]].append(temp[i])



# creating dictionary for next day air cost from ori to dest
# structure would be like {ori:{dest: ***, dest2:***} ori2: {dest:...}}
data = openpyxl.load_workbook('cost_matrix_nda.xlsx',read_only=True, data_only=True)
sheet = data['cost']

cost_nda = {}

for row in sheet.rows:
    temp = []
    for cell in row:
        temp.append(cell.value)
    ori = temp[0]
    temp.pop(0)
    cost_nda[ori] = {}
    for i in range(len(title)):
        cost_nda[ori][title[i]] = []
        cost_nda[ori][title[i]].append(temp[i])


# creating dictionary for 2nd day air cost from ori to dest
# structure would be like {ori:{dest: ***, dest2:***} ori2: {dest:...}}
data = openpyxl.load_workbook('cost_matrix_sda.xlsx',read_only=True, data_only=True)
sheet = data['cost']

cost_sda = {}

for row in sheet.rows:
    temp = []
    for cell in row:
        temp.append(cell.value)
    ori = temp[0]
    temp.pop(0)
    cost_sda[ori] = {}
    for i in range(len(title)):
        cost_sda[ori][title[i]] = []
        cost_sda[ori][title[i]].append(temp[i])

# creating dictionary for demand of high value and value of each 3 digit
# structure would be like {3digit: [high, low], ...}
data = openpyxl.load_workbook('demand.xlsx',read_only=True, data_only=True)
sheet = data['Demand']

demand = {}

for row in sheet.rows:
    temp = []
    for cell in row:
        temp.append(cell.value)
    zipcode = temp[0]
    temp.pop(0)
    demand[zipcode] = []
    for i in range(len(temp)):
        demand[zipcode].append(temp[i])

# creating dictionary for facility F
# structure would be like {3digit: cost, 3digit2: cost, ... }
data = openpyxl.load_workbook('facility_cost.xlsx',read_only=True, data_only=True)
sheet = data['Demand']

facility_cost = {}

for row in sheet.rows:
    temp = []
    for cell in row:
        temp.append(cell.value)
    zipcode = temp[0]
    temp.pop(0)
    facility_cost[zipcode] = temp[0]
    

#pear optimization code
# Data Section
CustLoc=[] #a set of customer locations
Zip=[] # set of 3-digit zip code
shipmax = 1 #change from next day to 2 days - connect to GUI
F = 10000 #fixed cost to operate/open facility  
V = [20,150]#V[k] = [Vl, Vh] - value of LV and HV goods = [20,150]
M = [1,2] #[air, truck]
cost_air={1:cost_nda, 2:cust_sda}

"""
Berth Assignment Problem Model

This version uses the relative position formulation, and
minimizes only total completion time of all vessels

Authors: Alan Erera 2014
"""

# Import PuLP modeler functions
from pulp import *

# Data Section
CustLoc=[] #a set of customer locations
Zip=[] # set of 3-digit zip code
shipmax = 1 #change from next day to 2 days - connect to GUI
F = 10000 #fixed cost to operate/open facility  
M = [1,2] #[air, truck]
costAir={1:cost_nda, 2:cust_sda} #{ori:{dest: ***, dest2:***} ori2: {dest:...}}
trucktime= delivery_day #{} #delivery_day={ori:{dest: ***, dest2:***} ori2: {dest:...}}
costTruck = cost_Truck #{ori:{dest: ***, dest2:***} ori2: {dest:...}}
d = demand #{3digit: [high, low], ...}
# Code Section

# Create the 'prob' object to contain the problem data
 #create prob object to contain optimization problem data
 problem = pulp.LpProblem("Facility Location Plan", pulp.LpMinimize)

# Decision variables

combo={}
for i in CustLoc:
	for j in Zip:
		combo['route'] =(i,j)
		combo['customer']=i
		combo['facility']=j

for a,a_dict in combo.iteritems():
	cl = a[0]
	z = a[1]
	
	# varAH = pulp.LpVariable("AirHigh(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	# a_dict['dvAH'] = varAH

	# varAL = pulp.LpVariable("AirLow(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	# a_dict['dvAL'] = varAL

	# varTH = pulp.LpVariable("TruckHigh(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	# a_dict['dvTH'] = varTH

	# varTL = pulp.LpVariable("TruckLow(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	# a_dict['dvTL'] = varTL

	varAH = pulp.LpVariable("AirHigh(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	varAL = pulp.LpVariable("AirLow(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	varTH = pulp.LpVariable("TruckHigh(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	varTL = pulp.LpVariable("TruckLow(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	
	a_dict['dv']=[varAH, varAL, varTH, varTL]

	objFn.append((varAH*D[cl][0]+varAL*D[cl][1])*costAir[shipmax][cl][z]*150 + (varTH*D[cl][0]+varAL*D[cl][1])*costTruck[cl][z]*20)

FacilityLocations={}
for j in Zip:
	FacilityLocations['Location']=j

for j,j_dict in FacilityLocations.iteritems():
	dvLoc = pulp.LpVariable("Zip(%s)"%str(j), lowBound=0, upBound=1, cat=pulp.LpInteger)
	j_dict["Locate?"] = dvLoc
	objFn.append(dvLoc*F)
	

# # The objective function is added to 'prob' first
# # lpSum takes a list of coefficients*LpVariables and makes a summation
prob += pulp.lpSum(objFn), "Total Cost"


# Flow balance at all nodes for all commodities
for k,k_dict in commods.iteritems():
	orig = k[0]
	dest = k[1]
	for i in nodes:
		# If i is the orig of the commodity, the net supply is q
		if i == orig:
			netsupply = k_dict['q']
		elif i == dest:
			netsupply = - k_dict['q']
		else:
			netsupply = 0
			
		# Create the flow balance constraint	
		prob += pulp.lpSum(arcs[a]['dvFlows'][k] for a in nodes[i]['outArcs']) - pulp.lpSum(arcs[a]['dvFlows'][k] for a in nodes[i]['inArcs']) == netsupply, "Node %s Commodity (%s,%s) Flow Balance" % (str(i),str(orig),str(dest))

# Round up the arc flows to trailer flows, across commodities
for a,a_dict in arcs.iteritems():
	i = a[0]
	j = a[1]
	# Sum over all of the commodity-specific arc flows
	prob += pulp.lpSum(a_dict['dvFlows'][k] for k in commods) <= a_dict['dvTrailerFlow'], "Arc (%s,%s) Trailer Roundup" % (str(i),str(j))



for m in M:
	for i in Zip:
		for j in CustLoc:
			objFn.append(([C[m][i][j]*d[j][val]*V[val]*x[i,j,m,val] for val in V]), "Cost to transport")

 #objective function
 #V[k]*C[i,j,m]*D[i,k]*x[i,j,m,k]   +F*y[j]
#C= {ori:{dest: ***, dest2:***} ori2: {dest:...}}

for a,a_dict in Combo.iteritems():
	origin=a['route'][1] #zip
	dest=a['route'][0] #customer

	prob +=LpSum(a['dvAH']+a['dvAL']+a['dvTH']+a['dvTL']>=1) #for all i,j: sum[m,k](x[i,j,m,k])>=1

for i in CustLoc:
	for a,a_dict in combo.iterItems():
		if a['customer']==i:












a_dict['dvAirHighFlow'] = varAH

	varAL = pulp.LpVariable("AirLow(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	a_dict['dvAirLowFlow'] = varAL

	varTH = pulp.LpVariable("TruckHigh(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	a_dict['dvTruckHighFlow'] = varTH

	varTL = pulp.LpVariable("TruckLow(%s,%s)"%(str(cl),str(z)), lowBound=0, cat=pulp.LpInteger)
	a_dict['dvTruckLowFlow'] = varTL

#constraint 1
for i in Zip:
	for j in CustLoc:
		for v in V:
			for m in M:
				if m == 1:
					prob+=t[m] <=shipmax+M*(1-x[i,j,m,v])
				else:
					prob += t[2][i][j] <= shipmax+M*(1-x[i,j,m,v])

     #for all i,j,m,k: t[i,j,m]<=1+M*(1-x[i,j,m,k])
     #for all i,j,m,k: t[i,j,m]<=2+M*(1-x[i,j,m,k]) 

 #constraint 2
 
 #constraint 3
 for j in CustLoc:
 	for m in M:
 		for k in V:
 			prob+=LpSum(x[i,j,m,k] for i in Zip) <=M*y[j]
 #for all j: sum[k,m,i](x[i,j,m,k])<=M[j]
 #M = big number

# Constraints
	
# Write out as a .LP file
prob.writeLP("UPS_network.lp")

# The problem is solved using PuLP's choice of Solver
prob.solve(GUROBI())
#prob.solve()

# The status of the solution is printed to the screen
print "Status:", LpStatus[prob.status]

# Each of the variables is printed with it's resolved optimum value
# for a in x:
#     print( v.name, "=", v.varValue)
#print out only locations that were planted. also print out #no.

# The optimised objective function value is printed to the screen    
print "Total Cost = ", value(prob.objective)
