README
---------------------
   
 * Introduction
 * Requirements
 * Output Files
 
 ---------------------
 INTRODUCTION 
 ---------------------
 This code is to be assisted with the GUI. Ideally, this code should not
 be touched.
 The code has 2-parts, the excel-spreadsheet-scraper & the optimization model.
 The excel-spreadsheet-scraper transforms the data into dictionaries to be 
 parsed through the model.
 Optimization model includes a data input section, and a code section.
 Data input parses through the dictionaries to modify them to suit the code.
 The code contains 2 parts - optimization function, and constraints (3).
 Comments on file for explanation.
 
 ---------------------
 REQUIREMENTS 
 ---------------------
 Can be run on any OS
 Python 2.7
 Gurobi license: http://www.gurobi.com/products/how-to-buy/direct-from-gurobi
 Gurobi Installation: http://www.gurobi.com/registration/download-reg
 
 Required Modules:
 PuLP: https://www.coin-or.org/PuLP/main/installing_pulp_at_home.html
 TkInter: http://www.tkdocs.com/tutorial/install.html
 Openpyxl: https://openpyxl.readthedocs.io/en/stable/
 
 Input Files: (all .xlsx)
 * Air Shipment costs for Next day deliveries
 * Air Shipment costs for Second day deliveries
 * Truck Shipment costs between 2 locations
 * Cost to build a facility at each locations
 * List of locations
 * Demand at each end location 
 
---------------------
OUTPUT FILES 

---------------------
 Output Files: 
 * UPS_network.lp - the optimization model itself, with all constraints
 * gurobi.txt - the solution through the solver
 * result.txt - all demand  variables that are equal to 1 
				this shows which facility is assigned to which customer,
				transporting which level of goods,
				using which mode of transportation,
				and which facility will be opened
				
				Data format:
				for assignment - [ship mode][good value](facility, customer)
				for location opening - Zip(facility)
				Example: 
				
				AirHigh(001, 005) = 1.0 means facility 1 ships high value
									to location 5 using air 
				Zip(001) = 1.0 means facility 1 will be opened
				Zip(176) = 1.0 means facility will be opened at 176xx
									
				
 