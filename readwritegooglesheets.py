#Library of gspread and oauth2client for the api
import gspread
from oauth2client.service_account import ServiceAccountCredentials
#import time

#Scopes Required for the running of the google docs api
scopes=[
    'https://www.googleapis.com/auth/spreadsheets'
    'https://www.googleapis.com/auth/drive'
]

#File which contains secret key should be in Data Folder
creds = ServiceAccountCredentials.from_json_keyfile_name("./Data/secret_key.json")

#Authorizing the api with secret key file
file = gspread.authorize(creds)

#Importing all three sheets with Hybrid Off grid and Grid Tied and Also Solar Panels Prices
Hybrid_Offgrid_Inverter_prices = file.open('Hydrid_Offgrid_Inverter_prices')
Grid_Tied_Inverter_prices = file.open("Grid_Tied_Inverter_Rates")
Solar_Panel_Prices = file.open("solar_panel_prices")


#Importing sheet 1 of all of them
Hybrid_Offgrid_Inverters = Hybrid_Offgrid_Inverter_prices.sheet1
Grid_Tied_Inverters = Grid_Tied_Inverter_prices.sheet1
Solar_Panels = Solar_Panel_Prices.sheet1

#Importing Inverter Names with seperated values
Grid_Tie_name_of_inverters = Grid_Tied_Inverters.col_values(1)

#Importing name wattage and price of hybrid off grid inverters
Hybrid_name_of_inverters = Hybrid_Offgrid_Inverters.col_values(1)

#Importing name wattage and price of Solar Panels
Solar_Panels_Names = Solar_Panels.col_values(1)
Solar_Panel_Price = Solar_Panels.col_values(2)
Solar_Panel_Wattage = Solar_Panels.col_values(3)

#clear selection of name and price and wattage of the inverters and solar panels.
Grid_Tie_Inverter_Names = Grid_Tie_name_of_inverters[1::2]

#Grid_Tied_Inverters_Wattage = []
#Grid_Tied_Inverters_Prices = []

#for i in range(2,30,2):
#    for x in Grid_Tied_Inverters.row_values(i):
#        Grid_Tied_Inverters_Wattage.append(x)

#Grid_Tied_Inverters_Wattage_Indexer = Grid_Tied_Inverters_Wattage.copy()

#while("" in Grid_Tied_Inverters_Wattage):
#    Grid_Tied_Inverters_Wattage.remove("")

#time.sleep(30)

#for i in range(1,30,2):
#    for x in Grid_Tied_Inverters.row_values(i):
#       Grid_Tied_Inverters_Prices.append(x)

#Grid_Tied_Inverters_Prices_Indexer = Grid_Tied_Inverters_Prices.copy()

#while("" in Grid_Tied_Inverters_Prices):
#    Grid_Tied_Inverters_Prices.remove("")

Hybrid_Inverter_Names = Hybrid_name_of_inverters[1::2]

#Hybrid_Inverters_Wattage = []
#Hybrid_Inverters_Prices = []

#for i in range(2,30,2):
#    for x in Hybrid_Offgrid_Inverters.row_values(i):
#        Hybrid_Inverters_Wattage.append(x)

#Hybrid_Inverters_Wattage_Indexer = Hybrid_Inverters_Wattage.copy()

#while("" in Hybrid_Inverters_Wattage):
#    Hybrid_Inverters_Wattage.remove("")

#for i in range(1,30,2):
#    for x in Hybrid_Offgrid_Inverters.row_values(i):
#        Hybrid_Inverters_Prices.append(x)

#while("" in Hybrid_Inverters_Prices):
#    Hybrid_Inverters_Prices.remove("")

#Hybrid_Inverters_Prices_Indexer = Hybrid_Inverters_Prices.copy()

#time.sleep(30)