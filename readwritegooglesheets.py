#Library of gspread and oauth2client for the api
import gspread
from oauth2client.service_account import ServiceAccountCredentials

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
Customer_Data = file.open("Customer_Quotation_Data")


#Importing sheet 1 of all of them
Hybrid_Offgrid_Inverters = Hybrid_Offgrid_Inverter_prices.sheet1
Grid_Tied_Inverters = Grid_Tied_Inverter_prices.sheet1
Solar_Panels = Solar_Panel_Prices.sheet1
Customer_Data_Sheet = Customer_Data.sheet1

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

#clear selection of name and price and wattage of the inverters and solar panels.
Hybrid_Inverter_Names = Hybrid_name_of_inverters[1::2]