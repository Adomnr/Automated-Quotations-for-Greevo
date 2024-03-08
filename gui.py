#Importing Libraries
import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl
from readwritegooglesheets import *
import docx
#For Value in words.
import inflect
from spire.doc import *
from spire.doc.common import *
from datetime import datetime
from doctest import *
import uuid
import time


Inverter_name = []
Inverter_wattage = []
template_file = ""
InverterPrice = 0
panelprice = 0
panelwattage = 0
TotalCostNormal = 0
TotalCostRaised = 0
UniqueID = 0
serialNumber = 0
valueinwords = ""

generalrow = 0
generalrowentry = 1

referredrow = 2
referredrowentry=3

inverterrow = 4
inverterrowentry = 5

panelrow = 6
panelrowentry=7

pvrow = 8
pvrowentry = 9

cinrow = 8
cinrowentry = 9

def getSerialNumber(*args):
    global serialNumber
    index = []
    for x in Customer_Data_Sheet.col_values(1):
        index.append(x)
    serialNumber = int(index[-1])

def enter_data():
    global valueinwords
    accepted = accept_var.get()
    total_cost_calculator()
    UniqueID = generate_unique_id()
    if structure_type_combobox.get() == "Raised":
        valueinwords = convert_to_words(TotalCostRaised)
    else:
        valueinwords = convert_to_words(TotalCostNormal)

    if accepted == "Accepted":
        # User info
        SystemSize = System_Size_combobox.get()
        ClientName = Client_Name_combobox.get()
        ClientLocation = Client_Location_combobox.get()
        ReferredBy = Reffered_combobox.get()
        if SystemSize and ClientName and ClientLocation and ReferredBy:
            Inverter_TYP = inverter_type_combobox.get()
            Inverter_Name = inverter_name_combobox.get()
            Inverter_Watt = inverter_wattage_combobox.get()

            if Inverter_TYP and Inverter_Name and Inverter_Watt:
                Name_of_Panels = panel_name_combobox.get()
                Number_of_Panels = number_of_panels_spinbox.get()
                if Name_of_Panels and Number_of_Panels:
                    # Course info
                    Structure_Type = structure_type_combobox.get()
                    pv_balance = pv_balance_combobox.get()
                    if Structure_Type and pv_balance:
                        Carriage_Cost = carriage_entry.get()
                        Installation = installation_entry.get()
                        Net_Metering = net_metering_entry.get()
                        if Carriage_Cost and Installation and Net_Metering:
                            print("System Size: ", SystemSize, "Client Name: ", ClientName, "Client Location: ", ClientLocation)
                            print("Reffered by: ", ReferredBy, "Inverter Type: ", Inverter_TYP, "Inverter Name: ", Inverter_Name)
                            print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels, "Panel Price: ", panelprice)
                            print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                            print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost, "Installation Cost: ", Installation)
                            print("Net Metering: ", Net_Metering, "Template File", template_file, "Inverter Price", InverterPrice)
                            print("Total Cost Normal: ",TotalCostNormal,"Total Cost Raised: ",TotalCostRaised, "Unique ID: ", UniqueID)
                            print("------------------------------------------")

                            record_data(SystemSize,UniqueID,ClientName, ClientLocation, ReferredBy, Inverter_TYP, Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice,Number_of_Panels, Structure_Type, pv_balance, Carriage_Cost, Installation,Net_Metering, template_file, InverterPrice,TotalCostNormal,TotalCostRaised)

                            document_creater(SystemSize, Inverter_Watt, Number_of_Panels,Net_Metering, template_file, TotalCostNormal, TotalCostRaised, valueinwords)
                        else:
                            tkinter.messagebox.showwarning(title="Error", message="Enter Carriage, Installation and Net Metering Cost")
                    else:
                        tkinter.messagebox.showwarning(title="Error", message="Enter Structure Type and PV Balance.")
                else:
                    tkinter.messagebox.showwarning(title="Error", message="Fill all box of Solar Panels.")
            else:
                tkinter.messagebox.showwarning(title="Error", message="Enter All box of Inverters.")
        else:
            tkinter.messagebox.showwarning(title="Error", message="Enter All boxes of Client Information.")
    else:
        tkinter.messagebox.showwarning(title="Error", message="You have not accepted the terms")

#This updates the template which is going to be subsituited the placeholder into GTGN Grid Tied General Normal

def total_cost_calculator(*args):
    global TotalCostNormal
    global TotalCostRaised
    Structure_Type = structure_type_combobox.get()
    Number_of_Panels = number_of_panels_spinbox.get()
    pv_balance = pv_balance_combobox.get()
    Carriage_Cost = carriage_entry.get()
    Installation = installation_entry.get()
    Net_Metering = net_metering_entry.get()
    TotalCostNormal = int(InverterPrice)+(20*int(Number_of_Panels)*int(panelwattage))+(int(panelwattage)*int(panelprice)*int(Number_of_Panels))+int(pv_balance)+int(Carriage_Cost)+int(Installation)+int(Net_Metering)
    TotalCostRaised = int(InverterPrice) + (6500*(int(Number_of_Panels)/2)) + (int(panelwattage) * int(panelprice) * int(Number_of_Panels)) + int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Net_Metering)

def convert_to_words(number):
    p = inflect.engine()
    words = p.number_to_words(number)
    return words

def generate_unique_id():
    # Get current timestamp in seconds
    timestamp = int(time.time())

    # Generate a random component within the range of 3 digits
    random_component = uuid.uuid4().int % 1000  # 3-digit random component

    # Combine timestamp and random component to create a unique ID
    unique_id = int(f"{timestamp:03d}{random_component:03d}")%10000000

    return unique_id

def record_data(SystemSize,UniqueID,ClientName, ClientLocation, ReferredBy, Inverter_TYP, Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice,Number_of_Panels, Structure_Type, pv_balance, Carriage_Cost, Installation,Net_Metering, template_file, InverterPrice,TotalCostNormal,TotalCostRaised):
    serial_list = []
    current_date_time = datetime.now()
    current_date_time = str(current_date_time)
    for x in Customer_Data_Sheet.col_values(1):
        serial_list.append(x)
    SerialNumber = int(serial_list[-1]) + 1
    if SerialNumber == 0:
        SerialNumber = 1
    Customer_Data_Sheet.append_row([SerialNumber,current_date_time,UniqueID,SystemSize, ClientName,ClientLocation, ReferredBy, Inverter_TYP, Inverter_Name,Inverter_Watt, Name_of_Panels, panelprice, Number_of_Panels,Structure_Type, pv_balance, Carriage_Cost, Installation,Net_Metering,InverterPrice,TotalCostNormal,TotalCostRaised])

def update_template_type(*args):
    global template_file
    if inverter_type_combobox.get() == "Grid Tie":
        template_file = ".\\Templates\\GTGN_Template.docx"
        if structure_type_combobox.get() == "Raised":
            template_file = ".\\Templates\\GTGR_Template.docx"
    elif inverter_type_combobox.get() == "Hybrid":
        template_file = ".\\Templates\\HGN_Template.docx"
        if structure_type_combobox.get() == "Raised":
            template_file = ".\\Templates\\HGR_Template.docx"
    print(template_file)

def inverter_selection(*args):
    inverter_name_combobox.set('')
    inverter_wattage_combobox.set('')
    if inverter_type_combobox.get() == "Grid Tie":
        Inverter_name = Grid_Tie_Inverter_Names.copy()
    else:
        Inverter_name = Hybrid_Inverter_Names.copy()
    inverter_name_combobox['values'] = Inverter_name

def update_panel(*args):
    global panelprice
    global panelwattage
    i = panel_name_combobox.current()
    panelprice = Solar_Panel_Price[i]
    print(panelprice)
    for x in Solar_Panel_Wattage:
        panelwattage = Solar_Panel_Wattage[i]
    print(panelwattage)

def inverter_wattage_selection(*args):
    inverter_wattage_combobox.set('')
    new_list = []
    if inverter_type_combobox.get() == "Grid Tie":
        if inverter_name_combobox.get() != "":
            index = (inverter_name_combobox.current()+1) * 2
            for x in Grid_Tied_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter_wattage_combobox.config(values = new_list)
    else:
        if inverter_name_combobox.get() != "":
            index = (inverter_name_combobox.current()+1) * 2
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter_wattage_combobox.config(values = new_list)

def inverter_price(*args):
    index = ((inverter_name_combobox.current() + 2) * 2)-1
    list = []
    if inverter_type_combobox.get() == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
        list.pop(0)
        while("" in list):
            list.remove("")
    else:
        if inverter_type_combobox.get() == "Hybrid":
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                list.append(x)
            while("" in list):
                list.remove("")
    print(list)
    if inverter_wattage_combobox.current() >= 0 and inverter_wattage_combobox.current() <= 20:
        index2 = int(inverter_wattage_combobox.current())
        global InverterPrice
        InverterPrice = list[index2]
        print(InverterPrice)

#Inverter_name,name_of_solar_panels,wattage_of_solar_panels

window = tkinter.Tk()
window.title("Quotation Automation")

frame = tkinter.Frame(window)
frame.pack()

# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="Panel Inverters and PV Balance")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

System_Size_label = tkinter.Label(user_info_frame, text="System Size")
System_Size_combobox = ttk.Entry(user_info_frame)
System_Size_label.grid(row=generalrow, column=0)
System_Size_combobox.grid(row=generalrowentry, column=0)

Client_Name_label = tkinter.Label(user_info_frame, text="Client Name")
Client_Name_combobox = ttk.Entry(user_info_frame)
Client_Name_label.grid(row=generalrow, column=1)
Client_Name_combobox.grid(row=generalrowentry, column=1)

Client_Location_label = tkinter.Label(user_info_frame, text="Location")
Client_Location_combobox = ttk.Entry(user_info_frame)
Client_Location_label.grid(row=generalrow, column=2)
Client_Location_combobox.grid(row=generalrowentry, column=2)

Reffered_label = tkinter.Label(user_info_frame, text="Reference By")
Reffered_combobox = ttk.Combobox(user_info_frame, values=["Madam Rafia", "Engr Sajjad", "Engr Shaban", "Engr Abid", "Engr Ammar", "Engr Ammar Butt","Engr Ubaid", "Sir Nabeel"])
Reffered_label.grid(row=referredrow, column=1)
Reffered_combobox.grid(row=referredrowentry, column=1)

sel = tkinter.StringVar(user_info_frame)
sel3 = tkinter.StringVar(user_info_frame)
sel4 = tkinter.StringVar(user_info_frame)
sel5 = tkinter.StringVar(user_info_frame)

inverter_type_label = tkinter.Label(user_info_frame, text="Inverter Type")
inverter_type_combobox = ttk.Combobox(user_info_frame, values=["Grid Tie", "Hybrid"], textvariable=sel)
inverter_type_label.grid(row=inverterrow, column=0)
inverter_type_combobox.grid(row=inverterrowentry, column=0)

sel.trace('w',inverter_selection)
inverter_name_label = tkinter.Label(user_info_frame, text="Inverter Name")
inverter_name_combobox = ttk.Combobox(user_info_frame, values=Inverter_name,textvariable=sel3)
inverter_name_label.grid(row=inverterrow, column=1)
inverter_name_combobox.grid(row=inverterrowentry, column=1)

inverter_wattage_label = tkinter.Label(user_info_frame, text="Wattage")
inverter_wattage_combobox = ttk.Combobox(user_info_frame, values=Inverter_wattage, textvariable=sel5)
inverter_wattage_label.grid(row=inverterrow, column=2)
inverter_wattage_combobox.grid(row=inverterrowentry, column=2)

sel3.trace('w',inverter_wattage_selection)
sel2 = tkinter.StringVar(user_info_frame)


panel_name_label = tkinter.Label(user_info_frame, text="Panel Name")
panel_name_combobox = ttk.Combobox(user_info_frame, values=Solar_Panels_Names, textvariable=sel2)
panel_name_label.grid(row=panelrow, column=0)
panel_name_combobox.grid(row=panelrowentry, column=0)

sel5.trace('w',inverter_price)
sel2.trace('w',update_panel)

number_of_panels_label = tkinter.Label(user_info_frame, text="No. of Panels")
number_of_panels_spinbox = tkinter.Spinbox(user_info_frame, from_=0, to=1000)
number_of_panels_label.grid(row=panelrow, column=1)
number_of_panels_spinbox.grid(row=panelrowentry, column=1)

structure_type = tkinter.Label(user_info_frame, text="Structure Type")
structure_type_combobox = ttk.Combobox(user_info_frame, values=["Normal", "Raised"], textvariable=sel4)
structure_type.grid(row=panelrow, column=2)
structure_type_combobox.grid(row=panelrowentry, column=2)


pv_balance_label = tkinter.Label(user_info_frame, text="PV Balance")
pv_balance_combobox = ttk.Entry(user_info_frame)
pv_balance_label.grid(row=pvrow, column=1)
pv_balance_combobox.grid(row=pvrowentry, column=1)

sel4.trace('w',update_template_type)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving Course Info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

carriage = tkinter.Label(courses_frame, text="Carriage")
carriage_entry = ttk.Entry(courses_frame)
carriage.grid(row = cinrow, column=0)
carriage_entry.grid(row=cinrowentry, column=0)

installation = tkinter.Label(courses_frame, text="Installation")
installation_entry = ttk.Entry(courses_frame)
installation.grid(row=cinrow, column=1)
installation_entry.grid(row=cinrowentry, column=1)

net_metering = tkinter.Label(courses_frame, text="Net Metering")
net_metering_entry = ttk.Entry(courses_frame)
net_metering.grid(row=cinrow, column=2)
net_metering_entry.grid(row=cinrowentry, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms and conditions.", variable=accept_var,
                                      onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = tkinter.Button(frame, text="Enter data", command=enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

window.mainloop()