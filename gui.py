import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl
from readwritegooglesheets import *
import docx

Inverter_name = []
Inverter_wattage = []
template_file = ""
InverterPrice = 0

generalrow = 0
generalrowentry = 1

inverterrow = 2
inverterrowentry = 3

panelrow = 4
panelrowentry=5

pvrow = 6
pvrowentry = 7

cinrow = 8
cinrowentry = 9
def update_template_type(*args):
    inverter_price()
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

def inverter_price(*args):
    index = (inverter_name_combobox.current() + 2) * 2
    list = []
    if inverter_type_combobox == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
        list.pop(0)
    else:
        if inverter_type_combobox == "Hybrid":
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                list.append(x)
            list.pop(0)
    if inverter_wattage_combobox.current() >= 0 and inverter_wattage_combobox.current() <= 20:
        index2 = inverter_wattage_combobox.current()
        global InverterPrice
        InverterPrice = list[index2]
        print(InverterPrice)

def enter_data():
    accepted = accept_var.get()

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
                Wattage_of_Panels = panel_wattage_combobox.get()
                Number_of_Panels = number_of_panels_spinbox.get()
                if Name_of_Panels and Wattage_of_Panels and Number_of_Panels:
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
                            print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels, "Panel Wattage", Wattage_of_Panels)
                            print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                            print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost, "Installation Cost: ", Installation)
                            print("Net Metering: ", Net_Metering, "Template File", template_file, "Inverter Price", InverterPrice)
                            print("------------------------------------------")

                            #filepath = "as"

                            #if not os.path.exists(filepath):
                            #    workbook = openpyxl.Workbook()
                            #    sheet = workbook.active
                            #    heading = ["First Name", "Last Name", "Title", "Age", "Nationality",
                            #               "# Courses", "# Semesters", "Registration status"]
                            #    sheet.append(heading)
                            #    workbook.save(filepath)
                            #workbook = openpyxl.load_workbook(filepath)
                            #sheet = workbook.active
                            #sheet.append([firstname, lastname, title, age, nationality, numcourses,
                            #              numsemesters, registration_status])
                            #workbook.save(filepath)
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

def inverter_selection(*args):
    inverter_name_combobox.set('')
    inverter_wattage_combobox.set('')
    if inverter_type_combobox.get() == "Grid Tie":
        Inverter_name = Grid_Tie_Inverter_Names.copy()
    else:
        Inverter_name = Hybrid_Inverter_Names.copy()
    inverter_name_combobox['values'] = Inverter_name

def update_panel(*args):
    inverter_wattage_combobox.set('')
    i = panel_name_combobox.current()
    panel_wattage_combobox['values'] = Solar_Panel_Wattage[i]


def inverter_wattage_selection(*args):
    new_list = []
    global InverterPrice
    list = []
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
    index = ((inverter_name_combobox.current() + 2) * 2)-1
    if inverter_type_combobox == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
    else:
        for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
            list.append(x)
    list.pop(0)
    if inverter_wattage_combobox.current() >= 0 and inverter_wattage_combobox.current() <= 20:
        index2 = int(inverter_wattage_combobox.current())
        InverterPrice = list[index2]


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
Reffered_combobox = ttk.Combobox(user_info_frame, values=["Madam Greevo", "Engr Sajjad", "Engr Shaban", "Engr Abid"])
Reffered_label.grid(row=generalrow, column=3)
Reffered_combobox.grid(row=generalrowentry, column=3)

sel = tkinter.StringVar(user_info_frame)
sel3 = tkinter.StringVar(user_info_frame)
sel4 = tkinter.StringVar(user_info_frame)

inverter_type_label = tkinter.Label(user_info_frame, text="Inverter Type")
inverter_type_combobox = ttk.Combobox(user_info_frame, values=["Grid Tie", "Hybrid"], textvariable=sel)
inverter_type_label.grid(row=inverterrow, column=0)
inverter_type_combobox.grid(row=inverterrowentry, column=0)

sel.trace('w',inverter_selection)
inverter_name_label = tkinter.Label(user_info_frame, text="Inverter Name")
inverter_name_combobox = ttk.Combobox(user_info_frame, values=Inverter_name)
inverter_name_label.grid(row=inverterrow, column=1)
inverter_name_combobox.grid(row=inverterrowentry, column=1)

inverter_wattage_label = tkinter.Label(user_info_frame, text="Wattage")
inverter_wattage_combobox = ttk.Combobox(user_info_frame, values=Inverter_wattage, textvariable=sel3)
inverter_wattage_label.grid(row=inverterrow, column=2)
inverter_wattage_combobox.grid(row=inverterrowentry, column=2)

sel3.trace('w',inverter_wattage_selection)
sel2 = tkinter.StringVar(user_info_frame)

panel_name_label = tkinter.Label(user_info_frame, text="Panel Name")
panel_name_combobox = ttk.Combobox(user_info_frame, values=Solar_Panels_Names, textvariable=sel2)
panel_name_label.grid(row=panelrow, column=0)
panel_name_combobox.grid(row=panelrowentry, column=0)

panel_wattage_label = tkinter.Label(user_info_frame, text="Panel Wattage")
panel_wattage_combobox = ttk.Combobox(user_info_frame, values=Solar_Panel_Wattage)
panel_wattage_label.grid(row=panelrow, column=1)
panel_wattage_combobox.grid(row=panelrowentry, column=1)

sel2.trace('w',update_panel)

number_of_panels_label = tkinter.Label(user_info_frame, text="No. of Panels")
number_of_panels_spinbox = tkinter.Spinbox(user_info_frame, from_=0, to=1000)
number_of_panels_label.grid(row=panelrow, column=2)
number_of_panels_spinbox.grid(row=panelrowentry, column=2)

structure_type = tkinter.Label(user_info_frame, text="Structure Type")
structure_type_combobox = ttk.Combobox(user_info_frame, values=["Normal", "Raised"])
structure_type.grid(row=pvrow, column=0)
structure_type_combobox.grid(row=pvrowentry, column=0)


pv_balance_label = tkinter.Label(user_info_frame, text="PV Balance")
pv_balance_combobox = ttk.Entry(user_info_frame)
pv_balance_label.grid(row=pvrow, column=1, padx=10)
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