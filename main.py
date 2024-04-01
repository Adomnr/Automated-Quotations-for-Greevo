# Importing Libraries
import tkinter
from tkinter import ttk
from tkinter import messagebox
from readwritegooglesheets import *
# For Value in words.
import inflect
from datetime import datetime
from doctest import *
import uuid
import time
from trademarkremover import *
import math
from pathlib import Path

Inverter_name, Inverter_wattage, Inverter2_name, Inverter2_wattage, Inverter3_name, Inverter3_wattage = [], [], [], [], [], []

Number_of_Inverter1, Number_of_Inverter2 = 0, 0

template_file = ""
InverterPrice, Inverter2Price = 0, 0

Number_of_Panels, panelprice, panelwattage, panel_structure_rate = 0, 0, 0, 0

TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI = 0, 0, 0, 0

UniqueID = 0
serialNumber = 0

structure_rate_normal, structure_rate_raised = 6500, 20

valueinwords = ""

generalrow, generalrowentry, inverterselectionrow, inverterselectionrowentry, inverterrow, inverterrowentry, inverter2row, inverter2rowentry, inverter3row, inverter3rowentry = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9

panelrow, panelrowentry, batteryrow, batteryentryrow, cinrow, cinrowentry = 10, 11, 12, 13, 14, 15


def capitalize_first_character_in_each_word(sentence):
    capitalized_sentence = sentence.title()
    return capitalized_sentence


def GetSerialNumber():
    global serialNumber
    index = []
    for x in Customer_Data_Sheet.col_values(1):
        index.append(x)
    serialNumber = int(index[-1])


def override_rates(*args):
    newwindow = tkinter.Tk()
    newwindow.title("Rates Overrider")

    newframe = tkinter.Frame(newwindow)
    newframe.pack()

    # Saving User Info
    Rates_Frames = tkinter.LabelFrame(newframe, text="Inverter and Panel Rates")
    Rates_Frames.grid(row=0, column=0, padx=20, pady=10)


    Inverter_Rate_Label = tkinter.Label(Rates_Frames, text="Inverter 1 Rate")
    Inverter_Rate_Entry = tkinter.Entry(Rates_Frames)
    Inverter_Rate_Label.grid(row=0, column=0, padx=20, pady=10)
    Inverter_Rate_Entry.grid(row=1, column=0, padx=20, pady=10)

    if Inverter_Rate_Entry.get() == "":
        Inverter_Rate_Entry.insert(0, str(InverterPrice))

    PanelRateLabel = tkinter.Label(Rates_Frames, text="Panel Rate")
    PanelRateValue = tkinter.Entry(Rates_Frames)
    PanelRateLabel.grid(row=2, column=0, padx=20, pady=10)
    PanelRateValue.grid(row=3, column=0, padx=20, pady=10)

    if PanelRateValue.get() == "":
        PanelRateValue.insert(0,str(panelprice))

    structure_rate_normal_label = tkinter.Label(Rates_Frames, text="Normal Structure Rate")
    structure_rate_normal_entry = tkinter.Entry(Rates_Frames)
    structure_rate_normal_label.grid(row=4, column=0, padx=20, pady=10)
    structure_rate_normal_entry.grid(row=5, column=0, padx=20, pady=10)

    if structure_rate_normal_entry.get() == "":
        structure_rate_normal_entry.insert(0, "6500")

    structure_rate_raised_label = tkinter.Label(Rates_Frames, text="Raised Structure Rate")
    structure_rate_raised_entry = tkinter.Entry(Rates_Frames)
    structure_rate_raised_label.grid(row=6, column=0, padx=20, pady=10)
    structure_rate_raised_entry.grid(row=7, column=0, padx=20, pady=10)

    if structure_rate_raised_entry.get() == "":
        structure_rate_raised_entry.insert(0, "20")

    def update_rates(*args):
        update_panel_price(PanelRateValue.get())
        update_inverter_price(Inverter_Rate_Entry.get())
        update_structure_rate_normal_rates(structure_rate_normal_entry.get())
        update_structure_rate_raised_rates(structure_rate_raised_entry.get())
        print(panelprice)
        print(InverterPrice)
        print(structure_rate_raised)
        print(structure_rate_normal)
        newwindow.destroy()

    ChangeButton = tkinter.Button(newframe, text="Override", command=update_rates)
    ChangeButton.grid(row=4,column=0)


def update_structure_rate_normal_rates(structurepricenormal):
    global structure_rate_normal
    structure_rate_normal = int(structurepricenormal)

def update_structure_rate_raised_rates(structurepriceraised):
    global structure_rate_raised
    structure_rate_raised = int(structurepriceraised)

def update_panel_price(PanelRateValue):
    global panelprice
    panelprice = int(PanelRateValue)


def update_inverter_price(Inverter_Rate_Entry):
    global InverterPrice
    InverterPrice = int(Inverter_Rate_Entry)


def enter_data():
    global valueinwords
    accepted = accept_var.get()
    total_cost_calculator()
    UniqueID = generate_unique_id()
    if Quotation_type_combobox.get() == "General Net Metering Not Included" or Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
        if structure_type_combobox.get() == "Raised":
            valueinwords = capitalize_first_character_in_each_word(convert_to_words(TotalCostRaised))
        else:
            valueinwords = capitalize_first_character_in_each_word(convert_to_words(TotalCostNormal))
    if Quotation_type_combobox.get() == "General Net Metering Included" or Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
        if structure_type_combobox.get() == "Raised":
            valueinwords = capitalize_first_character_in_each_word(convert_to_words(TotalCostRaisedNNI))
        else:
            valueinwords = capitalize_first_character_in_each_word(convert_to_words(TotalCostNormalNNI))
    print("1")
    print(valueinwords)
    if accepted == "Accepted":
        # User info
        print("2")
        SystemSize = System_Size_combobox.get()
        ClientName = Client_Name_combobox.get()
        ClientLocation = Client_Location_combobox.get()
        ReferredBy = Reffered_combobox.get()
        if SystemSize and ClientName and ClientLocation and ReferredBy:
            print("3")
            Inverter_TYP = inverter_type_combobox.get()
            Inverter2_TYP = inverter2_type_combobox.get()
            Inverter_Name = inverter_name_combobox.get()
            Inverter2_Name = inverter2_name_combobox.get()
            Inverter_Watt = inverter_wattage_combobox.get()
            Inverter2_Watt = inverter2_wattage_combobox.get()
            Number_of_Inverter1 = inverter_number_Entry.get()
            Number_of_Inverter2 = inverter2_number_Entry.get()
            Number_of_inverters = int(Number_of_Inverter1) + int(Number_of_Inverter2)
            if Inverter_TYP and Inverter_Name and Inverter_Watt:
                print("4")
                if inverter_selection_combobox.get() == "2":
                    print("5")
                    if Inverter2_TYP and Inverter2_Name and Inverter2_Watt:
                        print("6")
                        Name_of_Panels = panel_name_combobox.get()
                        Structure_Type = structure_type_combobox.get()
                        pv_balance = pv_balance_combobox.get()
                        if Name_of_Panels and Structure_Type and pv_balance:
                            print("7")
                            BatteryPrice = Battery_Price_Entry.get()
                            BatteryName = Battery_Name_Entry.get()
                            BatteryPieces = Number_of_Batteries_Entry.get()
                            if Inverter_TYP == "Hybrid" or Inverter2_TYP == "Hybrid":
                                print("8")
                                if BatteryPrice and BatteryName and BatteryPieces:
                                    print("9")
                                    Earthing_val = Earthing_entry.get()
                                    Foundation = foundation_work_entry.get()
                                    Carriage_Cost = carriage_entry.get()
                                    Installation = installation_entry.get()
                                    Net_Metering = net_metering_entry.get()
                                    if Carriage_Cost and Installation and Net_Metering and Foundation and Earthing_val:
                                        print("10")
                                        print("System Size: ", SystemSize, "Client Name: ", ClientName,
                                              "Client Location: ", ClientLocation)
                                        print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP,
                                              "Inverter Name: ", Inverter_Name)
                                        print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels,
                                              "Panel Price: ", panelprice)
                                        print("Inverter 1 Price", InverterPrice, "Inverter 2 Price", Inverter2Price)
                                        print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                                        print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost,
                                              "Installation Cost: ", Installation)
                                        print("Net Metering: ", Net_Metering, "Template File", template_file,
                                              "Inverter Price", InverterPrice)
                                        print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ",
                                              TotalCostRaised, "Unique ID: ", UniqueID)
                                        print("------------------------------------------")

                                        document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                                                         Inverter_Watt,
                                                         Inverter2_Watt, Number_of_Panels,
                                                         Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                                         TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords,
                                                         Inverter_Name,
                                                         Inverter2_Name,
                                                         Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces,
                                                         Number_of_inverters,
                                                         panelwattage,
                                                         Carriage_Cost, Installation, Foundation, Earthing_val)

                                        filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                                            ClientLocation) + " Quotation" + str(UniqueID)
                                        tradeMarkRemover(filename)

                                        record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy,
                                                    Inverter_TYP,
                                                    Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice,
                                                    Number_of_Panels,
                                                    Structure_Type, pv_balance, Carriage_Cost, Installation,
                                                    Net_Metering,
                                                    template_file, InverterPrice, TotalCostNormal, TotalCostRaised,
                                                    TotalCostNormalNNI,
                                                    TotalCostRaisedNNI)
                                    else:
                                        tkinter.messagebox.showwarning(title="Error", message="Enter Carriage, Installation and Net Metering Cost")
                                else:
                                    tkinter.messagebox.showwarning(title="Error", message="Battery Name, Rate and Pieces")
                            else:
                                Foundation = foundation_work_entry.get()
                                Carriage_Cost = carriage_entry.get()
                                Installation = installation_entry.get()
                                Net_Metering = net_metering_entry.get()
                                Earthing_val = Earthing_entry.get()
                                if Carriage_Cost and Installation and Net_Metering:
                                    print("System Size: ", SystemSize, "Client Name: ", ClientName, "Client Location: ",
                                        ClientLocation)
                                    print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP, "Inverter Name: ",
                                        Inverter_Name)
                                    print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels,
                                        "Panel Price: ", panelprice)
                                    print("Inverter 1 Price", InverterPrice, "Inverter 2 Price", Inverter2Price)
                                    print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                                    print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost, "Installation Cost: ",
                                        Installation)
                                    print("Net Metering: ", Net_Metering, "Template File", template_file, "Inverter Price",
                                        InverterPrice)
                                    print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ", TotalCostRaised,
                                        "Unique ID: ", UniqueID)
                                    print("------------------------------------------")

                                    document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                                                         Inverter_Watt,
                                                         Inverter2_Watt, Number_of_Panels,
                                                         Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                                         TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords, Inverter_Name,
                                                         Inverter2_Name,
                                                         Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces,
                                                         Number_of_inverters,
                                                         panelwattage,
                                                         Carriage_Cost, Installation, Foundation, Earthing_val)

                                    filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                                        ClientLocation) + " Quotation" + str(UniqueID)
                                    tradeMarkRemover(filename)

                                    record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy, Inverter_TYP,
                                                    Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice, Number_of_Panels,
                                                    Structure_Type, pv_balance, Carriage_Cost, Installation, Net_Metering,
                                                    template_file, InverterPrice, TotalCostNormal, TotalCostRaised,
                                                    TotalCostNormalNNI,
                                                    TotalCostRaisedNNI)
                                else:
                                    tkinter.messagebox.showwarning(title="Error", message="Enter Carriage, Installation and Net Metering Cost")
                        else:
                            tkinter.messagebox.showwarning(title="Error", message="Enter Structure Type and PV Balance.")
                else:
                    print("5")
                    print("6")
                    Name_of_Panels = panel_name_combobox.get()
                    Structure_Type = structure_type_combobox.get()
                    pv_balance = pv_balance_combobox.get()
                    if Name_of_Panels and Structure_Type and pv_balance:
                        print("7")
                        BatteryPrice = Battery_Price_Entry.get()
                        BatteryName = Battery_Name_Entry.get()
                        BatteryPieces = Number_of_Batteries_Entry.get()
                        if Inverter_TYP == "Hybrid":
                            print("8")
                            if BatteryPrice and BatteryName and BatteryPieces:
                                print("9")
                                Earthing_val = Earthing_entry.get()
                                Foundation = foundation_work_entry.get()
                                Carriage_Cost = carriage_entry.get()
                                Installation = installation_entry.get()
                                Net_Metering = net_metering_entry.get()
                                if Carriage_Cost and Installation and Net_Metering and Foundation and Earthing_val:
                                    print("10")
                                    print("System Size: ", SystemSize, "Client Name: ", ClientName,                                            "Client Location: ", ClientLocation)
                                    print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP,
                                        "Inverter Name: ", Inverter_Name)
                                    print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels,
                                        "Panel Price: ", panelprice)
                                    print("Inverter 1 Price", InverterPrice, "Inverter 2 Price", Inverter2Price)
                                    print("No of Panels: ", Number_of_Panels, "Structure Type: ",
                                        Structure_Type)
                                    print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost,
                                        "Installation Cost: ", Installation)
                                    print("Net Metering: ", Net_Metering, "Template File", template_file,                                            "Inverter Price", InverterPrice)
                                    print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ",
                                        TotalCostRaised, "Unique ID: ", UniqueID)
                                    print("------------------------------------------")

                                    document_creater(UniqueID, ClientName, ClientLocation, SystemSize,
                                        Inverter_TYP, Inverter_Watt, Inverter2_Watt, Number_of_Panels,
                                        Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                        TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords, Inverter_Name,
                                        Inverter2_Name, Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces,
                                        Number_of_inverters, panelwattage,
                                        Carriage_Cost, Installation, Foundation, Earthing_val)

                                    filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                                        ClientLocation) + " Quotation" + str(UniqueID)
                                    tradeMarkRemover(filename)

                                    record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy,
                                    Inverter_TYP, Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice,
                                    Number_of_Panels, Structure_Type, pv_balance, Carriage_Cost, Installation,
                                    Net_Metering, template_file, InverterPrice, TotalCostNormal, TotalCostRaised,
                                    TotalCostNormalNNI, TotalCostRaisedNNI)
                                else:
                                    tkinter.messagebox.showwarning(title="Error", message="Enter Carriage, Installation and Net Metering Cost")
                        else:
                            Foundation = foundation_work_entry.get()
                            Carriage_Cost = carriage_entry.get()
                            Installation = installation_entry.get()
                            Net_Metering = net_metering_entry.get()
                            Earthing_val = Earthing_entry.get()
                            if Carriage_Cost and Installation and Net_Metering:
                                print("System Size: ", SystemSize, "Client Name: ", ClientName,
                                    "Client Location: ", ClientLocation)
                                print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP,
                                    "Inverter Name: ", Inverter_Name)
                                print("Inver4ter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels,
                                    "Panel Price: ", panelprice)
                                print("Inverter 1 Price", InverterPrice, "Inverter 2 Price", Inverter2Price)
                                print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                                print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost,
                                    "Installation Cost: ", Installation)
                                print("Net Metering: ", Net_Metering, "Template File", template_file,
                                    "Inverter Price", InverterPrice)
                                print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ",
                                TotalCostRaised, "Unique ID: ", UniqueID)
                                print("------------------------------------------")

                                document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                                                     Inverter_Watt,
                                                     Inverter2_Watt, Number_of_Panels,
                                                     Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                                     TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords,
                                                     Inverter_Name,
                                                     Inverter2_Name,
                                                     Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces,
                                                     Number_of_inverters,
                                                    panelwattage,
                                                     Carriage_Cost, Installation, Foundation, Earthing_val)

                                filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                                    ClientLocation) + " Quotation" + str(UniqueID)
                                tradeMarkRemover(filename)

                                record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy,
                                            Inverter_TYP, Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice,
                                            Number_of_Panels,  Structure_Type, pv_balance, Carriage_Cost, Installation,
                                            Net_Metering, template_file, InverterPrice, TotalCostNormal, TotalCostRaised,
                                            TotalCostNormalNNI, TotalCostRaisedNNI)
                            else:
                                tkinter.messagebox.showwarning(title="Error", message="Enter Carriage, Installation and Net Metering Cost")
                    else:
                        tkinter.messagebox.showwarning(title="Error", message="Enter Structure Type and PV Balance.")
            else:
                tkinter.messagebox.showwarning(title="Error", message="Enter All box of Inverters.")
        else:
            tkinter.messagebox.showwarning(title="Error", message="Enter All boxes of Client Information.")
    else:
        tkinter.messagebox.showwarning(title="Error", message="You have not accepted the terms")


# This updates the template which is going to be subsituited the placeholder into GTGN Grid Tied General Normal

def round_up_to_nearest_thousand(number):
    if number % 1000 == 0:
        return number
    else:
        return ((number // 1000) + 1) * 1000


def total_cost_calculator(*args):
    global TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI
    pv_balance = pv_balance_combobox.get()
    Carriage_Cost = carriage_entry.get()
    Installation = installation_entry.get()
    Net_Metering = net_metering_entry.get()
    Foundation_Work = foundation_work_entry.get()
    Number_of_Inverter1 = inverter_number_Entry.get()
    Number_of_Inverter2 = inverter2_number_Entry.get()
    print(InverterPrice, "  ", Number_of_Inverter1, "  ", Inverter2Price, "  ", Number_of_Inverter2, "  ",
          Number_of_Panels, "  ", panelwattage, "  ", panelprice, "  ", pv_balance, "  ", Carriage_Cost, "  ",
          Installation, "  ", Net_Metering, "  ", Foundation_Work)
    TotalCostRaised = round_up_to_nearest_thousand(
            (int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) +
            (int(structure_rate_raised) * int(Number_of_Panels) * int(panelwattage)) + (int(panelwattage) * int(panelprice) *
            int(Number_of_Panels)) + int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Net_Metering) + int(Foundation_Work))
    print(TotalCostRaised)
    TotalCostRaisedNNI = round_up_to_nearest_thousand((int(InverterPrice) * int(Number_of_Inverter1)) +
            (int(Inverter2Price) * int(Number_of_Inverter2)) + (int(structure_rate_raised) *
            int(Number_of_Panels) * int(panelwattage)) + (int(panelwattage) * int(panelprice) * int(Number_of_Panels))
            + int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Foundation_Work))
    print(TotalCostRaisedNNI)
    TotalCostNormal = round_up_to_nearest_thousand(
            (int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) +
            int((int(structure_rate_normal) * (int(Number_of_Panels) / 2))) + (int(panelwattage) * int(panelprice) * int(Number_of_Panels)) +
            int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Net_Metering) + int(Foundation_Work))
    print(TotalCostNormal)
    TotalCostNormalNNI = round_up_to_nearest_thousand((int(InverterPrice) * int(Number_of_Inverter1)) +
            (int(Inverter2Price) * int(Number_of_Inverter2)) + int((int(structure_rate_normal) * (int(Number_of_Panels) / 2))) +
            (int(panelwattage) * int(panelprice) * int(Number_of_Panels)) + int(pv_balance) + int(Carriage_Cost) +
            int(Installation) + int(Foundation_Work))
    print(TotalCostNormalNNI)


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
    unique_id = int(f"{timestamp:03d}{random_component:03d}") % 10000000
    return unique_id


def record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy, Inverter_TYP, Inverter_Name,
                Inverter_Watt, Name_of_Panels, panelprice, Number_of_Panels, Structure_Type, pv_balance, Carriage_Cost,
                Installation, Net_Metering, template_file, InverterPrice, TotalCostNormal, TotalCostRaised,
                TotalCostNormalNNI, TotalCostRaisedNNI):
    serial_list = []
    current_date_time = datetime.now()
    current_date_time = str(current_date_time)
    for x in Customer_Data_Sheet.col_values(1):
        serial_list.append(x)
    SerialNumber = int(serial_list[-1]) + 1
    if SerialNumber == 0:
        SerialNumber = 1
    Customer_Data_Sheet.append_row([SerialNumber, current_date_time, UniqueID, SystemSize, ClientName, ClientLocation,
                                    ReferredBy, Inverter_TYP, Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice,
                                    Number_of_Panels, Structure_Type, pv_balance, Carriage_Cost, Installation,
                                    Net_Metering,
                                    InverterPrice, TotalCostNormal, TotalCostRaised, TotalCostNormalNNI,
                                    TotalCostRaisedNNI])


def update_template_type(*args):
    home_dir = Path.home()
    print(home_dir)
    global template_file
    if inverter_selection_combobox.get() == "1":
        if inverter_type_combobox.get() == "Grid Tie":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRNISPI_Template.docx")
        if inverter_type_combobox.get() == "Hybrid":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGNSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_Template.docx")
    if inverter_selection_combobox.get() == "2":
        if inverter_type_combobox.get() == "Grid Tie" and inverter2_type_combobox.get() == "Hybrid":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_WHI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_WHI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WHI_Template.docx")
        if inverter_type_combobox.get() == "Grid Tie" and inverter2_type_combobox.get() == "Grid Tie":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_WGTI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WGTI_Template.docx")
        if inverter_type_combobox.get() == "Hybrid" and inverter2_type_combobox.get() == "Grid Tie":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGNSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_WGTI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_WGTI_Template.docx")
        if inverter_type_combobox.get() == "Hybrid" and inverter2_type_combobox.get() == "Hybrid":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_WHI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_WHI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_WHI_Template.docx")
    print(template_file)


def inverter_selection(*args):
    battery_area_selection()
    inverter_name_combobox.set('')
    inverter_wattage_combobox.set('')
    if inverter_type_combobox.get() == "Grid Tie":
        Inverter_name = Grid_Tie_Inverter_Names.copy()
    else:
        Inverter_name = Hybrid_Inverter_Names.copy()
    inverter_name_combobox['values'] = Inverter_name


def inverter2_selection(*args):
    battery_area_selection()
    inverter2_name_combobox.set('')
    inverter2_wattage_combobox.set('')
    if inverter2_type_combobox.get() == "Grid Tie":
        Inverter2_name = Grid_Tie_Inverter_Names.copy()
    else:
        Inverter2_name = Hybrid_Inverter_Names.copy()
    inverter2_name_combobox['values'] = Inverter2_name


def inverter3_selection(*args):
    battery_area_selection()
    inverter3_name_combobox.set('')
    inverter3_wattage_combobox.set('')
    if inverter3_type_combobox.get() == "Grid Tie":
        Inverter3_name = Grid_Tie_Inverter_Names.copy()
    else:
        Inverter3_name = Hybrid_Inverter_Names.copy()
    inverter3_name_combobox['values'] = Inverter3_name


def update_panel(*args):
    global panelprice
    global panelwattage
    i = panel_name_combobox.current()
    panelprice = Solar_Panel_Price[i]
    print(panelprice)
    for x in Solar_Panel_Wattage:
        panelwattage = Solar_Panel_Wattage[i]
    PanelWattageInt = int(panelwattage)
    print(panelwattage)
    Panels_Numbers(PanelWattageInt)


def Panels_Numbers(panelwattage):
    global Number_of_Panels
    SZ = int(System_Size_combobox.get()) * 1000
    Number_of_Panels = SZ / int(panelwattage)
    Number_of_Panels = math.ceil(Number_of_Panels)
    print(Number_of_Panels)


def inverter_wattage_selection(*args):
    inverter_wattage_combobox.set('')
    new_list = []
    if inverter_type_combobox.get() == "Grid Tie":
        if inverter_name_combobox.get() != "":
            index = (inverter_name_combobox.current() + 1) * 2
            for x in Grid_Tied_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter_wattage_combobox.config(values=new_list)
    else:
        if inverter_name_combobox.get() != "":
            index = (inverter_name_combobox.current() + 1) * 2
            for x in Hybrid_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter_wattage_combobox.config(values=new_list)


def inverter2_wattage_selection(*args):
    inverter2_wattage_combobox.set('')
    new_list = []
    if inverter2_type_combobox.get() == "Grid Tie":
        if inverter2_name_combobox.get() != "":
            index = (inverter2_name_combobox.current() + 1) * 2
            for x in Grid_Tied_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter2_wattage_combobox.config(values=new_list)
    else:
        if inverter2_name_combobox.get() != "":
            index = (inverter2_name_combobox.current() + 1) * 2
            for x in Hybrid_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter2_wattage_combobox.config(values=new_list)


def inverter3_wattage_selection(*args):
    inverter3_wattage_combobox.set('')
    new_list = []
    if inverter3_type_combobox.get() == "Grid Tie":
        if inverter3_name_combobox.get() != "":
            index = (inverter3_name_combobox.current() + 1) * 2
            for x in Grid_Tied_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter3_wattage_combobox.config(values=new_list)
    else:
        if inverter3_name_combobox.get() != "":
            index = (inverter3_name_combobox.current() + 1) * 2
            for x in Hybrid_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter3_wattage_combobox.config(values=new_list)


def inverter_price(*args):
    index = ((inverter_name_combobox.current() + 2) * 2) - 1
    list = []
    if inverter_type_combobox.get() == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
        list.pop(0)
        while ("" in list):
            list.remove("")
    else:
        if inverter_type_combobox.get() == "Hybrid":
            for x in Hybrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print(list)
    if inverter_wattage_combobox.current() >= 0 and inverter_wattage_combobox.current() <= 20:
        index2 = int(inverter_wattage_combobox.current())
        global InverterPrice
        InverterPrice = list[index2]
        print(InverterPrice)


def inverter2_price(*args):
    index = ((inverter2_name_combobox.current() + 2) * 2) - 1
    list = []
    if inverter2_type_combobox.get() == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
        list.pop(0)
        while ("" in list):
            list.remove("")
    else:
        if inverter2_type_combobox.get() == "Hybrid":
            for x in Hybrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print(list)
    if inverter2_wattage_combobox.current() >= 0 and inverter2_wattage_combobox.current() <= 20:
        index2 = int(inverter2_wattage_combobox.current())
        global Inverter2Price
        Inverter2Price = list[index2]
        print(Inverter2Price)


def inverter3_price(*args):
    index = ((inverter3_name_combobox.current() + 2) * 2) - 1
    list = []
    if inverter3_type_combobox.get() == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
        list.pop(0)
        while ("" in list):
            list.remove("")
    else:
        if inverter3_type_combobox.get() == "Hybrid":
            for x in Hybrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print(list)
    if inverter3_wattage_combobox.current() >= 0 and inverter3_wattage_combobox.current() <= 20:
        index2 = int(inverter3_wattage_combobox.current())
        global Inverter3Price
        Inverter3Price = list[index2]
        print(Inverter3Price)


def inverter_number_selection(*args):
    if inverter_selection_combobox.get() == '1':
        inverter2_frame.grid_remove()
        inverter3_frame.grid_remove()
    else:
        if inverter_selection_combobox.get() == '2':
            inverter2_frame.grid()
            if inverter3_frame.winfo_ismapped():
                inverter3_frame.grid_remove()
        else:
            if inverter2_frame != inverter2_frame.grid():
                inverter2_frame.grid()
            inverter3_frame.grid()


def battery_area_selection(*args):
    if inverter_type_combobox.get() == "Grid Tie" or inverter2_type_combobox.get() == "Grid Tie" or inverter3_type_combobox.get() == "Grid Tie":
        battery_frame.grid_remove()
    if inverter_type_combobox.get() == "Hybrid" or inverter2_type_combobox.get() == "Hybrid" or inverter3_type_combobox.get() == "Hybrid":
        if not battery_frame.winfo_ismapped():
            battery_frame.grid()


# Inverter_name,name_of_solar_panels,wattage_of_solar_panels
window = tkinter.Tk()
window.title("Quotation Automation")

frame = tkinter.Frame(window)
frame.pack()

# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="Customer Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

systemtracker = tkinter.StringVar()

System_Size_label = tkinter.Label(user_info_frame, text="System Size")
System_Size_combobox = ttk.Entry(user_info_frame, textvariable=systemtracker)
System_Size_label.grid(row=generalrow, column=0)
System_Size_combobox.grid(row=generalrowentry, column=0)

systemtracker.trace('w', update_panel)

Client_Name_label = tkinter.Label(user_info_frame, text="Client Name")
Client_Name_combobox = ttk.Entry(user_info_frame)
Client_Name_label.grid(row=generalrow, column=1)
Client_Name_combobox.grid(row=generalrowentry, column=1)

Client_Location_label = tkinter.Label(user_info_frame, text="Location")
Client_Location_combobox = ttk.Entry(user_info_frame)
Client_Location_label.grid(row=generalrow, column=2)
Client_Location_combobox.grid(row=generalrowentry, column=2)

Reffered_label = tkinter.Label(user_info_frame, text="Referred By")
Reffered_combobox = ttk.Combobox(user_info_frame,
                                 values=["Madam Rafia", "Engr Sajjad", "Engr Shaban", "Engr Abid", "Engr Ammar Butt",
                                         "Engr Ubaid", "Sir Nabeel", "Engr Osama"])
Reffered_label.grid(row=generalrow, column=3)
Reffered_combobox.grid(row=generalrowentry, column=3)

inverter_selection_frame = tkinter.LabelFrame(frame, text="DIS")
inverter_selection_frame.grid(row=inverterselectionrow, column=0, sticky="news", padx=20, pady=15)

tracker_inverters = tkinter.StringVar(inverter_selection_frame)

inverter_selection_label = tkinter.Label(inverter_selection_frame, text="Number of Different Inverters")
inverter_selection_combobox = ttk.Combobox(inverter_selection_frame, values=['1', '2', '3'],
                                           textvariable=tracker_inverters)
inverter_selection_label.grid(row=inverterselectionrow, column=0)
inverter_selection_combobox.grid(row=inverterselectionrowentry, column=0)

tracker_inverters.trace('w', inverter_number_selection)

foundation_work = tkinter.Label(inverter_selection_frame, text="Foundation Work")
foundation_work_entry = ttk.Entry(inverter_selection_frame)
foundation_work.grid(row=inverterselectionrow, column=1)
foundation_work_entry.grid(row=inverterselectionrowentry, column=1)

if foundation_work_entry.get() == "":
    foundation_work_entry.insert(0, "0")

OverrideRate = ttk.Button(text="Override Rates")

inverter1_frame = tkinter.LabelFrame(frame, text="Inverter 1 Area")
inverter1_frame.grid(row=inverterrow, column=0, sticky="news", padx=20, pady=15)

sel = tkinter.StringVar(inverter1_frame)
sel3 = tkinter.StringVar(inverter1_frame)

sel5 = tkinter.StringVar(inverter1_frame)

inverter_type_label = tkinter.Label(inverter1_frame, text="Inverter 1 Type")
inverter_type_combobox = ttk.Combobox(inverter1_frame, values=["Grid Tie", "Hybrid"], textvariable=sel)
inverter_type_label.grid(row=inverterrow, column=0, padx=5)
inverter_type_combobox.grid(row=inverterrowentry, column=0, padx=5)

sel.trace('w', inverter_selection)

inverter_name_label = tkinter.Label(inverter1_frame, text="Inverter 1 Name")
inverter_name_combobox = ttk.Combobox(inverter1_frame, values=Inverter_name, textvariable=sel3)
inverter_name_label.grid(row=inverterrow, column=1, padx=5)
inverter_name_combobox.grid(row=inverterrowentry, column=1, padx=5)

inverter_wattage_label = tkinter.Label(inverter1_frame, text=" Inverter 1 Wattage")
inverter_wattage_combobox = ttk.Combobox(inverter1_frame, values=Inverter_wattage, textvariable=sel5)
inverter_wattage_label.grid(row=inverterrow, column=2, padx=5)
inverter_wattage_combobox.grid(row=inverterrowentry, column=2, padx=5)

sel3.trace('w', inverter_wattage_selection)

inverter_number_Label = tkinter.Label(inverter1_frame, text="No of Type 1 Inverters")
inverter_number_Entry = ttk.Entry(inverter1_frame)
inverter_number_Label.grid(row=inverterrow, column=3, padx=5)
inverter_number_Entry.grid(row=inverterrowentry, column=3, padx=5)

if inverter_number_Entry.get() == "":
    inverter_number_Entry.insert(0, '1')

inverter2_frame = tkinter.LabelFrame(frame, text="Inverter 2 Area")
inverter2_frame.grid(row=inverter2row, column=0, sticky="news", padx=20, pady=15)

sel6 = tkinter.StringVar(inverter2_frame)
sel7 = tkinter.StringVar(inverter2_frame)
sel8 = tkinter.StringVar(inverter2_frame)

inverter2_type_label = tkinter.Label(inverter2_frame, text="Inverter 2 Type")
inverter2_type_combobox = ttk.Combobox(inverter2_frame, values=["Grid Tie", "Hybrid"], textvariable=sel6)
inverter2_type_label.grid(row=inverter2row, column=0, padx=5)
inverter2_type_combobox.grid(row=inverter2rowentry, column=0, padx=5)

sel6.trace('w', inverter2_selection)

inverter2_name_label = tkinter.Label(inverter2_frame, text="Inverter 2 Name")
inverter2_name_combobox = ttk.Combobox(inverter2_frame, values=Inverter2_name, textvariable=sel7)
inverter2_name_label.grid(row=inverter2row, column=1, padx=5)
inverter2_name_combobox.grid(row=inverter2rowentry, column=1, padx=5)

inverter2_wattage_label = tkinter.Label(inverter2_frame, text=" Inverter 2 Wattage")
inverter2_wattage_combobox = ttk.Combobox(inverter2_frame, values=Inverter2_wattage, textvariable=sel8)
inverter2_wattage_label.grid(row=inverter2row, column=2, padx=5)
inverter2_wattage_combobox.grid(row=inverter2rowentry, column=2, padx=5)

inverter2_number_Label = tkinter.Label(inverter2_frame, text="No of Type 2 Inverters")
inverter2_number_Entry = ttk.Entry(inverter2_frame)
inverter2_number_Label.grid(row=inverter2row, column=3, padx=5)
inverter2_number_Entry.grid(row=inverter2rowentry, column=3, padx=5)

if inverter_selection_combobox.get() == '1':
    if inverter2_number_Entry.get() == "":
        inverter2_number_Entry.insert(0, '0')
else:
    if inverter2_number_Entry.get() == "":
        inverter2_number_Entry.insert(0, '1')

sel7.trace('w', inverter2_wattage_selection)

inverter3_frame = tkinter.LabelFrame(frame, text="Inverter 3 Area")
inverter3_frame.grid(row=inverter3row, column=0, sticky="news", padx=20, pady=15)

sel9 = tkinter.StringVar(inverter3_frame)
sel10 = tkinter.StringVar(inverter3_frame)
sel11 = tkinter.StringVar(inverter3_frame)

inverter3_type_label = tkinter.Label(inverter3_frame, text="Inverter 3 Type")
inverter3_type_combobox = ttk.Combobox(inverter3_frame, values=["Grid Tie", "Hybrid"], textvariable=sel9)
inverter3_type_label.grid(row=inverter3row, column=0, padx=5)
inverter3_type_combobox.grid(row=inverter3rowentry, column=0, padx=5)

sel9.trace('w', inverter3_selection)

inverter3_name_label = tkinter.Label(inverter3_frame, text="Inverter 3 Name")
inverter3_name_combobox = ttk.Combobox(inverter3_frame, values=Inverter2_name, textvariable=sel10)
inverter3_name_label.grid(row=inverter3row, column=1, padx=5)
inverter3_name_combobox.grid(row=inverter3rowentry, column=1, padx=5)

inverter3_wattage_label = tkinter.Label(inverter3_frame, text=" Inverter 3 Wattage")
inverter3_wattage_combobox = ttk.Combobox(inverter3_frame, values=Inverter2_wattage, textvariable=sel11)
inverter3_wattage_label.grid(row=inverter3row, column=2, padx=5)
inverter3_wattage_combobox.grid(row=inverter3rowentry, column=2, padx=5)

inverter3_number_Label = tkinter.Label(inverter3_frame, text="No of Type 3 Inverters")
inverter3_number_Entry = ttk.Entry(inverter3_frame)
inverter3_number_Label.grid(row=inverter3row, column=3, padx=5)
inverter3_number_Entry.grid(row=inverter3rowentry, column=3, padx=5)

if inverter_selection_combobox.get() == "3":
    if inverter3_number_Entry.get() == "":
        inverter3_number_Entry.insert(0, '1')
else:
    if inverter3_number_Entry.get() == "":
        inverter3_number_Entry.insert(0, '0')

sel10.trace('w', inverter2_wattage_selection)

if inverter_selection_combobox.get() == "":
    inverter_selection_combobox.set('1')

panel_frame = tkinter.LabelFrame(frame, text="Panel Area")
panel_frame.grid(row=panelrow, column=0, sticky="news", padx=20, pady=15)

sel2 = tkinter.StringVar(panel_frame)

panel_name_label = tkinter.Label(panel_frame, text="Panel Name")
panel_name_combobox = ttk.Combobox(panel_frame, values=Solar_Panels_Names, textvariable=sel2)
panel_name_label.grid(row=panelrow, column=0, padx=5)
panel_name_combobox.grid(row=panelrowentry, column=0, padx=5)

sel5.trace('w', inverter_price)
sel8.trace('w', inverter2_price)
sel11.trace('w', inverter3_price)
sel2.trace('w', update_panel)
sel4 = tkinter.StringVar(panel_frame)

pv_balance_label = tkinter.Label(panel_frame, text="PV Balance")
pv_balance_combobox = ttk.Entry(panel_frame)
pv_balance_label.grid(row=panelrow, column=1, padx=5)
pv_balance_combobox.grid(row=panelrowentry, column=1, padx=5)

structure_type = tkinter.Label(panel_frame, text="Structure Type")
structure_type_combobox = ttk.Combobox(panel_frame, values=["Normal", "Raised"], textvariable=sel4)
structure_type.grid(row=panelrow, column=2, padx=5)
structure_type_combobox.grid(row=panelrowentry, column=2, padx=5)

if structure_type_combobox.get() == "":
    structure_type_combobox.set("Normal")

update_tracker = tkinter.StringVar()

Quotation_type = tkinter.Label(panel_frame, text="Quotation Type")
Quotation_type_combobox = ttk.Combobox(panel_frame,
                                       values=["General Net Metering Included", "Specify Brand Net Metering Included",
                                               "General Net Metering Not Included",
                                               "Specify Brand Net Metering Not Included"], textvariable=update_tracker)
Quotation_type.grid(row=panelrow, column=3, padx=5)
Quotation_type_combobox.grid(row=panelrowentry, column=3, padx=5)

update_tracker.trace('w', update_template_type)
if Quotation_type_combobox.get() == "":
    Quotation_type_combobox.set("Net Metering Not Included")

sel4.trace('w', update_template_type)

battery_frame = tkinter.LabelFrame(frame, text="Battery Area")
battery_frame.grid(row=batteryrow, column=0, sticky="news", padx=20, pady=15)

Battery_Name_Label = tkinter.Label(battery_frame, text="Battery Name")
Battery_Name_Entry = ttk.Entry(battery_frame)
Battery_Name_Label.grid(row=batteryrow, column=0, padx=5)
Battery_Name_Entry.grid(row=batteryentryrow, column=0, padx=5)

if Battery_Name_Entry.get() == "":
    Battery_Name_Entry.insert(0, "Daewoo Deep Cycle")

Battery_Price_Label = tkinter.Label(battery_frame, text="Battery Price")
Battery_Price_Entry = ttk.Entry(battery_frame)
Battery_Price_Label.grid(row=batteryrow, column=1, padx=5)
Battery_Price_Entry.grid(row=batteryentryrow, column=1, padx=5)

if Battery_Price_Entry.get() == "":
    Battery_Price_Entry.insert(0, "45000")

Number_of_Batteries_Label = tkinter.Label(battery_frame, text="Number of Batteries")
Number_of_Batteries_Entry = ttk.Entry(battery_frame)
Number_of_Batteries_Label.grid(row=batteryrow, column=2, padx=5)
Number_of_Batteries_Entry.grid(row=batteryentryrow, column=2, padx=5)

if Number_of_Batteries_Entry.get() == "":
    Number_of_Batteries_Entry.insert(0, '4')

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving Course Info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=cinrow, column=0, sticky="news", padx=20, pady=10)

carriage = tkinter.Label(courses_frame, text="Carriage")
carriage_entry = ttk.Entry(courses_frame)
carriage.grid(row=cinrow, column=0)
carriage_entry.grid(row=cinrowentry, column=0)

if carriage_entry.get() == "":
    carriage_entry.insert(0, '0')

installation = tkinter.Label(courses_frame, text="Installation")
installation_entry = ttk.Entry(courses_frame)
installation.grid(row=cinrow, column=1)
installation_entry.grid(row=cinrowentry, column=1)

if installation_entry.get() == "":
    installation_entry.insert(0, '0')

net_metering = tkinter.Label(courses_frame, text="Net Metering")
net_metering_entry = ttk.Entry(courses_frame)
net_metering.grid(row=cinrow, column=2)
net_metering_entry.grid(row=cinrowentry, column=2)

if net_metering_entry.get() == "":
    net_metering_entry.insert(0, '0')

Earthing = tkinter.Label(courses_frame, text="Earthing")
Earthing_entry = ttk.Entry(courses_frame)
Earthing.grid(row=cinrow, column=3)
Earthing_entry.grid(row=cinrowentry, column=3)

if Earthing_entry.get() == "":
    Earthing_entry.insert(0, "0")

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=12, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Last Check")
terms_frame.grid(row=17, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="Checked All.", variable=accept_var, onvalue="Accepted",
                                  offvalue="Not Accepted")
terms_check.grid(row=18, column=0)

overrride_button = tkinter.Button(frame, text="Override Rates", command=override_rates)
overrride_button.grid(row=19, column=0, sticky="news", padx=20, pady=10)
# Button
button = tkinter.Button(frame, text="Enter data", command=enter_data)
button.grid(row=20, column=0, sticky="news", padx=20, pady=10)

window.mainloop()
