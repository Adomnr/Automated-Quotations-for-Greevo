from spire.doc import *
import os
from datetime import datetime

def new_folder(filename):
    # Get the current date
    current_date = datetime.now()

    # Define the filename
    filename = "example.txt"  # Replace "example.txt" with your actual filename

    # Define the directory structure
    directory_year_month = current_date.strftime("%B %Y")
    directory_day = current_date.strftime("%B %d")

    # Create the directories if they don't exist
    os.makedirs(os.path.join(directory_year_month, directory_day), exist_ok=True)

    # Combine the directory paths
    folder_path = os.path.join(directory_year_month, directory_day)

    # Path to the new file
    file_path = os.path.join(folder_path, filename)

    # Create the new file
    with open(file_path, "w") as file:
        # You can write to the file here if needed
        pass

    print(f"New file '{filename}' created in folder '{folder_path}'")


def document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                    Inverter_Watt, Inverter2_Watt, Number_of_Panels,
                    Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                    TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords,
                    Inverter_Name, Inverter2_Name,
                    Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces, BatterySpecs,
                    Number_of_inverters, panelwattage,
                    Carriage_Cost, Installation, Foundation, Earthing_val,
                    panelprice, structure_rate_normal, structure_rate_raised, InverterPrice,
                    pv_balance, Structure_TYP, advancepanelnames, advanceinverternames,advancecablename, NumberofDifferentInverters):
    # Create a Document object
    document = Document()

    # Load the template
    document.LoadFromFile(template_file)
    NOS = int(int(Number_of_Panels)/2)

    print("was here")
    batteryprice = int(BatteryPrice) * int(BatteryPieces)

    panelname = Name_of_Panels[:-6]

    TCN = ('{:,}'.format(int(TotalCostNormal)))
    TCR = ('{:,}'.format(int(TotalCostRaised)))
    TCNNI = ('{:,}'.format(int(TotalCostNormalNNI)))
    TCRNI = ('{:,}'.format(int(TotalCostRaisedNNI)))

    Net_Metering_cost = ('{:,}'.format(int(Net_Metering)))

    net_metering_symbol = ""

    if int(Net_Metering) == 0:
        net_metering_symbol = "N/A"
    else:
        net_metering_symbol = "01 Job"

    inverter_warranty_year = ""

    if str(Inverter_Name) == "FoxESS":
        inverter_warranty_year = "10 years replacement"
    else:
        inverter_warranty_year = "5 years"

    carriage_symbol = ""
    if int(Carriage_Cost) == 0:
        carriage_symbol = "N/A"
    else:
        carriage_symbol = "01 Job"

    foundation_symbol = ""
    if int(Foundation) == 0:
        foundation_symbol = "N/A"
    else:
        foundation_symbol = "01 Job"

    if batteryprice == 0:
        str(batteryprice)
        batteryprice = "N/A"

    if Installation == 0:
        installation_symbol = "N/A"
    else:
        installation_symbol = "01 Job"

    if Earthing_val == 0:
        earthing_symbol = "N/A"
    else:
        earthing_symbol = "01 Set"

    TotalPanelPrice = int(panelprice) * int(panelwattage) * int(Number_of_Panels)

    structure_rate = 0
    if Structure_TYP == "Normal":
        structure_rate = int(structure_rate_normal)
    else:
        structure_rate = int(structure_rate_raised)

    # Store the placeholders and new strings in a dictionary
    dictionary = {
                    '[uid]': str(UniqueID),
                    '[sw]': str(SystemSize),
                    '[nop]': str(Number_of_Panels),
                    '[nos]': str(NOS),
                    '[noi]': str(Number_of_inverters),
                    '[iw]': str(Inverter_Watt),
                    '[iwy]': str(inverter_warranty_year),
                    '[calculated_price]': str(TCR),
                    '[calculated_price_normal]': str(TCN),
                    '[calculated_price_NNI]': str(TCRNI),
                    '[calculated_price_normal_NNI]': str(TCNNI),
                    '[value_in_words]': str(valueinwords),
                    '[bn]': str(BatteryName),
                    '[bp]': str(batteryprice),
                    '[bv]': str(BatteryPieces),
                    '[nmc]': str(Net_Metering_cost),
                    '[nms]': str(net_metering_symbol),
                    '[pn]': str(panelname),
                    '[pw]': str(panelwattage),
                    '[in]': str(Inverter_Name),
                    '[in2]': str(Inverter2_Name),
                    '[fs]': str(foundation_symbol),
                    '[cs]': str(carriage_symbol),
                    '[es]': str(earthing_symbol),
                    '[is]': str(installation_symbol),
                    '[pp]': str(TotalPanelPrice),
                    '[sp]': str(structure_rate),
                    '[pvp]': str(pv_balance),
                    '[ep]': str(Earthing_val),
                    '[isp]': str(Installation),
                    '[cp]': str(Carriage_Cost),
                    '[fp]': str(Foundation),
                    '[apn]': str(advancepanelnames),
                    '[ain]': str(advanceinverternames),
                    '[acn]': str(advancecablename)
                }

    # Loop through the items in the dictionary
    for key, value in dictionary.items():
        # Replace a placeholder (key) with a new string (value)
        document.Replace(key, value, False, True)

    filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
        ClientLocation) + " Quotation" + str(UniqueID)
    print(filename)
    savename = "output/"+filename+".docx"

    # Save the resulting document
    document.SaveToFile(savename, FileFormat.Docx2016)
    document.Close()