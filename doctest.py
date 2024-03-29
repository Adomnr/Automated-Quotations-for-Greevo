from spire.doc import *

def document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP, Inverter_Watt,
                 Inverter2_Watt, Number_of_Panels,
                 Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                 TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords, Inverter_Name,
                 Inverter2_Name,
                 Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces, Number_of_inverters,
                 panelwattage,
                 Carriage_Cost, Installation, Foundation, Earthing):
    # Create a Document object
    document = Document()

    # Load the template
    document.LoadFromFile(template_file)
    NOS = int(int(Number_of_Panels)/2)

    print("was here")
    batteryprice = int(BatteryPrice) * int(BatteryPieces)

    panelname = Name_of_Panels[:5-6]

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

    if Earthing == 0:
        earthing_symbol = "N/A"
    else:
        earthing_symbol = "01 Set"


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
                    '[pn]': panelname,
                    '[pw]': str(panelwattage),
                    '[in]': str(Inverter_Name),
                    '[in2]': str(Inverter2_Name),
                    '[fs]': str(foundation_symbol),
                    '[cs]': str(carriage_symbol),
                    '[es]': str(earthing_symbol),
                    '[is]': str(installation_symbol)
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