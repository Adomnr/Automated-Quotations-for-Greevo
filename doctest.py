from spire.doc import *
from spire.doc.common import *

def document_creater(SystemSize,ClientName, ClientLocation, ReferredBy, Inverter_TYP, Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice, Number_of_Panels, Structure_Type, pv_balance, Carriage_Cost, Installation, Net_Metering, template_file, InverterPrice, TotalCostNormal, TotalCostRaised):
    # Create a Document object
    document = Document()

    # Load the template
    document.LoadFromFile(template_file)
    NOS = int(Number_of_Panels)/2

    # Store the placeholders and new strings in a dictionary
    dictionary = {
                    '[sw]': SystemSize,
                    '[nop]': Number_of_Panels,
                    '[nos]': str(NOS),
                    '[noi]': '1',
                    '[iw]': Inverter_Watt,
                    '[iwy]': '5',
                    '[calculated_price]':TotalCostRaised,
                    '[calculated_price_normal]':TotalCostNormal,
                    '[value_in_words]':'Five Hundred Thousand',
                    '[bp]':'180,000',
                    '[bv]':'4',
                    '[nmc]':'100,000'
                }

    # Loop through the items in the dictionary
    for key, value in dictionary.items():
        # Replace a placeholder (key) with a new string (value)
        document.Replace(key, value, False, True)


    # Save the resulting document
    document.SaveToFile("output/ReplacePlaceholder.docx", FileFormat.Docx2016)
    document.Close()