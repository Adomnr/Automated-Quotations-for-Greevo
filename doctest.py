from spire.doc import *
from spire.doc.common import *

def document_creater(SystemSize,Inverter_Watt, Number_of_Panels, Net_Metering, template_file, TotalCostNormal, TotalCostRaised, valueinwords):
    # Create a Document object
    document = Document()

    # Load the template
    document.LoadFromFile(template_file)
    NOS = int(Number_of_Panels)/2

    # Store the placeholders and new strings in a dictionary
    dictionary = {
                    '[sw]': str(SystemSize),
                    '[nop]': str(Number_of_Panels),
                    '[nos]': str(NOS),
                    '[noi]': '1',
                    '[iw]': str(Inverter_Watt),
                    '[iwy]': '5',
                    '[calculated_price]':str(TotalCostRaised),
                    '[calculated_price_normal]':str(TotalCostNormal),
                    '[value_in_words]':str(valueinwords),
                    '[bp]':'180,000',
                    '[bv]':'4',
                    '[nmc]': str(Net_Metering)
                }

    # Loop through the items in the dictionary
    for key, value in dictionary.items():
        # Replace a placeholder (key) with a new string (value)
        document.Replace(key, value, False, True)


    # Save the resulting document
    document.SaveToFile("output/ReplacePlaceholder.docx", FileFormat.Docx2016)
    document.Close()