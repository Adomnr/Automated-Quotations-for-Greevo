from spire.doc import *

def document_creater(UniqueID, ClientName, SystemSize, Inverter_Watt, Number_of_Panels, Net_Metering, template_file, TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords):
    # Create a Document object
    document = Document()

    # Load the template
    document.LoadFromFile(template_file)
    NOS = int(Number_of_Panels)/2

    # Store the placeholders and new strings in a dictionary
    dictionary = {
                    '[uid]': str(UniqueID),
                    '[sw]': str(SystemSize),
                    '[nop]': str(Number_of_Panels),
                    '[nos]': str(NOS),
                    '[noi]': '1',
                    '[iw]': str(Inverter_Watt),
                    '[iwy]': '5',
                    '[calculated_price]':str(TotalCostRaised),
                    '[calculated_price_normal]':str(TotalCostNormal),
                    '[calculated_price_NNI]': str(TotalCostRaisedNNI),
                    '[calculated_price_normal_NNI]': str(TotalCostNormalNNI),
                    '[value_in_words]':str(valueinwords),
                    '[bp]':'180,000',
                    '[bv]':'4',
                    '[nmc]': str(Net_Metering)
                }

    # Loop through the items in the dictionary
    for key, value in dictionary.items():
        # Replace a placeholder (key) with a new string (value)
        document.Replace(key, value, False, True)

    filename = str(UniqueID)+str(ClientName)+ "_Quotation"
    print(filename)
    savename = "output/"+filename+".docx"


    # Save the resulting document
    document.SaveToFile(savename, FileFormat.Docx2016)
    document.Close()