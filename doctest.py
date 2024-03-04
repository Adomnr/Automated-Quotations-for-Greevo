from spire.doc import *
from spire.doc.common import *

# Create a Document object
document = Document()

# Load the template
document.LoadFromFile("template.docx")

# Store the placeholders and new strings in a dictionary
dictionary = {
                '[sw]':'10kw',
                '[nop]': '20',
                '[nos]': '10',
                '[noi]': '1',
                '[iw]': '10',
                '[iwy]': '5',
                '[calculated_price]':'500,000',
                '[calculated_price_normal]':'450,000',
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