from trademarkremover import *

def convert_docx_to_pdf(input_docx, output_pdf):
    # Read the .docx file
    doc = Document(input_docx)

    # Create a PDF file
    pdf_canvas = canvas.Canvas(output_pdf, pagesize=letter)

    for para in doc.paragraphs:
        pdf_canvas.drawString(50, 800, para.text)  # Adjust the position as needed

    pdf_canvas.save()