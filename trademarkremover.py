from docx import Document
import os

def main():
    template_file_path = './output/ReplacePlaceholder.docx'
    output_file_path = './wr/WR.docx'

    variables = {
		'Evaluation Warning: The document was created with Spire.Doc for Python.':' '
	}

    template_document = Document(template_file_path)

    for variable_key, variable_value in variables.items():
        for paragraph in template_document.paragraphs:
            replace_text_in_paragraph(paragraph, variable_key, variable_value)

    template_document.save(output_file_path)

def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)


if __name__ == '__main__':
    main()
