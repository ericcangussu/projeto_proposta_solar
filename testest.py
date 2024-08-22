import pythoncom
from win32com.client import Dispatch

def replace_text_in_textboxes(docx_path, old_text, new_text, output_path):
    # Inicializa o Word para COM
    pythoncom.CoInitialize()
    word = Dispatch("Word.Application")
    word.Visible = False

    # Abre o documento Word
    doc = word.Documents.Open(docx_path)

    # Percorre todas as formas no documento para encontrar caixas de texto
    for shape in doc.Shapes:
        if shape.TextFrame.HasText:
            text_range = shape.TextFrame.TextRange
            if old_text in text_range.Text:
                text_range.Text = text_range.Text.replace(old_text, new_text)

    # Salva o documento modificado
    doc.SaveAs(output_path)
    doc.Close(False)
    word.Quit()

# Caminho para o seu documento Word
docx_path = 'Template_Proposta_Comercial.docx'

# Substitui o texto '_varCliente' por 'Nome do Cliente' e salva como 'documento_modificado.docx'
replace_text_in_textboxes(docx_path, '_varCliente', 'Nome do Cliente','documento_modificado.docx')

print("Documento salvo como 'documento_modificado.docx'")