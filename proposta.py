from docx import Document

# Caminho para o seu documento Word
docx_path = 'Template_Proposta_Comercial.docx'

# Abre o documento
doc = Document(docx_path)

# Itera sobre todos os estilos e imprime aqueles que são de tipo 'Numbering'
print("Estilos de marcadores disponíveis no documento:")
for style in doc.styles:
    if style.type == 4:  # 4 indica estilos de numeração e marcadores
        print(style.name)
