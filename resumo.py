from docx import Document

titulo = "arroz"
corporesumo = "Escreva um resumo da sua invenção aqui em um único parágrafo de no máximo 25 linhas. Indique o setor técnico da sua invenção e faça uma breve descrição dela dando informações essenciais sobre o que a caracteriza e o que a diferencia do estado da técnica. Esta seção do pedido de patente é muito utilizada nas buscas feitas pelos examinadores e também por outros interessados."
# input()

document = Document('C:/teste/resumo.docx')
font = document.styles['Normal'].font
for paragraph in document.paragraphs:
    if 'ESCREVA AQUI O TÍTULO DO SEU PEDIDO DE PATENTE (deve ser idêntico ao informado no formulário de depósito)'in paragraph.text:
        estilo = paragraph.style
        paragraph.text = titulo
        font.name = 'Arial'
        font.bold = True
        break
    if 'texto de exemplo' in paragraph.text:
       paragraph.text = corporesumo 


document.save('C:\\teste\\resumo_editado.docx')