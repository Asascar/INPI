from docx import Document
import sys
import os
import comtypes.client

document = Document('relatorio descritivo.docx')

titulo = input('ESCREVA AQUI O TÍTULO DO SEU PEDIDO DE PATENTE (deve ser idêntico ao informado no formulário de depósito): \n')
campoinvecao = input('Descreva aqui o setor técnico ao qual se refere sua invenção. O setor técnico pode ser composições de tintura capilar, máquinas para semeadura ou comunicações de rede sem fio, por exemplo. Se sua invenção puder ser aplicada em mais de um campo técnico cite todos eles: \n')
estadotec = input('Escreva aqui o estado da técnica relacionado à sua invenção, ou seja, aquilo que já se conhece sobre inventos parecidos com o seu. Procure apresentar as características mais importantes desses inventos. É isso o que pede o artigo 2°, inciso IV, da Instrução Normativa n° 30/2013. Use quantos parágrafos forem necessários:\n' )
problematec = input('Em seguida, você deve apresentar o problema técnico que ainda não foi solucionado pelo estado da técnica e mostrar como sua invenção resolve esse problema. Ou seja, você deve mostrar as diferenças da sua invenção em relação às invenções do estado da técnica e apresentar as vantagens da sua. É muito importante destacar o benefício ou efeito técnico da sua invenção (mais eficiente, mais barata, ocupa menos espaço, não contém elementos tóxicos para o meio ambiente etc), pois o examinador de patentes levará isso em consideração durante o exame do seu pedido de patente:\n')
desenhos = input('Se o seu pedido de patente tiver desenhos (podem ser figuras, gráficos ou desenhos propriamente ditos) descreva de forma breve as informações apresentadas em cada um dos desenhos. Uma a duas linhas são suficientes para essa descrição. As linhas que contêm as descrições dos desenhos não precisam conter numeração sequencial dos parágrafos:\n')
descinvencao = input('Essa é a maior seção do relatório descritivo, que pode ter de poucas até centenas de páginas. Apresente de forma detalhada sua invenção nessa seção e inclua todas as suas possibilidades de concretização. Você pode iniciar por uma ideia geral da invenção para detalhá-la melhor nos parágrafos seguintes. Mais importante do que escrever muitas páginas sobre sua invenção é descrevê-la de forma clara e precisa, de forma que o examinador de patentes possa entender o que você inventou e como sua invenção funciona:\n')
concretizacao = input('Nesta seção do relatório descritivo você deve apresentar exemplos de concretizações da sua invenção, seja ela um composto, uma composição, um equipamento, um processo etc. Se for o caso, você deve também indicar qual é a forma preferida de concretizar sua invenção. Por exemplo, se sua invenção for uma composição, você deve indicar qual composição (ou tipo de composição) é preferida dentre as várias possíveis composições que sua invenção representa:\n')
estadotecnica = input('Entre com as informações do estado da técnica: \n')

print("Você gostaria de adicionar um banco de palavras: (S/N)")
resp = input()
relacao = 1
if resp == 'S':
 dicionario = {}
 while resp == 'S':
    dicionario[relacao] = input(f'Qual palavra gostaria de acrescentar como {relacao} ? \n')
    relacao = relacao + 1
    resp = input('Gostaria de acrescentar outra entrada ao dicionario ? (S/N)')   

font = document.styles['Normal'].font
for paragraph in document.paragraphs:
    if 'ESCREVA AQUI O TÍTULO DO SEU PEDIDO DE PATENTE (deve ser idêntico ao informado no formulário de depósito)'in paragraph.text:
        estilo = paragraph.style
        paragraph.text = titulo
        font.name = 'Arial'
        font.bold = True
    if 'Descreva aqui o setor técnico ao qual se refere sua invenção. O setor técnico pode ser composições de tintura capilar, máquinas para semeadura ou comunicações de rede sem fio, por exemplo. Se sua invenção puder ser aplicada em mais de um campo técnico cite todos eles.'in paragraph.text:
        estilo = paragraph.style
        paragraph.text = campoinvecao
        font.name = 'Arial'
        font.bold = False
    if 'Escreva aqui o estado da técnica relacionado à sua invenção, ou seja, aquilo que já se conhece sobre inventos parecidos com o seu.' in paragraph.text:
       paragraph.text = estadotec
       font.name = 'Arial'
       font.bold = False
    if 'Em seguida, você deve apresentar o problema técnico que ainda não foi solucionado pelo estado da técnica e mostrar como sua invenção resolve esse problema. Ou seja, você deve mostrar as diferenças da sua invenção em relação às invenções do estado da técnica e apresentar as vantagens da sua. É muito importante destacar o benefício ou efeito técnico da sua invenção (mais eficiente, mais barata, ocupa menos espaço, não contém elementos tóxicos para o meio ambiente etc), pois o examinador de patentes levará isso em consideração durante o exame do seu pedido de patente.' in paragraph.text:
       paragraph.text = problematec
       font.name = 'Arial'
       font.bold = False
    if 'Se o seu pedido de patente tiver desenhos (podem ser figuras, gráficos ou desenhos propriamente ditos) descreva de forma breve as informações apresentadas em cada um dos desenhos. Uma a duas linhas são suficientes para essa descrição. As linhas que contêm as descrições dos desenhos não precisam conter numeração sequencial dos parágrafos. Por exemplo:' in paragraph.text:
       paragraph.text = desenhos
       font.name = 'Arial'
       font.bold = False
    if 'Essa é a maior seção do relatório descritivo, que pode ter de poucas até centenas de páginas. Apresente de forma detalhada sua invenção nessa seção e inclua todas as suas possibilidades de concretização. Você pode iniciar por uma ideia geral da invenção para detalhá-la melhor nos parágrafos seguintes. Mais importante do que escrever muitas páginas sobre sua invenção é descrevê-la de forma clara e precisa, de forma que o examinador de patentes possa entender o que você inventou e como sua invenção funciona' in paragraph.text:
       paragraph.text = descinvencao
       font.name = 'Arial'
       font.bold = False
    if 'Nesta seção do relatório descritivo você deve apresentar exemplos de concretizações da sua invenção, seja ela um composto, uma composição, um equipamento, um processo etc. Se for o caso, você deve também indicar qual é a forma preferida de concretizar sua invenção. Por exemplo, se sua invenção for uma composição, você deve indicar qual composição (ou tipo de composição) é preferida dentre as várias possíveis composições que sua invenção representa.' in paragraph.text:
       paragraph.text = concretizacao
       font.name = 'Arial'
       font.bold = False
    if 'Estado da técnica' in paragraph.text:
       paragraph.text = estadotecnica
       font.name = 'Arial'
       font.bold = False
    


document.save('relatorio descritivo_editado.docx')

wdFormatPDF = 17

in_file = ('relatorio descritivo.docx')
out_file = ('relatorio descritivo_editado.pdf')

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()