import openpyxl
from openpyxl.utils import get_column_letter
import time
import docx


try:
    Tabela_Orcamento = openpyxl.load_workbook('Orçamento.xlsx')['ORÇAMENTO']
except:
    print('A arquivo do orçamento deve ter o nome "Orçamento.xlsx" e a tabela dentro do arquivo deve ter o nome "ORÇAMENTO"')
    time.sleep(15)
Coluna_A = Tabela_Orcamento['A1':'A'+str(Tabela_Orcamento.max_row)]

Itens = [cell[0].value for cell in Coluna_A]
try:
    Linha_Comeca = Itens.index('Item')+1
except:
    print('Os itens devem estar na coluna A da tabela de Orçamento')
    time.sleep(15)
try:
    Ultima_Linha = (Itens.index(None,Linha_Comeca))
except:
    Ultima_Linha = len(Itens)
Itens = Itens[Linha_Comeca-1:Ultima_Linha]



Linha_Inicial_Tabela = Tabela_Orcamento['A'+str(Linha_Comeca):get_column_letter(Tabela_Orcamento.max_column)+str(Linha_Comeca)]
Linha_Inicial_Tabela = [cell.value for cell in Linha_Inicial_Tabela[0]]
Coluna_Codigo = Linha_Inicial_Tabela.index('Código')+1

Codigos = Tabela_Orcamento[get_column_letter(Coluna_Codigo)+str(Linha_Comeca):get_column_letter(Coluna_Codigo)+str(Ultima_Linha)]
Codigos = [str(cell[0].value).strip() for cell in Codigos]


Coluna_Titulos = Linha_Inicial_Tabela.index('Descrição')+1

Titulos = Tabela_Orcamento[get_column_letter(Coluna_Titulos)+str(Linha_Comeca):get_column_letter(Coluna_Titulos)+str(Ultima_Linha)]

Titulos = [cell[0].value for cell in Titulos]
#print(Titulos[214])
#print(Codigos[214])
#print(Itens[214])
try:
    Tabela_Banco_Dados_1 = openpyxl.load_workbook('Base de Dados.xlsx')['Plan1']
    Tabela_Banco_Dados_2 = openpyxl.load_workbook('Base de Dados.xlsx')['Plan2']
except:
    print('A o arquivo da base de dados deve ter o nome "Base de Dados.xlsx" e a primeira tabela deve ter o nome "Plan1" e a segunda "Plan2')
    time.sleep(15)
codigos_aux = [cell.value for cell in Tabela_Banco_Dados_1['A']]
titulos_aux = [cell.value for cell in Tabela_Banco_Dados_1['B']]
Banco_Dados_Codigos_Titulos = dict(zip(codigos_aux,titulos_aux))


codigos_aux2 = [str(cell.value).strip() for cell in Tabela_Banco_Dados_2['A']]
descricoes_aux = [cell.value for cell in Tabela_Banco_Dados_2['B']]
Banco_Dados_Codigos_Descricoes = dict(zip(codigos_aux2,descricoes_aux))

doc = docx.Document()
lista_sem_codigos = docx.Document()
font = doc.styles['Normal'].font
font.name = 'Arial'
sem_codigo =[]
for item,titulo,codigo in zip(Itens[1:],Titulos[1:],Codigos[1:]):

    if item != '':
        if codigo =='None' and len(item.strip())<=2:
        
            p = doc.add_paragraph()
            p.paragraph_format.space_after = docx.shared.Pt(0)
            r=p.add_run(item.strip()+'. ' +titulo+'\n')
        
            r.font.highlight_color = docx.enum.text.WD_COLOR.GRAY_25
            
            r.bold =True
        
        else:
            p2 = doc.add_paragraph()
            p2.paragraph_format.space_after = docx.shared.Pt(0)
            p2.add_run(item.strip() +' '+ titulo).bold = True
        
            
            p3 = doc.add_paragraph()   
            try:
                p3.add_run(Banco_Dados_Codigos_Descricoes[codigo])
                p3.paragraph_format.space_after = docx.shared.Pt(0)
            except:
                p4 = lista_sem_codigos.add_paragraph(item).add_run('---'+codigo)

doc.save('Memorial Descritivo.docx')
lista_sem_codigos.save('Itens Sem Código.docx')
print('Deu Tudo certo :) ')
time.sleep(10)