#Bibliotecas para manipulação do Word
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.shared import Inches
#Biblioteca para escrever números por extenso
from num2words import num2words

#Declarando variáveis
ctrl = 1

#Definindo a função
def geradorRecibo(inquilino, endereco, tipo, valor, mes, vencimento):
    doc = docx.Document('recibo_branco.docx')
    
    num_ext = num2words(v, lang="pt-br")

    #Criando o texto "RECIBO"
    h1 = doc.add_paragraph("RECIBO DE ALUGUEL")
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in h1.runs:
        r.font.size = Pt(24)
        r.bold = True
        r.font.name = 'Arial'
        
    space = doc.add_paragraph()
    space.add_run().add_break()
    space.add_run().add_break()


    #Criando o corpo do texto
    p = doc.add_paragraph()
    p.add_run('Recebi do(a) Sr.(a) ')
    p.add_run(f'{inquilino},').bold = True

    p.add_run().add_break()

    p.add_run('A quantia de ')
    p.add_run(f'R${v}').bold = True
    p.add_run(f' ({num_ext}),').bold = True

    p.add_run().add_break()

    p.add_run('Referente ao aluguel do mês de ')
    p.add_run(f'{mes},').bold = True

    p.add_run().add_break()

    p.add_run('Do imóvel de endereço ')
    p.add_run(f'{endereco}, ').bold = True
    p.add_run('de cunho ')
    p.add_run(f'{tipo},').bold = True

    p.add_run().add_break()

    p.add_run('Com vencimento em ')
    p.add_run(f'{vencimento}',).bold = True

    #Espaçamento
    p.add_run().add_break()
    p.add_run().add_break()
    p.add_run().add_break()

    #Campo de assinatura 
    sig = doc.add_paragraph()
    sig.add_run('__________________________________')
    sig.add_run().add_break()
    sig.add_run('Assinatura').bold = True
    sig.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #Estilizando
    for r in p.runs:
        r.font.size = Pt(16)
        r.font.name = 'Arial'

    for r in sig.runs:
        r.font.size = Pt(16)
        r.font.name = 'Arial'

    #Salvando o documento
    doc.save(f'recibo_{inquilino}_{mes}.docx')

while ctrl == 1:
    inq = input("Inquilino: ")
    address = input("Endereco: ")
    tipo = int(input("Tipo [1- Residencial/ 2- Comercial]"))
    v = int(input("Valor: "))
    mes = input("Mês: ")
    vencimento = input("Vencimento: ")
    
    
    #Chamando a função
    geradorRecibo(inq, address, tipo, v, mes, vencimento)
    
    p = input("Deseja gerar outro recibo? [s/n]").lower()
    
    if p == "n":
        ctrl = 0