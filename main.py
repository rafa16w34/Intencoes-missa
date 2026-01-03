from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Border,Side
from openpyxl.utils import get_column_letter
from docx import Document

data_hoje  = datetime.now()

data_ajustada = data_hoje + timedelta(hours=1)#Soma 1 na hora porque temos que imprimir 10 minutos antes
data_formatada = data_ajustada.strftime("%d/%m/%Y - %H:00")

#Listas para cada uma das intenções
categorias = {
    "HONRA E LOUVOR": [],
    "INTENÇÃO": [],
    "ANIVERSÁRIO NATALÍCIO": [],
    "RECUPERAÇÃO DE SAÚDE": [],
    "ANIVERSÁRIO DE CASAMENTO": [],
    "AÇÃO DE GRAÇAS": [],
    "IRMÃOS FALECIDOS": [],
    "7º DIA": []
}

pix = []

#Variavel do dinheiro
valor_total = 0

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#FUNÇÃO DE ESCANEAR O DOCUMENTO WORD

def extrair_docx():

    def normalizar(txt):
        return " ".join(txt.strip().upper().split())

    arquivo = "intencoes.docx"
    doc = Document(arquivo)

    TITULOS = [
        "HONRA E LOUVOR",
        "INTENÇÃO",
        "ANIVERSÁRIO NATALÍCIO",
        "ANIVERSÁRIO DE CASAMENTO",
        "RECUPERAÇÃO DE SAÚDE",
        "AÇÃO DE GRAÇAS",
        "IRMÃOS FALECIDOS",
    ]

    TITULOS_NORM = [normalizar(t) for t in TITULOS]

    dados = {}
    titulo_atual = None

    # LER TABELAS EM VEZ DE PARÁGRAFOS
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:

                linha = cell.text.strip()
                if not linha:
                    continue

                ln = normalizar(linha)

                if ln in TITULOS_NORM:  # detecta título
                    titulo_atual = ln
                    dados[titulo_atual] = []
                    continue

                if titulo_atual:
                    
                    if (linha != "+") and not (linha.startswith("INTENÇÕES")):

                        dados[titulo_atual].append(linha)


    return dados



dados_importados = extrair_docx()



for titulo, itens in dados_importados.items():

    for item in itens:

        # se NÃO começar com 7º, mantém na categoria original
        if not item.startswith("7º"):
            categorias[titulo].append(item)

        # se começar com 7º, manda para a categoria 7º DIA
        else:
            categorias["7º DIA"].append(item)


#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#FUNÇÃO DE MARCAR MISSA

def menu_pessoa(valor_pessoa):

    sair_menu2 = 0

    while sair_menu2 == 0:#Menu para marcar a intenção de uma pessoa

        try:

            opcao = int(input('\nDigite uma das opções abaixo:\n1- Honra e Louvor\n2- Aniversário\n3- Aniversário de Casamento\n4- Ação de Graças\n5- Intenções\n6- Recuperação de Saúde\n7- 7º Dia\n8- Almas\n------------------------------------------------------------\n9- Voltar\n-> '))

            match(opcao):

                case 1:#Adiciona uma intenção de Honra e Louvor

                    
                    nome = str(input('\nDigite a intenção em Honra e Louvor:\n-> '))

                    nome = nome.lower()

                    categorias["HONRA E LOUVOR"].append(nome.title())
                    valor_pessoa += 3


                case 2:#Adiciona uma intenção para Aniversário


                    nome = str(input('\nDigite o nome do aniversariante:\n-> '))

                    nome = nome.lower()

                    categorias["ANIVERSÁRIO NATALÍCIO"].append(nome.title())
                    valor_pessoa += 3


                case 3:#Adiciona uma intenção de Aniversário de Casamento
                    
                            
                    nome = str(input('\nDigite o nome do casal aniversariante:\n-> '))

                    nome = nome.lower()

                    categorias["ANIVERSÁRIO DE CASAMENTO"].append(nome.title())
                    valor_pessoa += 3


                case 4:#Adiciona uma intenção de Ação de Graças
                    
                            
                    nome = str(input('\nDigite a intenção de acão de graças:\n-> '))

                    nome = nome.lower()

                    categorias["AÇÃO DE GRAÇAS"].append(nome.title())
                    valor_pessoa += 3


                case 5:#Adiciona uma intenção de Intenções
                    
                            
                    nome = str(input('\nDigite a intenção:\n-> '))

                    nome = nome.lower()

                    categorias["INTENÇÃO"].append(nome.title())
                    valor_pessoa += 3



                case 6:#Adiciona uma intenção de Recuperação de Saúde
                    
                            
                    nome = str(input('\nDigite o nome do enfermo:\n-> '))

                    nome = nome.lower()

                    categorias["RECUPERAÇÃO DE SAÚDE"].append(nome.title())
                    valor_pessoa += 3


                case 7:#Adiciona uma intenção de Sétimo dia
                    
                            
                    nome = str(input('\nDigite o nome do sétimo dia:\n-> '))

                    nome = nome.lower()

                    categorias['7º DIA'].append(nome.title())
                    valor_pessoa += 3

                case 8:#Adiciona uma intenção de Alma
                    
                            
                    nome = str(input('\nDigite o nome da alma:\n-> '))

                    nome = nome.lower()

                    categorias['IRMÃOS FALECIDOS'].append(nome.title())
                    valor_pessoa += 3

                case 9:

                
                    print(f'\nO valor a ser pago é de R$ {valor_pessoa:.2f}\n')

                    metodo_de_pagamento = int(input('\nQual vai ser a forma de pagamento?\n1- Dinheiro\n2- Pix\n-> '))

                    match(metodo_de_pagamento):

                        case 1:#Dinheiro

                            while True:
                                
                                try:

                                    valor_recebido = int(input('\nDigite o valor recebido:\n-> R$ '))

                                    if (valor_recebido < valor_pessoa):

                                        print('\nDigite um valor válido!\n')


                                    else:
                                        
                                        if (valor_recebido > valor_pessoa):

                                            print(f'\nTroco: R$ {(valor_recebido - valor_pessoa):.2f}\n')
                                        
                                        break

                                
                                
                                except(ValueError):
                                    
                                    print('\nDigite um valor válido!\n')

                        case 2:#Pix

                            nome_pix = str(input("\nQual é o nome quem irá fazer o pix?\n-> "))

                            nome_pix = nome_pix.lower()

                            pix.append(nome_pix.title())
                            pix.append(valor_pessoa)

                            with open ('Lista_Pix.txt','w') as arq:

                                arq.write(f'Lista do Pix {data_formatada}:\n')

                                for i in range(len(pix)):

                                    if ((i+1)%2 != 0):                                    
                                        arq.write(f'- {pix[i]} : R$ {pix[i+1]},00\n')
                                        i+2
                                    
                                    if i+2 >= len(pix):
                                        break



                    confirmacao = str(input('\nAperte enter para continuar\n'))

                    print('\nVoltando ao menu inicial...\n')

                    sair_menu2 = 1
                

                case _: 

                    print('\nErro: Digite uma das opções mostradas no menu!\n')
        
        except(ValueError):
            print('\nErro: Digite uma das opções mostradas no menu!\n')
    
    return(valor_pessoa)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#FUNÇÃO DE IMPRIMIR NO EXCEL

def imprimir():

    # Criar planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Intenções"

#-------------------------------------------------------------------------------------------------------------
    ws.append([f"INTENÇÕES DA MISSA {data_formatada}"])
    ws.append([])  # linha em branco

#-------------------------------------------------------------------------------------------------------------
    
    for titulo, lista in categorias.items():
        ws.append([titulo + ":"])
        for item in lista:
            ws.append([item])
        ws.append([])

#-------------------------------------------------------------------------------------------------------------

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # Aplicar bordas em todas as células usadas
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = borda

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)

        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass

        ws.column_dimensions[col_letter].width = max_length + 5

#-------------------------------------------------------------------------------------------------------------
    # Salvar arquivo
    wb.save("Intenções.xlsx")

    print("\nArquivo Excel 'Intenções.xlsx' criado com sucesso!\n")


#-------------------------------------------------------------------------------------------------------------
    print(f'\nO valor total do dia é de: R$ {valor_total:.2f}\n')
    
    sair_menu2 = 1

#-------------------------------------------------------------------------------------------------------------

#FUNÇÃO DE EDITAR A INTENÇÃO JÁ MARCADA

def editar(lista):

    divida = 0

    if (len(lista) != 0):
        
        for i in range(len(lista)+1):

            if (i < len(lista)):
                print(f'\n{i} : {lista[i]}') #Lista todas as intenções marcadas no tipo escolhido com o indice na frente
            else:
                print(f'\n{i} : Sair')
            
        print('\nEscolha qual será editado pelo índice:\n')

        while True:

            try:
                edicao = int(input('-> '))

            except ValueError:
                print('\nDigite um valor válido!\n')
                continue

            # Verifica se o índice existe
            if (0 <= edicao) and (edicao < len(lista)):
                sair = False
                break

            elif (edicao >= len(lista)):
                sair = True
                break

            else:

                print('\nÍndice inválido! Tente novamente.\n')

        while sair == False:

            print('\nGostaria de remover, editar ou sair?\n')
            opcao_edicao = input('-> ').lower()

            if opcao_edicao == 'remover':
                lista.pop(edicao)
                divida -= 3
                print("\nIntenção removida com sucesso!\n")
                break

            elif opcao_edicao == 'editar':
                nova_intencao = input('\nDigite a nova intenção:\n-> ')
                lista[edicao] = nova_intencao
                print('\nIntenção editada com sucesso!\n')
                break

            elif opcao_edicao == 'sair':
                break

            else:
                print('Digite "remover" ou "editar" !')

    else:
        print('\nErro: Lista vazia!\n')

    return(divida)


#-------------------------------------------------------------------------------------------------------------

#MENU PRINCIPAL

sair_menu1 = 0

while (sair_menu1 == 0):
            
            try:
                opcao = int(input('\nDigite uma das opções a seguir:\n1- Marcar intenção\n2- Dados\n3- Editar\n4- Imprimir\n0- Finalizar programa\n-> '))

                match(opcao):

                    case 1:
                        valor_pessoa = 0
                        

                        valor_total += menu_pessoa(valor_pessoa)

                    case 2:

                        print(f'\nDados atuais:\nTotal de intenções: {valor_total/3}\nValor total: R$ {valor_total:.2f}')

                    case 3:

                        try:
                            opcao_editar = int(input('\nGostaria de editar qual dos tipos de intenção?\n1- Honra e Louvor\n2- Aniversário\n3- Aniversário de Casamento\n4- Ação de Graças\n5- Intenções\n6- Recuperação de Saúde\n7- 7º Dia\n8- Almas\n9- Voltar\n-> '))
                            
                            lista = None

                            match(opcao_editar):

                                case 1:
                                    lista = categorias['HONRA E LOUVOR']

                                case 2:
                                    lista = categorias['ANIVERSÁRIO NATALÍCIO']

                                case 3:
                                    lista = categorias['ANIVERSÁRIO DE CASAMENTO']

                                case 4:
                                    lista = categorias['AÇÃO DE GRAÇAS']

                                case 5:
                                    lista = categorias['INTENÇÃO']

                                case 6:
                                    lista = categorias['RECUPERAÇÃO DE SAÚDE']

                                case 7:
                                    lista = categorias['7º DIA']

                                case 8:
                                    lista = categorias['IRMÃOS FALECIDOS']

                                case 9:
                                    print('\nVoltando ao menu..\n')
                                    continue

                                case _:
                                    print('\nDigite uma das opções exibidas no menu!\n')
                                    continue

                            if lista is not None:
                                valor_total += editar(lista)


                        except(ValueError):
                            print('\nDigite um valor válido!\n')    


                    case 4:#Imprime a lista formatada
                        

                        imprimir()


                    case 0:

                        confirmacao = str(input('\nDigite "sim" para sair.\n(Tudo o que foi digitado será perdido caso não tenha sido impresso!)\n-> '))

                        if (confirmacao == 'sim'):
                            print('\nFinalizando programa...\n')
                            sair_menu1 = 1



                    case _: 

                        print('\nErro: Digite uma das opções mostradas no menu!\n')


            except(ValueError):
                print('\nErro: Digite uma das opções mostradas no menu!\n')