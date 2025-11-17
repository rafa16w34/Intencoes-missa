from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Border,Side

data_hoje  = datetime.now()

data_formatada = data_hoje.strftime("%d/%m/%Y - %H:00")

#Listas para cada uma das intenções
honra_e_louvor = []
aniversario = []
aniversario_de_casamento = []
acao_de_gracas = []
intencoes = []
recuperacao_de_saude = []
setimo_dia = []
almas = []

pix = []

#Variavel do dinheiro
valor_total = 0



def menu_pessoa(valor_pessoa):

    sair_menu2 = 0

    while sair_menu2 == 0:#Menu para marcar a intenção de uma pessoa

        try:

            opcao = int(input('\nDigite uma das opções abaixo:\n1- Honra e Louvor\n2- Aniversário\n3- Aniversário de Casamento\n4- Ação de Graças\n5- Intenções\n6- Recuperação de Saúde\n7- 7º Dia\n8- Almas\n------------------------------------------------------------\n9- Voltar\n-> '))

            match(opcao):

                case 1:#Adiciona uma intenção de Honra e Louvor

                    
                    nome = str(input('\nDigite a intenção em Honra e Louvor:\n-> '))

                    nome = nome.lower()

                    honra_e_louvor.append(nome.title())
                    valor_pessoa += 3


                case 2:#Adiciona uma intenção para Aniversário


                    nome = str(input('\nDigite o nome do aniversariante:\n-> '))

                    nome = nome.lower()

                    aniversario.append(nome.title())
                    valor_pessoa += 3


                case 3:#Adiciona uma intenção de Aniversário de Casamento
                    
                            
                    nome = str(input('\nDigite o nome do casal aniversariante:\n-> '))

                    nome = nome.lower()

                    aniversario_de_casamento.append(nome.title())
                    valor_pessoa += 3


                case 4:#Adiciona uma intenção de Ação de Graças
                    
                            
                    nome = str(input('\nDigite a intenção de acão de graças:\n-> '))

                    nome = nome.lower()

                    acao_de_gracas.append(nome.title())
                    valor_pessoa += 3


                case 5:#Adiciona uma intenção de Intenções
                    
                            
                    nome = str(input('\nDigite a intenção:\n-> '))

                    nome = nome.lower()

                    intencoes.append(nome.title())
                    valor_pessoa += 3



                case 6:#Adiciona uma intenção de Recuperação de Saúde
                    
                            
                    nome = str(input('\nDigite o nome do enfermo:\n-> '))

                    nome = nome.lower()

                    recuperacao_de_saude.append(nome.title())
                    valor_pessoa += 3


                case 7:#Adiciona uma intenção de Sétimo dia
                    
                            
                    nome = str(input('\nDigite o nome do sétimo dia:\n-> '))

                    nome = nome.lower()

                    setimo_dia.append(nome.title())
                    valor_pessoa += 3

                case 8:#Adiciona uma intenção de Alma
                    
                            
                    nome = str(input('\nDigite o nome da alma:\n-> '))

                    nome = nome.lower()

                    almas.append(nome.title())
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



                    confirmacao = str(input('\nAperte qualquer tecla para continuar\n'))

                    print('\nVoltando ao menu inicial...\n')

                    sair_menu2 = 1
                

                case _: 

                    print('\nErro: Digite uma das opções mostradas no menu!\n')
        
        except(ValueError):
            print('\nErro: Digite uma das opções mostradas no menu!\n')
    
    return(valor_pessoa)


def imprimir():

    # Criar planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Intenções"

#############################################################################################################
    ws.append([f"INTENÇÕES DA MISSA {data_formatada}"])
    ws.append([])  # linha em branco

#############################################################################################################
    # Inserir Honra e Louvor
    ws.append(['HONRA E LOUVOR:'])

    for i in range(len(honra_e_louvor)):

        ws.append([honra_e_louvor[i]])
    
    ws.append([])  # linha em branco

#############################################################################################################
    # Inserir Aniversário
    ws.append(['ANIVERSÁRIO:'])

    for i in range(len(aniversario)):

        ws.append([aniversario[i]])

    ws.append([])  # linha em branco

#############################################################################################################
    # Inserir Aniversário de Casamento
    ws.append(['ANIVERSÁRIO DE CASAMENTO:'])

    for i in range(len(aniversario_de_casamento)):

        ws.append([aniversario_de_casamento[i]])

    ws.append([])  # linha em branco

#############################################################################################################
    # Inserir Ação de Graças
    ws.append(['AÇÃO DE GRAÇAS:'])

    for i in range(len(acao_de_gracas)):
        ws.append([acao_de_gracas[i]])
    ws.append([])  # linha em branco

#############################################################################################################
    # Inserir Intenções
    ws.append(['INTENÇÕES:'])

    for i in range(len(intencoes)):

        ws.append([intencoes[i] ])

    ws.append([])  # linha em branco

#############################################################################################################
    # Inserir Recuperação de Saúde
    ws.append(['RECUPERAÇÃO DE SAÚDE:'])

    for i in range(len(recuperacao_de_saude)):

        ws.append([recuperacao_de_saude[i]])

    ws.append([])  # linha em branco

#############################################################################################################
    # Inserir Sétimo Dia
    ws.append(["7º DIA:"])

    for i in range(len(setimo_dia)):

        ws.append([f'+ {setimo_dia[i]}'])

    ws.append([])  # linha em branco

#############################################################################################################
    # Inserir Almas
    ws.append(["IRMÃOS FALECIDOS:"])

    for i in range(len(almas)):

        ws.append([f'+ {almas[i]}'])

    ws.append([])  # linha em branco

#############################################################################################################

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

#############################################################################################################
    # Salvar arquivo
    wb.save("Intenções.xlsx")

    print("\nArquivo Excel 'Intenções.xlsx' criado com sucesso!\n")


#############################################################################################################
    print(f'\nO valor total do dia é de: R$ {valor_total:.2f}\n')
    
    sair_menu2 = 1

#############################################################################################################

def editar(lista):

    divida = 0

    if (len(lista) != 0):
        
        for i in range(len(lista)):
            print(f'\n{i} : {lista[i]}')
        print('\nEscolha qual será editado pelo índice:\n')

        while True:

            try:
                edicao = int(input('-> '))
            except ValueError:
                print('\nDigite um valor válido!\n')
                continue

            # Verifica se o índice existe
            if (0 <= edicao) and (edicao < len(lista)):
                break
            else:
                print('\nÍndice inválido! Tente novamente.\n')

        while True:

            print('\nGostaria de remover ou editar?\n')
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

            else:
                print('Digite "remover" ou "editar" !')

    else:
        print('\nErro: Lista vazia!\n')

    return(divida)



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
                            opcao_editar = int(input('\nGostaria de editar qual dos tipos de intenção?\n1- Honra e Louvor\n2- Aniversário\n3- Aniversário de Casamento\n4- Ação de Graças\n5- Intenções\n6- Recuperação de Saúde\n7- 7º Dia\n8- Almas\n-> '))
                            
                            lista = None

                            match(opcao_editar):

                                case 1:
                                    lista = honra_e_louvor

                                case 2:
                                    lista = aniversario

                                case 3:
                                    lista = aniversario_de_casamento

                                case 4:
                                    lista = acao_de_gracas

                                case 5:
                                    lista = intencoes

                                case 6:
                                    lista = recuperacao_de_saude

                                case 7:
                                    lista = setimo_dia

                                case 8:
                                    lista = almas

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