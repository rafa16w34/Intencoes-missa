
from config import categorias, registro
from pagamento import menu


#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

def adicionar_intencao(categoria, texto):

    linhas = quebrar_linha(texto.title())

    for linha in linhas:
        categorias[categoria].append(linha)


#FUNÇÃO DE MARCAR MISSA

def menu_pessoa(valor_pessoa):

    sair_menu2 = 0

    categoria = None
    texto = None

    while sair_menu2 == 0:#Menu para marcar a intenção de uma pessoa

        try:

            opcao = int(input('\nDigite uma das opções abaixo:\n1- Honra e Louvor\n2- Aniversário\n3- Aniversário de Casamento\n4- Ação de Graças\n5- Intenções\n6- Recuperação de Saúde\n7- 7º Dia\n8- Almas\n------------------------------------------------------------\n9- Voltar\n-> '))

            match(opcao):

                case 1:#Adiciona uma intenção de Honra e Louvor

                    categoria = "HONRA E LOUVOR"

                    texto = '\nDigite a intenção em Honra e Louvor:\n-> '


                case 2:#Adiciona uma intenção para Aniversário

                    categoria = "ANIVERSÁRIO NATALÍCIO"

                    texto = '\nDigite o nome do aniversariante:\n-> '


                case 3:#Adiciona uma intenção de Aniversário de Casamento
                    
                    categoria = "ANIVERSÁRIO DE CASAMENTO"

                    texto = '\nDigite o nome do casal aniversariante:\n-> '



                case 4:#Adiciona uma intenção de Ação de Graças
                    
                    categoria = "AÇÃO DE GRAÇAS"
                            
                    texto = '\nDigite a intenção de acão de graças:\n-> '



                case 5:#Adiciona uma intenção de Intenções
                    
                    categoria = "INTENÇÃO"
                            
                    texto = '\nDigite a intenção:\n-> '




                case 6:#Adiciona uma intenção de Recuperação de Saúde
                    
                    categoria = "RECUPERAÇÃO DE SAÚDE"

                    texto = '\nDigite o nome do enfermo:\n-> '

                    


                case 7:#Adiciona uma intenção de Sétimo dia
                    
                    categoria = "7º DIA"

                    texto = '\nDigite o nome do sétimo dia:\n-> '

                   

                case 8:#Adiciona uma intenção de Alma
                    
                    categoria = "IRMÃOS FALECIDOS"

                    texto = '\nDigite o nome da alma:\n-> '

                    

                case 9:

                    if (valor_pessoa > 0):

                        sair_menu2 = menu(valor_pessoa)

                    else:#Voltar para o menu caso nao tenha marcado nenhuma intenção

                        confirmacao = str(input('\nAperte enter para continuar\n'))

                        print('\nVoltando ao menu inicial...\n')

                        sair_menu2 = 1


                case _: 

                    print('\nErro: Digite uma das opções mostradas no menu!\n')
                    
                    

        except(ValueError):
            print('\nErro: Digite uma das opções mostradas no menu!\n')
            continue


        if opcao != 9:

            nome = str(input(texto))
            nome = nome.lower()

            adicionar_intencao(categoria, nome)
            valor_pessoa += 3

    
    return(valor_pessoa)

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

            elif (edicao == len(lista)):
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


                registro.append(f"Intenção removida - R$ {divida:.2f} \n")


                break

            elif opcao_edicao == 'editar':
                nova_intencao = input('\nDigite a nova intenção:\n-> ')
                lista[edicao] = nova_intencao.title()
                print('\nIntenção editada com sucesso!\n')
                break

            elif opcao_edicao == 'sair':
                break

            else:
                print('Digite "remover", "editar" ou "sair" !')

    else:
        print('\nErro: Lista vazia!\n')

    return(divida)


def quebrar_linha(texto, limite=100):
    palavras = texto.split()
    linhas = []
    linha_atual = ""

    for palavra in palavras:
        # testa se a palavra cabe na linha atual
        if len(linha_atual) + len(palavra) + 1 <= limite:
            if linha_atual:
                linha_atual += " "
            linha_atual += palavra
        else:
            # fecha a linha atual com "-"
            linhas.append(linha_atual + "-")
            linha_atual = "-" + palavra

    # adiciona a última linha
    if linha_atual:
        linhas.append(linha_atual)

    return linhas

