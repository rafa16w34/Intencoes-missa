
from intencao import menu_pessoa
from intencao import editar
from imprimir import imprimir
from leitura import extrair_docx
from config import categorias
from config import valor_total

#-------------------------------------------------------------------------------------------------------------

#MENU PRINCIPAL


extrair_docx()

while True:
    try:
        opcao = int(input(
            '\nDigite uma das opções a seguir:\n'
            '1- Marcar intenção\n'
            '2- Dados\n'
            '3- Editar\n'
            '4- Imprimir\n'
            '0- Finalizar programa\n-> '
        ))
    except ValueError:
        print('\nErro: Digite um número válido!\n')
        continue

    match opcao:

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

            print(f'\nO valor total do dia é de: R$ {valor_total:.2f}\n')


        case 0:

            confirmacao = str(input('\nDigite "sim" para sair.\n(Tudo o que foi digitado será perdido caso não tenha sido impresso!)\n-> '))

            if (confirmacao == 'sim'):
                print('\nFinalizando programa...\n')
                break



        case _: 

            print('\nErro: Digite uma das opções mostradas no menu!\n')
