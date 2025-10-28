from datetime import datetime

data_hoje  = datetime.now()

data_formatada = data_hoje.strftime("%d/%m/%Y - %H:00")

#Listas para cada uma das intenções
honra_e_louvor = []
aniversario = []
aniversario_de_casamento = []
acao_de_gracas = []
intencoes = []
recuperacao_de_saude = []
almas = []

#Variavel do dinheiro
valor_total = 0.0

sair = 0

while sair == 0:#Menu para o sistema das intenções

    try:

        opcao = int(input('\nDigite uma das opções abaixo:\n1- Honra e Louvor\n2- Aniversário\n3- Aniversário de Casamento\n4- Ação de Graças\n5- Intenções\n6- Recuperação de Saúde\n7- Almas\n8- Imprimir\n-> '))

        match(opcao):

            case 1:#Adiciona uma intenção de Honra e Louvor

                try:

                    numero = int(input('\nQuantas inteções gostaria de marcar?\n'));
                    

                    for i in range(numero):

                        nome = str(input('\nDigite a intenção em Honra e Louvor:\n-> '))

                        nome = nome.lower()

                        honra_e_louvor.append(nome.title())
                        valor_total += 3.0
                        
                    print(f'\nO valor a ser pago é de R$ {numero * 3:.2f}\n')

                except(ValueError):

                    print('\nErro: Digite um número inteiro!\n')


            case 2:#Adiciona uma intenção para Aniversário

                try:

                    numero = int(input('\nQuantas inteções gostaria de marcar?\n'));

                    for i in range(numero):
                        
                        nome = str(input('\nDigite o nome do aniversariante:\n-> '))

                        nome = nome.lower()

                        aniversario.append(nome.title())
                        valor_total += 3.0

                    print(f'\nO valor a ser pago é de R$ {numero * 3:.2f}\n')
                
                except(ValueError):

                    print('\nErro: Digite um número inteiro!\n')

            case 3:#Adiciona uma intenção de Aniversário de Casamento
                
                try:

                    numero = int(input('\nQuantas inteções gostaria de marcar?\n'));

                    for i in range(numero):
                        
                        nome = str(input('\nDigite o nome do casal aniversariante:\n-> '))

                        nome = nome.lower()

                        aniversario_de_casamento.append(nome.title())
                        valor_total += 3.0

                    print(f'\nO valor a ser pago é de R$ {numero * 3:.2f}\n')
                
                except(ValueError):

                    print('\nErro: Digite um número inteiro!\n')

            case 4:#Adiciona uma intenção de Ação de Graças
                
                try:

                    numero = int(input('\nQuantas inteções gostaria de marcar?\n'));

                    for i in range(numero):
                        
                        nome = str(input('\nDigite a intenção de acão de graças:\n-> '))

                        nome = nome.lower()

                        acao_de_gracas.append(nome.title())
                        valor_total += 3.0

                    print(f'\nO valor a ser pago é de R$ {numero * 3:.2f}\n')
                
                except(ValueError):

                    print('\nErro: Digite um número inteiro!\n')

            case 5:#Adiciona uma intenção de Intenções
                
                try:

                    numero = int(input('\nQuantas inteções gostaria de marcar?\n'));

                    for i in range(numero):
                        
                        nome = str(input('\nDigite a intenção:\n-> '))

                        nome = nome.lower()

                        intencoes.append(nome.title())
                        valor_total += 3.0

                    print(f'\nO valor a ser pago é de R$ {numero * 3:.2f}\n')
                
                except(ValueError):

                    print('\nErro: Digite um número inteiro!\n')


            case 6:#Adiciona uma intenção de Recuperação de Saúde
                
                try:

                    numero = int(input('\nQuantas inteções gostaria de marcar?\n'));

                    for i in range(numero):
                        
                        nome = str(input('\nDigite o nome do enfermo:\n-> '))

                        nome = nome.lower()

                        recuperacao_de_saude.append(nome.title())
                        valor_total += 3.0

                    print(f'\nO valor a ser pago é de R$ {numero * 3:.2f}\n')
                
                except(ValueError):

                    print('\nErro: Digite um número inteiro!\n')


            case 7:#Adiciona uma intenção de Alma
                
                try:

                    numero = int(input('\nQuantas inteções gostaria de marcar?\n'));

                    for i in range(numero):
                        
                        nome = str(input('\nDigite o nome da alma:\n-> '))

                        nome = nome.lower()

                        almas.append(nome.title())
                        valor_total += 3.0

                    print(f'\nO valor a ser pago é de R$ {numero * 3:.2f}\n')
                
                except(ValueError):

                    print('\nErro: Digite um número inteiro!\n')



            case 8:#Imprime a lista formatada

                with open("Intenções.txt",'a') as arq:

                    arq.write(f'\nINTENÇÕES DA MISSA\t{data_formatada}\n')

                ###################################################################################
                    arq.write('\nHONRA E LOUVOR :\n')

                    for i in range(len(honra_e_louvor)):

                        arq.write(f'\n- {honra_e_louvor[i]}\n')

                ###################################################################################
                    arq.write('\nANIVERSARIANTES :\n')

                    for i in range(len(aniversario)):

                        arq.write(f'\n- {aniversario[i]}\n')

                ###################################################################################
                    arq.write('\nANIVERSÁRIO DE CASAMENTO :\n')

                    for i in range(len(aniversario_de_casamento)):

                        arq.write(f'\n- {aniversario_de_casamento[i]}\n')

                ###################################################################################
                    arq.write('\nAÇÃO DE GRAÇAS :\n')

                    for i in range(len(acao_de_gracas)):

                        arq.write(f'\n- {acao_de_gracas[i]}\n')

                ###################################################################################
                    arq.write('\nINTENÇÕES :\n')

                    for i in range(len(intencoes)):

                        arq.write(f'\n- {intencoes[i]}\n')

                ###################################################################################
                    arq.write('\nRECUPERAÇÃO DE SAÚDE :\n')

                    for i in range(len(recuperacao_de_saude)):

                        arq.write(f'\n- {recuperacao_de_saude[i]}\n')

                ##################################################################################
                    arq.write(f'\nIRMÃOS FALECIDOS :\n')

                    for i in range(len(almas)):

                        arq.write(f'\n+ {almas[i]}\n')

                print(f'\nO valor total do dia é de: R$ {valor_total:.2f}\n')
                
                sair = 1


            case _: 

                print('\nErro: Digite uma das opções mostradas no menu!\n')


    except(ValueError):
        print('\nErro: Digite uma das opções mostradas no menu!\n')