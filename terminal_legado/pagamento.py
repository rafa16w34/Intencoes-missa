import config


def menu(valor_pessoa):
                    
    print(f'\nO valor a ser pago é de R$ {valor_pessoa:.2f}\n')

    metodo_de_pagamento = int(input('\nQual vai ser a forma de pagamento?\n1- Dinheiro\n2- Pix\n3- Voltar\n-> '))

    match(metodo_de_pagamento):

        case 1:#Dinheiro

            pagamento_confirmado = dinheiro(valor_pessoa)
            
        case 2:#Pix

            pagamento_confirmado = pix(valor_pessoa)
            
        case 3:#sair

            pagamento_confirmado = False

            return 0

        case _:

            print('\nErro: Digite uma das opções mostradas no menu!\n')
            pagamento_confirmado = False

            return 0


    if (pagamento_confirmado == True):

        confirmacao = str(input('\nAperte enter para continuar\n'))

        print('\nVoltando ao menu inicial...\n')

        return 1





def dinheiro(valor_pessoa):

    doacao_dinheiro = 0

    while True:
        
        try:

            valor_recebido = int(input('\nDigite o valor recebido:\n-> R$ '))

            if (valor_recebido < valor_pessoa):

                print('\nDigite um valor válido!\n')


            else:
                
                if (valor_recebido > valor_pessoa):

                    

                    doacao_dinheiro = troco(valor_recebido,valor_pessoa)
                    
                
                break

        
        
        except(ValueError):
            
            print('\nDigite um valor válido!\n')

    config.registro.append(f"Valor: R$ {valor_pessoa:.2f} -- Valor Pago: R$ {valor_recebido:.2f} (Troco: R$ {((valor_recebido - valor_pessoa)- doacao_dinheiro):.2f} -- Doação: R$ {(doacao_dinheiro):.2f}) -- Dinheiro\n")

    config.valor_dinheiro += valor_pessoa

    return True


def pix(valor_pessoa):


    doacao_pix = 0

    nome_pix = str(input("\nQual é o nome quem irá fazer o pix?\n-> "))

    nome_pix = nome_pix.lower()

    config.lista_pix.append(nome_pix.title())

    while True:

        try:

            valor_pix = int(input(f"\nInforme o valor pago no Pix (solicite o comprovante para {nome_pix.title()})\n->"))

            if (valor_pix < valor_pessoa):

                print('\nDigite um valor válido!\n')


            else:
                
                if (valor_pix > valor_pessoa):


                    doacao_pix = troco(valor_pix,valor_pessoa)
                    
                
                break

        except(ValueError):
                    
            print('\nDigite um valor válido!\n')

            
    config.lista_pix.append(valor_pix)

    with open ('output/Lista_Pix.txt','w') as arq:

        arq.write(f'Lista do Pix {config.data_formatada}:\n')

        for i in range(len(config.lista_pix)):

            if ((i+1)%2 != 0):          

                arq.write(f'- {config.lista_pix[i]} : R$ {config.lista_pix[i+1]},00\n')

                pix_registro = (f"Valor: R$ {valor_pessoa:.2f} -- Valor Pago: R$ {config.lista_pix[i+1]:.2f} (Troco: R$ {((config.lista_pix[i+1] - valor_pessoa)- doacao_pix):.2f} -- Doação: R$ {(doacao_pix):.2f}) -- Pix\n")

                i+2
            
            if i+2 >= len(config.lista_pix):
                break

    config.registro.append(pix_registro)

    config.valor_pix += valor_pessoa

    return True


def troco(valor,valor_pessoa):

    print(f'\nTroco: R$ {(valor - valor_pessoa):.2f}\n')
    
    while True:

        try:

            doacao = int(input(f"\nDigite o valor do troco que será destinado a doação:\n-> R$ "))

            if (doacao > (valor - valor_pessoa)):

                print("\nValor de doação maior do que o de troco!\n")

            else:

                print(f'\nTroco: R$ {((valor - valor_pessoa)- doacao):.2f} -- Doação: R$ {(doacao):.2f} \n')

                break

        except(ValueError):
                
                print('\nDigite um valor válido!\n')

    return doacao


