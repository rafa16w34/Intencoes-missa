from datetime import datetime, timedelta


#-------------------------------------------------------------------------------------------------------------

data_hoje  = datetime.now()

data_ajustada = data_hoje + timedelta(hours=1)#Soma 1 na hora porque temos que imprimir 10 minutos antes

data_formatada = data_ajustada.strftime("%d/%m/%Y - %H:00")

hora_formatada = data_ajustada.strftime("%Hh")

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

registro = []
lista_pix = []


#Variavel do dinheiro
valor_total = 0
valor_dinheiro = 0
valor_pix = 0
