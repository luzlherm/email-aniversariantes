import pandas as pd
from datetime import date
import numpy as np

#arquivoAniversario = pd.read_excel("A:\\TKinter\\Email Aniversarios\\Aniversario.xlsx")
arquivoAniversario = pd.read_excel("Aniversario.xlsx")

#Converte a coluna Nascimento para texto
arquivoAniversario["Nascimento"] = arquivoAniversario["Nascimento"].astype(str)

#Criando a coluna Ano
arquivoAniversario["Ano"] = arquivoAniversario["Nascimento"].str[:4]

#Criando a coluna Mes
arquivoAniversario["Mes"] = arquivoAniversario["Nascimento"].str[5:7]

#Criando a coluna Dia
arquivoAniversario["Dia"] = arquivoAniversario["Nascimento"].str[-2:]

#Cria uma coluna com a data atual que é a data de hoje
arquivoAniversario["Data Atual"] = date.today()

#Converte a coluna Data Atual para texto
arquivoAniversario["Data Atual"] = arquivoAniversario["Data Atual"].astype(str)

#Criando a coluna Ano Atual
arquivoAniversario["Ano Atual"] = arquivoAniversario["Data Atual"].str[:4]

#Criando a coluna Mes Atual
arquivoAniversario["Mes Atual"] = arquivoAniversario["Data Atual"].str[5:7]

#Criando a coluna Dia Atual
arquivoAniversario["Dia Atual"] = arquivoAniversario["Data Atual"].str[-2:]

#Comparando mes e dia e descobrindo os aniversariantes
arquivoAniversario["Aniversario"] = np.where((arquivoAniversario["Mes"] == arquivoAniversario["Mes Atual"]) &
                                             (arquivoAniversario["Dia"] == arquivoAniversario["Dia Atual"]), "Sim", "")

#loc - Localiza e limita por um critério
#Filtra e deixa somente os aniversariantes do dia
arquivoAniversario = arquivoAniversario.loc[arquivoAniversario["Aniversario"] != "", ["Nome", "Nascimento", "Email"]]

#print(arquivoAniversario)

#for - para
for linha in range(len(arquivoAniversario)):

    print(" Hoje é aniversário de:\n",arquivoAniversario.iloc[linha, 2,])
    
