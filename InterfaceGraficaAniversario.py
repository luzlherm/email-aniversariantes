from tkinter import *
from tkinter import ttk # para criar treeview
from tkinter import messagebox
import win32com.client as win32

#Preparando o outlook
outlook = win32.Dispatch("outlook.application")

#tk - Biblioteca do tkinter
#Tk - Janela / Tela
janela = Tk()

#Deixa o tamanho da janela fixo
#janela.geometry("750x350")

#Temas = clam, alt, default, classic
stilo = ttk.Style()
stilo.theme_use("clam")
stilo.configure(".", font=("Arial 18"), rowheight=30)

treeviewDados = ttk.Treeview(janela, columns=(1, 2, 3), show="headings")

treeviewDados.column("1", anchor=CENTER)
treeviewDados.heading("1", text="Nome")

treeviewDados.column("2", anchor=CENTER)
treeviewDados.heading("2", text="Nascimento")

treeviewDados.column("3", anchor=CENTER)
treeviewDados.heading("3", text="Email")

#grid - Divide a tela em grades / partes
#row - Linha
#column - Coluna
#columnspan - Colocamos para dizer quantas colunas do grid o item vai oculpar
#sticky - Usamos para preencher os espaços em bracos
#sticky - NSEW (Norte, Sul, Leste e Oeste)
treeviewDados.grid(row=2, column=0, columnspan=8, sticky="NSEW")

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

    #Email
    #print(arquivoAniversario.iloc[linha, 2])
    #Populando os itens na Treeview com os dados dos aniversariantes do dia
    treeviewDados.insert("", "end",
                         values=(str( arquivoAniversario.iloc[linha, 0] ), #Nome
                                 str( arquivoAniversario.iloc[linha, 1] ), #Aniversario
                                 str( arquivoAniversario.iloc[linha, 2] ))) #Email

#Criando a função que vai deletar o item
def deletarItemTreeview():

    itens = treeviewDados.selection()

    #for - para
    for item in itens:

        #Deletando o item que está selecionado
        treeviewDados.delete(item)

        messagebox.showinfo(title="Atenção!", message="Nome deletado com sucesso!")

        # Chama a função que conta linhas
        contarNumeroLinhas()

#Criando o botão de deletar
botaoDeletar = Button(text="Deletar",
                      font="Arial 20",
                      command = deletarItemTreeview)

botaoDeletar.grid(row=1, column=0, columnspan=2, sticky="NSEW")

#Label que exibe informações para o usuário
nome = Label(text="Nome: ", font="Arial 12")
nome.grid(row=0, column=0, sticky="W")

#Campo de entrada de dados
exibirNome = Entry(font="Arial 12")
exibirNome.grid(row=0, column=1, sticky="W")

#Label que exibe informações para o usuário
nascimento = Label(text="Nascimento: ", font="Arial 12")
nascimento.grid(row=0, column=2, sticky="W")

#Campo de entrada de dados
exibirNascimento = Entry(font="Arial 12")
exibirNascimento.grid(row=0, column=3, sticky="W")

#Label que exibe informações para o usuário
email = Label(text="Email: ", font="Arial 12")
email.grid(row=0, column=4, sticky="W")

#Campo de entrada de dados
exibirEmail = Entry(font="Arial 12")
exibirEmail.grid(row=0, column=5, sticky="W")

#Criando a função para adicionar um novo item na Treeview
def addItemTreeview():

    #if - se
    #elif - Senão Se
    if exibirNome.get() == "":
        messagebox.showinfo(title="Atenção!", message="Por favor digite um nome!")

    elif exibirNascimento.get() == "":
        messagebox.showinfo(title="Atenção!", message="Por favor digite uma data de nascimento!")

    elif exibirEmail.get() == "":
        messagebox.showinfo(title="Atenção!", message="Por favor digite um email!")

    else:

        #Cadastra um item no final
        treeviewDados.insert("", "end",
                             values=(str(exibirNome.get()),
                                     str(exibirNascimento.get()),
                                     str(exibirEmail.get())))

        messagebox.showinfo(title="Atenção!", message="Nome cadastrado com sucesso!")

        # Chama a função que conta linhas
        contarNumeroLinhas()

        #Limpando os campos digitaveis
        exibirNome.delete(0, "end")
        exibirNascimento.delete(0, "end")
        exibirEmail.delete(0, "end")

#Criando o botão de deletar
botaoAdicionar = Button(text="Adicionar",
                      font="Arial 20",
                      command = addItemTreeview)

botaoAdicionar.grid(row=1, column=2, columnspan=2, sticky="NSEW")

def alterarItemTreeview():

    # if - se
    # elif - Senão Se
    if exibirNome.get() == "":
        messagebox.showinfo(title="Atenção!", message="Por favor digite um nome!")

    elif exibirNascimento.get() == "":
        messagebox.showinfo(title="Atenção!", message="Por favor digite uma data de nascimento!")

    elif exibirEmail.get() == "":
        messagebox.showinfo(title="Atenção!", message="Por favor digite um email!")

    else:

        #Alterao os dados
        itemSelecionado = treeviewDados.selection()[0]
        treeviewDados.item(itemSelecionado,
                           values=(str(exibirNome.get()),
                                   str(exibirNascimento.get()),
                                   str(exibirEmail.get())))

        messagebox.showinfo(title="Atenção!", message="Nome alterado com sucesso!")

        # Limpando os campos digitaveis
        exibirNome.delete(0, "end")
        exibirNascimento.delete(0, "end")
        exibirEmail.delete(0, "end")


#Criando o botão de deletar
botaoAlterar = Button(text="Alterar",
                      font="Arial 20",
                      command = alterarItemTreeview)

botaoAlterar.grid(row=1, column=4, columnspan=2, sticky="NSEW")

def criarEmail():

    #for - para
    for numeroLinha in treeviewDados.get_children():

        #Pego os dados da linha que estiver passando / seleciona naquele momento
        #dadosDaLinha = treeviewDados.item(numeroLinha)["values"]
        #print(dadosDaLinha)

        #Criar um email em branco / Novo email
        emailOutlook = outlook.CreateItem(0)

        nome = treeviewDados.item(numeroLinha)["values"][0]
        aniversario = treeviewDados.item(numeroLinha)["values"][1]
        email = treeviewDados.item(numeroLinha)["values"][2]

        #split quebra o nome em colunas de acordo com o criterio
        #Pego a posicao 0 que é a primeira coluna
        variavelNome = nome.split(" ")[0]

        #<b> - Negrito / Bold
        #<font color="blue"> - Alterando a cor letra
        #<a href="" - Para colocar um hyperlink
        #img src= - Para colocar uma imagem
        emailOutlook.To = email #quem recebe o email
        emailOutlook.Subject = "Feliz Aniversário " + str(nome) #titulo da mensagem
        emailOutlook.HTMLBody = f"""
        <p>Parabéns, <b>{variavelNome}</b>!</p>
        <p><font color="green">Esse é um dia especial, aproveite seu dia!</font></p>
        <p><a href="https://www.google.com.br/">Clique aqui para acessar seu presente.</a></p>
        <p>Atenciosamente.</p>
        <p><img src="A:\\TKinter\\Email Aniversarios\\Assinatura_Email.jpeg">.</p>
        """

        # save - Salvar como rascunho / draft
        # send - envia
        emailOutlook.save()

    messagebox.showinfo(title="Atenção!", message="Emails criados com sucesso!")


#Criando o botão de deletar
botaoCriarEmail = Button(text="Criar Email",
                      font="Arial 20",
                      command = criarEmail)

botaoCriarEmail.grid(row=1, column=6, columnspan=2, sticky="NSEW")


labelNumeroLinhas = Label(text="Linhas: ", font="Arial 20")

#Criando o labelNumeroLinhas na linha 4 que é embaixo da treeview
labelNumeroLinhas.grid(row=4, column=0, columnspan=8, sticky="W")

#Função que conta o total de linhas / Aniversariantes
def contarNumeroLinhas(item=""):

    numero = 0

    linhas = treeviewDados.get_children(item)

    for linha in linhas:

        #numero = numero + 1
        numero += 1

    #Imprimindo o total de linhas
    labelNumeroLinhas.config(text="Aniversariantes: " + str(numero))

#Chama a função que conta linhas
contarNumeroLinhas()

def passaDadosParaEntry(Event):

    item = treeviewDados.selection()

    for linha in item:

        #Limpa os campos
        exibirNome.delete(0, END)
        exibirNascimento.delete(0, END)
        exibirEmail.delete(0, END)

        #Passando os itens da treeviewDados para os campos digitaveis
        exibirNome.insert(0, treeviewDados.item(linha, "values")[0])
        exibirNascimento.insert(0, treeviewDados.item(linha, "values")[1])
        exibirEmail.insert(0, treeviewDados.item(linha, "values")[2])

#Programa o evento de duplo clique para chamar a função que passa os dados para os Entry
treeviewDados.bind("<Double-1>", passaDadosParaEntry)

#mainloop - No tkinter é uma janela em funcionamento como um loop
#A janela que o python mostra na verdade é um programa em funcionamento
janela.mainloop()