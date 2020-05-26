from tkinter import *
from Excel import Planilha
import Manipula as m
from tkinter import ttk

# Iniciando variaveis globais
planilha = None
arquivo = None
verifica = None


# Funçao Chamada para preencher a planilha
def preencherTudo(texto, planilha):
    global verifica
    dados = open(texto)
    vendas = m.itensVendas(dados)
    planilha = Planilha(vendas, planilha)
    planilha.preencher()
    informa['text'] = 'Planilha Preenchida com sucesso'


#Função do Botão para coletar os textos digitados
def b_click():
    global arquivo, verifica, informa
    texto = entradtela.get()
    plan = entradtela2.get()
    if texto == '' or plan == '':
        informa["text"] = 'Por favor preencha os campos'
    else:
        arquivo = texto
        # Esta Função "Pergunta" a função se o arquivo esta acessivel
        verifica = m.abreArq(arquivo)
        if (verifica == True):
            informa["text"] = 'Arquivos encontrados'
            preencherTudo(texto, plan)


        else:
            informa["text"] = 'Um dos arquivos não foi encontrado'


# Cria a janela
janela = Tk()
janela.title("Excpy")
janela.geometry("500x300")
janela.iconbitmap(default="Imagens/msexcel_93695.ico")
janela.resizable(0, 0)

# Divisão de frames
direito = Frame(janela, width=210, height=300)
direito['bg'] = 'green'
direito.pack(side=RIGHT)

# Widgets
entradtela = ttk.Entry(janela, width=30)
entradtela.place(x=40, y=52)
entradtela2 = ttk.Entry(janela, width=30)
entradtela2.place(x=40, y=130)

# Labels
saidatela = Label(janela, text="Nome do arquivo txt", fg='black', font='Arial 10 bold')
saidatela.place(x=36, y=25)
saidatela2 = Label(janela, text="Nome da Planilha", fg='black', font='Arial 10 bold')
saidatela2.place(x=36, y=100)
informa = Label(janela, text=" ", fg='black', font='Arial 10 bold')
informa.place(x=50, y=280)

# Botão
b_ok = Button(janela, text="OK", width=30, command=b_click)
b_ok.place(x=40, y=230)

# Imagem
imagem = PhotoImage(file="Imagens/imagem.png")
logoimage = Label(janela, image=imagem, bg='green')
logoimage.place(x=350, y=80)


janela.mainloop()