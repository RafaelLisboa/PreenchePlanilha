from tkinter import *
from Excel import *
import Manipula as m
from tkinter import ttk

arquivo = None
verifica = None


def b_click():
    global arquivo, verifica, informa
    texto = entradtela.get()
    if texto == '':
        informa["text"] = 'Por favor preencha os campos'
    else:
        arquivo = texto
        # Esta Função "Pergunta" a função se o arquivo esta acessivel
        verifica = m.abreArq(arquivo)
        if (verifica == True):
            informa["text"] = 'Arquivo de texto encontrado'
        else:
            informa["text"] = 'Arquivo de texto não encontrado'


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

# Caso consiga ele começa a tratar os dados do arquivo .txt e o da planilha
# if verifica == True:
# Abre o arquivo de texto e cria uma instancia da classe Planilha
# dados = open (arquivo)
# vendas = m.itensVendas(dados)
# planilha = Planilha(vendas)
# planilha.preencher()


janela.mainloop()
