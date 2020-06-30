from tkinter import *
from Excel import Planilha
import Manipula as m
from tkinter import ttk
from tkinter import filedialog

# Iniciando variaveis globais
planilha_buscada = None
planilha = None
arquivo = None
texto_buscado = None
verifica = None


def searchTxt():
    global texto_buscado
    texto_buscado = filedialog.askopenfilename(initialdir="Desktop", title="Selecione o Arquivo de texto",
                                               filetypes=(("Texto", "*.txt"), ("Todos os Arquivos", "*.*")))
    entradtela['text'] = "Arquivo encontrado"
    search_txt['state'] = DISABLED


def searchPlan():
    global planilha_buscada
    planilha_buscada = filedialog.askopenfilename(initialdir="Desktop", title="Selecione o Arquivo de texto",
                                                  filetypes=(("Planilha", "*.xlsx"), ("Todos os Arquivos", "*.*")))
    entradtela2['text'] = "Arquivo encontrado"
    search_plan['state'] = DISABLED


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
    global arquivo, verifica, informa, planilha_buscada, texto_buscado
    plan = planilha_buscada
    texto = texto_buscado
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
entradtela = Label(janela, text=" ", fg='black', font='Arial 10 bold')
entradtela.place(x=155, y=52)
entradtela2 = Label(janela, text=" ", fg='black', font='Arial 10 bold')
entradtela2.place(x=155, y=130)

# Labels
saidatela = Label(janela, text="Arquivo de texto", fg='black', font='Arial 10 bold')
saidatela.place(x=36, y=25)
saidatela2 = Label(janela, text="Planilha", fg='black', font='Arial 10 bold')
saidatela2.place(x=36, y=100)
informa = Label(janela, text=" ", fg='black', font='Arial 10 bold')
informa.place(x=50, y=280)

# Botão
b_ok = Button(janela, text="OK", width=30, command=b_click)
b_ok.place(x=40, y=230)
search_txt = Button(janela, text="Buscar Arquivo", width=15, command=searchTxt)
search_txt.place(x=40, y=52)
search_plan = Button(janela, text="Buscar Arquivo", width=15, command=searchPlan)
search_plan.place(x=40, y=130)
# Imagem
imagem = PhotoImage(file="Imagens/imagem.png")
logoimage = Label(janela, image=imagem, bg='green')
logoimage.place(x=350, y=80)

janela.mainloop()
