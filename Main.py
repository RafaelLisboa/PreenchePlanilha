import Manipula as m
from Excel import Planilha

iniciar = input('APERTE ENTER PARA INICIAR O SCRIPT')
arquivo = 'Vendas.txt'
# Esta Função "Pergunta" a função se o arquivo esta acessivel
verifica = m.abreArq(arquivo)

#Caso consiga ele começa a tratar os dados do arquivo .txt e o da planilha
if verifica == True:
    #Abre o arquivo de texto e cria uma instancia da classe Planilha
    dados = open (arquivo)
    vendas = m.itensVendas(dados)
    planilha = Planilha(vendas)
    planilha.preencher()
    finalizar = input('Aperte Enter Para Finalizar...'.upper())

else:
    print('Por favor coloque o arquivo de texto na mesma pasta que o script'.upper())
    input()