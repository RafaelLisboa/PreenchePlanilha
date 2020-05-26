from openpyxl import *
from datetime import *

#No caso desta planilha, este dicionario relaciona numeros com letras referente a coluna
def buscaCol():
    #Busca o dia em que o script foi executado e retorna para quem chamou
    dia = datetime.today().day
    dias = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G', 6: 'H', 7: 'I', 8: 'J', 9: 'K', 10: 'L', 11: 'M',
            12: 'N', 13: 'O', 14: 'P', 15: 'Q', 16: 'R', 17: 'S', 18: 'T', 19: 'U', 20: 'V', 21: 'W',
            22: 'X', 23: 'Y', 24: 'Z', 25: 'AA', 26: 'AB', 27: 'AC', 28: 'AD', 29: 'AE', 30: 'AF', 31: 'AG'}

    return dias[dia]

#Classe Planilha com metodo contrutor com atributos relacionados a planilha, como seu nome por exemplo
class Planilha:

    def __init__(self, lista, nomeplan):
        self.lista = lista
        self.nomedaplan = nomeplan
        try:
            self.wb = load_workbook(self.nomedaplan)
        except:
            pass
        else:
            self.coluna = buscaCol()
            self.linha = 2
            self.intervaloFolhas = self.wb['Sheet1']


    """Metodo que pega o codigo do produto no arquivo de texto
    e vai percorrendo todas a linhas da planilha procurando um codigo compativel
    caso seja compativel ele preenche a quantidade de vendas do arquivo txt na linha do codigo e na coluna
    referente a o dia"""

    def preencher(self):
        vendas = self.lista
        for cod, quant in enumerate(vendas):
            linha = 2
            codigo = quant[0]
            encontrada = False
            while not encontrada and linha < 50:
                if self.intervaloFolhas[f'A{linha}'].value == codigo:
                    encontrada = True
                    self.intervaloFolhas[f'{self.coluna}{linha}'].value = quant[2]
                linha+=1
        self.wb.save('Vendas.xlsx')















