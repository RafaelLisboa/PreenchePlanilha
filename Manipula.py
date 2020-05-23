from time import sleep

#Função que tenta abrir o arquivo .txt, se ela conseguir, retorna verdadeiro, se não retorna falso
def abreArq(arq):
    try: 
        file = open(arq, 'r')
    except (FileExistsError, FileNotFoundError):
        print ('Arquivo de texto não encontrado'.upper())
        sleep(2)
        return False
    else:
        file.close()
        print('Arquivo encontrado com sucesso'.upper())
        print('-'*40)
        sleep(2)
        return True

#Separa cada linha do arquivo de texto, colocando elas em uma lista e removendo espaços e o "\n"
#Transforma também os codigos e quantidades de string para o tipo int
def itensVendas(arq):
    listaVendas = []
    for linha in arq:
        linha = linha.strip().split()
        listaVendas.append(linha)
    for prod, qtd in enumerate(listaVendas):
        cod = qtd[0]
        qtd[0] = int(cod)
        venda = qtd[2]
        qtd[2] = int(venda)
    return listaVendas





