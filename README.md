# PreenchePlanilha
O projeto visa abrir um arquivo txt, ler os dados de venda e preencher uma planilha já previamente tratada com celulas definidas no proprio script.

## O arquivo executavel precisa estar na mesma pasta da planilha e do arquivo txt e precisa estar acompanhado da pasta Imagens. 

A linguagem escolhida foi Python pela maior familiaridade com comandos e paradigmas da mesma.

Este projeto foi feito apenas para fins de aprendizado

OBJETIVOS

O script deve abrir o arquivo e fecha-lo sem perder nenhum dado nele contido
Deve interpretar os dados de vendas e armazena-los de forma simples antes de implementa-los na planilha
Deve abrir a planilha com as celulas a serem preenchidas já  especificadas para o funcionamento completo
Reconhecer a coluna a ser preenchida com base no dia em que o script esta sendo executado

Inicialmente estou disponibilizando uma planilha modelo, além do arquivo txt.

- O arquivo Excpy.py é quem cria a interface grafica da aplicação e coleta dados como o nome da Planilha e do Arquivo de texto para envia-los aos módulos

- O arquivo Excel.py é quem preenche a planilha com dados do arquivos txt que foram tratados pelo modulo Manipula

- O arquivo Manipula.py é quem trata os dados do arquivos txt e armazena eles em uma lista e entrega a mesma para o arquivo Excpy, que depois é enviado para o Excel.py que preenche a planilha.

- O Excpy.exe é o executavel da aplicação, inicialmente somente para o Windows
