#CRIANDO UM ARQUIVO EXCEL COM FORMULAS DE SOMA, SUBTRAÇÃO, MULTIPLICAÇÃO, DIVISÃO, CONCATENANDO COLUNAS DE NOME E AJUSTANDO O TAMANHO DAS COLUNAS

import xlsxwriter as opcoesDoXlsxWriter
import os

#1 - indicando onde será criado o arquivo, seu nome e sua extensão. Importante a questão das barras duplas (testar).
nomeCaminhoArquivo = 'C:\\Users\\Windows\\Desktop\\Python Projetos\\xlsxwriter\Formulas.xlsx'
minhaPlanilha = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)
sheetDados = minhaPlanilha.add_worksheet("Dados") #Para renomear o nome da Sheet1 para Dados.


#Adicionando titulos
sheetDados.write("A1", "Número 1")
sheetDados.write("B1", "Número 2")
sheetDados.write("C1", "Fórmulas")

#Adicionando valores na coluna A
sheetDados.write("A2", 10)
sheetDados.write("A3", 6)
sheetDados.write("A4", 8)
sheetDados.write("A5", 6)
sheetDados.write("A8", "Ana")

#Adicionando valores na coluna B
sheetDados.write("B2", 7)
sheetDados.write("B3", 5)
sheetDados.write("B4", 3)
sheetDados.write("B5", 1)
sheetDados.write("B8", "Luiza")

#Adicionando fórmulas
sheetDados.write_formula("C2", "A2+B2") #fórmula para adição
sheetDados.write_formula("C3", "A3-B3") #fórmula para subtração
sheetDados.write_formula("C4", "A4*B4") #fórmula para multiplicação
sheetDados.write_formula("C5", "A2/B2") #fórmula para divisão
sheetDados.write_formula("C8", '=CONCATENATE(A8," ",B8)') #Aqui usamos para concatenar os nomes Ana + Luiza

#Coluna com tamanho 15
sheetDados.set_column('A:C', 15)

#3 - Para fechar e salvar as informações
minhaPlanilha.close()

#4 - Abrir o arquivo para verificar o resultado
os.startfile(nomeCaminhoArquivo)