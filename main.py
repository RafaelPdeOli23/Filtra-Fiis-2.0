import requests
from bs4 import BeautifulSoup
import locale
from openpyxl import Workbook
import tabulate
from openpyxl.worksheet.table import Table, TableStyleInfo
from modelos import FundoImobiliario, Estrategia
import os

#Criando a pasta "saida"
if os.path.exists('./saida'):
    pass
else:
    dir = './saida'
    os.makedirs(dir)

#Tratamento dos numeros
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def trata_porcentagem(porcentagem_str):
    return  locale.atof(porcentagem_str.split('%')[0])

def trata_decimal(decimal_str):
    return locale.atof(decimal_str)

#Coleta dos Dados
headers = {'User-agent':'Mozilla/5.0'}
r = requests.get('https://www.fundamentus.com.br/fii_resultado.php', headers=headers)

soup = BeautifulSoup(r.text,'html.parser')

linhas = soup.find(id = "tabelaResultado").find('tbody').find_all('tr')

#Lista com todos os fundos filtrados
resultado = []

#Estratégia de investimento definida de forma abritária
estrategia = Estrategia(
    segmento='',
    dividend_yield_minimo=8,
    p_pv_maximo=1,
    liquidez_minima=800000,
    vacancia_media_maxima=10
)

#Intera na lista com todos os fundos, colocando os devidos valores às variáveis abaixo
for linha in linhas:
    dados_fundo = linha.find_all('td')
    codigo = dados_fundo[0].text
    segmento = dados_fundo[1].text
    cotacao_atual = trata_decimal(dados_fundo[2].text)
    dividend_yield = trata_porcentagem(dados_fundo[4].text)
    p_pv = trata_decimal(dados_fundo[5].text)
    liquidez = trata_decimal(dados_fundo[7].text)
    vacancia_media = trata_porcentagem(dados_fundo[12].text)

    fundo_imobiliario = FundoImobiliario(codigo, segmento, cotacao_atual, dividend_yield,
                                         p_pv, liquidez, vacancia_media)

    #Filtro aplicado de acordo com a estratégia predefinida
    if estrategia.aplica_estrategia(fundo_imobiliario):
       resultado.append(fundo_imobiliario)

#Esqueleto da tabela
cabecalho = ['CÓDIGO', 'SEGMENTO', 'COTAÇÃO ATUAL', 'DIVIDEND YIELD', 'P/PV']
tabela = []

#Colocando os dados filtrados dentro da tabela
for elemento in resultado:
    tabela.append([
        elemento.codigo, elemento.segmento,
        locale.currency(elemento.cotacao_atual),
        f'{locale.str(elemento.dividend_yield)} %', elemento.p_pv
    ])

#Retorno da tabela dentro do console Python
print(tabulate.tabulate(tabela, headers=cabecalho,showindex='always', tablefmt='fancy_grid'))

#Criando uma planillha Excel
workbook = Workbook()
planilha_ativa = workbook.active
planilha_ativa.title = 'FIIs'

#Colocando o cabeçalho da tabela na planilha
planilha_ativa.append(cabecalho)

#Colocando os dados para dentro do Excel
indice = 2

for elemento in resultado:
    planilha_ativa[f'A{indice}'] = elemento.codigo
    planilha_ativa[f'B{indice}'] = elemento.segmento
    planilha_ativa[f'C{indice}'] = elemento.cotacao_atual
    planilha_ativa[f'D{indice}'] = elemento.dividend_yield
    planilha_ativa[f'E{indice}'] = elemento.p_pv

    indice += 1

#Formatando como tabela
tab = Table(displayName="Table1", ref=F"A1:E{indice-1}")
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
tab.tableStyleInfo = style
planilha_ativa.add_table(tab)

#Efetivamente criando o arquivo Xlsx
workbook.save('./saida/Planilha.xlsx')
print('Planilha Excel criada!')
