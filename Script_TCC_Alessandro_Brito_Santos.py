from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.select import Select
from email.mime.multipart import MIMEMultipart
from selenium.webdriver.common.by import By
from datetime import date,timedelta
import matplotlib.pyplot as plt
from selenium import webdriver
import win32com.client
from sys import exit
import smtplib, ssl
import pandas as pd
import numpy as np
import time
import glob
import os
from reportlab.platypus import SimpleDocTemplate, Table, Image, Spacer, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from io import BytesIO
from reportlab.lib.pagesizes import letter
import seaborn as sns

inicio = time.time()

cnpj = str(input('Insira o cnpj do fundo: '))

options = webdriver.ChromeOptions()
options.headless = False

chrome_driver_path = 'C:\\Users\\aless\\Downloads\\chromedriver-win64\\chromedriver.exe'

driver = webdriver.Chrome(executable_path=chrome_driver_path, options=options)

params = {'behavior': 'allow', 'downloadPath': 'C:\\Users\\alessandro.santos\\Downloads'}
driver.execute_cdp_cmd('Page.setDownloadBehavior', params)

driver.get('https://cvmweb.cvm.gov.br/swb/default.asp?sg_sistema=fundosreg')
driver.maximize_window()
time.sleep(2)

driver.switch_to.frame("Main")

driver.find_element(By.XPATH, '//*[@id="txtCNPJNome"]').send_keys(cnpj)

driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[7]/td/input').click()

driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr/td[2]/a').click()

nome = driver.find_element(By.XPATH, '/html/body/form/table[1]/tbody/tr[2]/td[1]/span').text

data_inicio = driver.find_element(By.XPATH, '/html/body/form/table[1]/tbody/tr[4]/td[2]/span').text

driver.find_element(By.XPATH, '//*[@id="Hyperlink2"]').click()

todas_opcoes = driver.find_elements_by_css_selector('option')

opcoes_selecionadas = len(todas_opcoes)

print("Número de opções selecionadas:", opcoes_selecionadas)

i = 1

dados = []

while i <= opcoes_selecionadas:
    
    driver.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/select/option["+str(i)+"]").click()
    
    tabela = driver.find_element(By.ID, "dgDocDiario")
    
    linhas = tabela.find_elements(By.TAG_NAME, 'tr')

    num_linhas = len(linhas)+1
    
    # Iterar sobre as linhas da tabela
    for k in range(2, num_linhas):
        
        linha = driver.find_element(By.XPATH, f"//*[@id='dgDocDiario']/tbody/tr[{k}]")
        
        # Verificando o comprimento da linha
        if len(linha.text) < 17:
            continue  # Se o comprimento for menor que 17, vá para a próxima iteração do loop k
    
        linha_dados = []
        
        for j in range(1, 9):
        
            linha_dados.append(linha.find_element(By.XPATH, f"td[{j}]").text.strip())
            
        mes_ano = driver.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/select/option["+str(i)+"]").text
    
        linha_dados[0] += f"/{mes_ano}"
        dados.append(linha_dados)

    i = i + 1
    
df = pd.DataFrame(dados, columns=["Dia", "Quota", "Captação no Dia", "Resgate no Dia", "Patrimônio Líquido", "Total da Carteira", "N°. Total de Cotistas", "Data da próxima informação do PL"])

driver.quit()

caminho_excel = r'C:\Users\aless\Downloads\dados_fundo.xlsx'

df.to_excel(caminho_excel, index=False)

df = pd.read_excel(caminho_excel)

df = df.drop('Data da próxima informação do PL', axis=1)

# Converter a coluna 'Dia' para o tipo datetime
df['Dia'] = pd.to_datetime(df['Dia'], format='%d/%m/%Y')

# Ordenar o DataFrame pela data
df = df.sort_values(by='Dia', ascending=False)

# Encontre a cota inicial
cota_inicial = df.loc[df.index[-1], 'Quota']

# Calcule a rentabilidade diária acumulada em relação à cota inicial
df['Rentabilidade'] = (df['Quota'] / cota_inicial - 1).map(lambda x: float(x))

caminho_excel_ibov = r'C:\Users\aless\Downloads\Dados Históricos - Ibovespa.xlsx'

df['Quota'] = df['Quota'].str.replace(',', '.').astype(float)
df['Captação no Dia'] = df['Captação no Dia'].str.replace('.', '').str.replace(',', '.').astype(float)
df['Resgate no Dia'] = df['Resgate no Dia'].str.replace('.', '').str.replace(',', '.').astype(float)
df['Patrimônio Líquido'] = df['Patrimônio Líquido'].str.replace('.', '').str.replace(',', '.').astype(float).astype(float)
df['Total da Carteira'] = df['Total da Carteira'].str.replace('.', '').str.replace(',', '.').astype(float)
df['N°. Total de Cotistas'] = df['N°. Total de Cotistas'].astype(int)

df.to_excel(caminho_excel, index=False)

# Calculando as estatísticas descritivas
media = df['Quota'].mean()
desvio_padrao = df['Quota'].std()
quartil_25 = df['Quota'].quantile(0.25)
quartil_50 = df['Quota'].quantile(0.5)
quartil_75 = df['Quota'].quantile(0.75)
maximo = df['Quota'].max()
minimo = df['Quota'].min()

# Exibindo as estatísticas descritivas
print("Média:", media)
print("Desvio Padrão:", desvio_padrao)
print("Quartil 25:", quartil_25)
print("Quartil 50 (Mediana):", quartil_50)
print("Quartil 75:", quartil_75)
print("Máximo:", maximo)
print("Mínimo:", minimo)

# Criando o gráfico de box plot
plt.figure(figsize=(5, 3))
df.boxplot(column=['Quota'])
plt.title('Box plot do Quota')
plt.ylabel('Valor (R$)')
plt.show()

# Calculando as estatísticas descritivas
media = df['Captação no Dia'].mean()
desvio_padrao = df['Captação no Dia'].std()
quartil_25 = df['Captação no Dia'].quantile(0.25)
quartil_50 = df['Captação no Dia'].quantile(0.5)
quartil_75 = df['Captação no Dia'].quantile(0.75)
maximo = df['Captação no Dia'].max()
minimo = df['Captação no Dia'].min()

# Exibindo as estatísticas descritivas
print("Média:", media)
print("Desvio Padrão:", desvio_padrao)
print("Quartil 25:", quartil_25)
print("Quartil 50 (Mediana):", quartil_50)
print("Quartil 75:", quartil_75)
print("Máximo:", maximo)
print("Mínimo:", minimo)

# Criando o gráfico de box plot
plt.figure(figsize=(5, 3))
df.boxplot(column=['Captação no Dia'])
plt.title('Box plot do Captação no Dia')
plt.ylabel('Valor (R$)')
plt.show()

# Calculando as estatísticas descritivas
media = df['Resgate no Dia'].mean()
desvio_padrao = df['Resgate no Dia'].std()
quartil_25 = df['Resgate no Dia'].quantile(0.25)
quartil_50 = df['Resgate no Dia'].quantile(0.5)
quartil_75 = df['Resgate no Dia'].quantile(0.75)
maximo = df['Resgate no Dia'].max()
minimo = df['Resgate no Dia'].min()

# Exibindo as estatísticas descritivas
print("Média:", media)
print("Desvio Padrão:", desvio_padrao)
print("Quartil 25:", quartil_25)
print("Quartil 50 (Mediana):", quartil_50)
print("Quartil 75:", quartil_75)
print("Máximo:", maximo)
print("Mínimo:", minimo)

# Criando o gráfico de box plot
plt.figure(figsize=(5, 3))
df.boxplot(column=['Resgate no Dia'])
plt.title('Box plot do Resgate no Dia')
plt.ylabel('Valor (R$)')
plt.show()

# Calculando as estatísticas descritivas
media = df['Patrimônio Líquido'].mean()
desvio_padrao = df['Patrimônio Líquido'].std()
quartil_25 = df['Patrimônio Líquido'].quantile(0.25)
quartil_50 = df['Patrimônio Líquido'].quantile(0.5)
quartil_75 = df['Patrimônio Líquido'].quantile(0.75)
maximo = df['Patrimônio Líquido'].max()
minimo = df['Patrimônio Líquido'].min()

# Exibindo as estatísticas descritivas
print("Média:", media)
print("Desvio Padrão:", desvio_padrao)
print("Quartil 25:", quartil_25)
print("Quartil 50 (Mediana):", quartil_50)
print("Quartil 75:", quartil_75)
print("Máximo:", maximo)
print("Mínimo:", minimo)

# Criando o gráfico de box plot
plt.figure(figsize=(5, 3))
df.boxplot(column=['Patrimônio Líquido'])
plt.title('Box plot do Patrimônio Líquido')
plt.ylabel('Valor (R$)')
plt.show()

# Calculando as estatísticas descritivas
media = df['Total da Carteira'].mean()
desvio_padrao = df['Total da Carteira'].std()
quartil_25 = df['Total da Carteira'].quantile(0.25)
quartil_50 = df['Total da Carteira'].quantile(0.5)
quartil_75 = df['Total da Carteira'].quantile(0.75)
maximo = df['Total da Carteira'].max()
minimo = df['Total da Carteira'].min()

# Exibindo as estatísticas descritivas
print("Média:", media)
print("Desvio Padrão:", desvio_padrao)
print("Quartil 25:", quartil_25)
print("Quartil 50 (Mediana):", quartil_50)
print("Quartil 75:", quartil_75)
print("Máximo:", maximo)
print("Mínimo:", minimo)

# Criando o gráfico de boxplot
plt.figure(figsize=(5, 3))
df.boxplot(column=['Total da Carteira'])
plt.title('Box plot do Total da Carteira')
plt.ylabel('Valor (R$)')
plt.show()

# Calculando as estatísticas descritivas
media = df['N°. Total de Cotistas'].mean()
desvio_padrao = df['N°. Total de Cotistas'].std()
quartil_25 = df['N°. Total de Cotistas'].quantile(0.25)
quartil_50 = df['N°. Total de Cotistas'].quantile(0.5)
quartil_75 = df['N°. Total de Cotistas'].quantile(0.75)
maximo = df['N°. Total de Cotistas'].max()
minimo = df['N°. Total de Cotistas'].min()

# Exibindo as estatísticas descritivas
print("Média:", media)
print("Desvio Padrão:", desvio_padrao)
print("Quartil 25:", quartil_25)
print("Quartil 50 (Mediana):", quartil_50)
print("Quartil 75:", quartil_75)
print("Máximo:", maximo)
print("Mínimo:", minimo)

# Criando o gráfico de boxplot
plt.figure(figsize=(5, 3))
df.boxplot(column=['N°. Total de Cotistas'])
plt.title('Box plot do N°. Total de Cotistas')
plt.ylabel('Valor (R$)')
plt.show()

# Convertendo a coluna 'Dia' para formato de data
df['Dia'] = pd.to_datetime(df['Dia'], format='%d/%m/%Y')

# Extraindo mês e ano
df['Mes'] = df['Dia'].dt.month
df['Ano'] = df['Dia'].dt.year

# Calculando a rentabilidade
df_rentabilidade = pd.DataFrame(columns=['Ano', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro', 'Acum. Ano'])
for ano in df['Ano'].unique():
    rentabilidades = {}
    for mes in df[df['Ano'] == ano]['Mes'].unique():
        ultimo_valor_mes_atual = df[(df['Ano'] == ano) & (df['Mes'] == mes)]['Quota'].iloc[0]
        if mes == 1:
            ultimo_valor_mes_anterior = df[(df['Ano'] == ano - 1) & (df['Mes'] == 12)]['Quota'].iloc[0]
    
        else:
            mes_anterior = mes - 1
            if mes_anterior == 0:
                mes_anterior = 12
                mes_anterior = mes_anterior - 1
            else:
                ano_anterior = ano
            if len(df[(df['Ano'] == ano_anterior) & (df['Mes'] == mes_anterior)]) > 0:
                ultimo_valor_mes_anterior = df[(df['Ano'] == ano_anterior) & (df['Mes'] == mes_anterior)]['Quota'].iloc[0]
                
            else:
                ultimo_valor_mes_anterior = None
        if ultimo_valor_mes_anterior is not None:
            rentabilidades[mes] = (ultimo_valor_mes_atual / ultimo_valor_mes_anterior) - 1
    rentabilidades['Acum. Ano'] = sum(rentabilidades.values()) if len(rentabilidades.values()) > 0 else None
    df_rentabilidade = df_rentabilidade.append({'Ano': ano, 'Janeiro': rentabilidades.get(1), 'Fevereiro': rentabilidades.get(2), 'Março': rentabilidades.get(3), 'Abril': rentabilidades.get(4), 'Maio': rentabilidades.get(5), 'Junho': rentabilidades.get(6), 'Julho': rentabilidades.get(7), 'Agosto': rentabilidades.get(8), 'Setembro': rentabilidades.get(9), 'Outubro': rentabilidades.get(10), 'Novembro': rentabilidades.get(11), 'Dezembro': rentabilidades.get(12), 'Acum. Ano': rentabilidades['Acum. Ano']}, ignore_index=True)
    
# Aplicando a formatação apenas da segunda coluna em diante
for coluna in df_rentabilidade.columns[1:]:
    df_rentabilidade[coluna] = df_rentabilidade[coluna].apply(lambda x: '{:.2%}'.format(x) if isinstance(x, float) else x)
df_rentabilidade['Ano'] = df_rentabilidade['Ano'].astype(int)
df_rentabilidade.replace('nan%', '-', inplace=True)

df = pd.read_excel(r'C:\Users\aless\Downloads\Dados Históricos - Ibovespa.xlsx')

# Converter a coluna 'Dia' para o tipo datetime
df['Dia'] = pd.to_datetime(df['Dia'], format='%d/%m/%Y')

# Ordenar o DataFrame pela data
df = df.sort_values(by='Dia', ascending=False)

# Extraindo mês e ano
df['Mes'] = df['Dia'].dt.month
df['Ano'] = df['Dia'].dt.year

# Calculando a rentabilidade
df_rentabilidade_ib = pd.DataFrame(columns=['Ano', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro', 'Acum. Ano'])
df_rentabilidade_ib.dropna(subset=['Ano'], inplace=True)
for ano in df['Ano'].unique():
    rentabilidades = {}
    for mes in df[df['Ano'] == ano]['Mes'].unique():
        ultimo_valor_mes_atual = df[(df['Ano'] == ano) & (df['Mes'] == mes)]['Último'].iloc[0]
        if mes == 1:
            ultimo_valor_mes_anterior = df[(df['Ano'] == ano - 1) & (df['Mes'] == 12)]['Último'].iloc[0]
    
        else:
            mes_anterior = mes - 1
            if mes_anterior == 0:
                mes_anterior = 12
                mes_anterior = mes_anterior - 1
            else:
                ano_anterior = ano
            if len(df[(df['Ano'] == ano_anterior) & (df['Mes'] == mes_anterior)]) > 0:
                ultimo_valor_mes_anterior = df[(df['Ano'] == ano_anterior) & (df['Mes'] == mes_anterior)]['Último'].iloc[0]
                
            else:
                ultimo_valor_mes_anterior = None
        if ultimo_valor_mes_anterior is not None:
            rentabilidades[mes] = (ultimo_valor_mes_atual / ultimo_valor_mes_anterior) - 1
    rentabilidades['Acum. Ano'] = sum(rentabilidades.values()) if len(rentabilidades.values()) > 0 else None
    df_rentabilidade_ib = df_rentabilidade_ib.append({'Ano': ano, 'Janeiro': rentabilidades.get(1), 'Fevereiro': rentabilidades.get(2), 'Março': rentabilidades.get(3), 'Abril': rentabilidades.get(4), 'Maio': rentabilidades.get(5), 'Junho': rentabilidades.get(6), 'Julho': rentabilidades.get(7), 'Agosto': rentabilidades.get(8), 'Setembro': rentabilidades.get(9), 'Outubro': rentabilidades.get(10), 'Novembro': rentabilidades.get(11), 'Dezembro': rentabilidades.get(12), 'Acum. Ano': rentabilidades['Acum. Ano']}, ignore_index=True)
    
# Aplicando a formatação apenas da segunda coluna em diante
for coluna in df_rentabilidade_ib.columns[1:]:
    df_rentabilidade_ib[coluna] = df_rentabilidade_ib[coluna].apply(lambda x: '{:.2%}'.format(x) if isinstance(x, float) else x)
df_rentabilidade_ib['Ano'] = df_rentabilidade_ib['Ano'].astype(int)
df_rentabilidade_ib.replace('nan%', '-', inplace=True)

# Carregar os DataFrames
df_fundo = pd.read_excel(caminho_excel)
df_ibov = pd.read_excel(caminho_excel_ibov)

# Criar o documento PDF
pdf_filename = r'C:\Users\aless\Downloads\mba.pdf'
pdf = SimpleDocTemplate(pdf_filename, pagesize=letter)

# Lista para armazenar elementos do PDF
elements = []

# Estilos para o PDF
styles = getSampleStyleSheet()

# Adicionar Razão Social, CNPJ e Data inicial
elements.append(Paragraph('<b>Razão Social:</b> ' + nome, styles["Normal"]))
elements.append(Paragraph('<b>CNPJ:</b> ' + cnpj, styles["Normal"]))
elements.append(Paragraph('<b>Data inicial:</b> ' + data_inicio, styles["Normal"]))
elements.append(Spacer(1, 12))  # Espaço de 12 pontos

# Criar a imagem do gráfico de comparação entre Itaú e Ibovespa
plt.figure(figsize=(6, 4))  # Define o mesmo tamanho para o gráfico de comparação
plt.plot(df_fundo['Dia'], df_fundo['Rentabilidade'], label='Rentabilidade do Fundo')
plt.plot(df_ibov['Dia'], df_ibov['Rentabilidade'], label='Rentabilidade do Ibovespa')

# Adicionar rótulos e título
plt.xlabel('Data')
plt.ylabel('Rentabilidade (%)')
plt.title('Comparação entre Fundo e Ibovespa')
plt.yticks([0, 1, 2, 3], ['0%', '100%', '200%', '300%'])
plt.legend()

# Salvar a imagem em um buffer BytesIO
buffer_comparison = BytesIO()
plt.savefig(buffer_comparison, format='png')
plt.close()

# Adicionar imagem ao PDF
buffer_comparison.seek(0)
img_comparison = Image(buffer_comparison)
elements.append(img_comparison)

# Converter a coluna 'Dia' para o tipo datetime
df_fundo['Dia'] = pd.to_datetime(df_fundo['Dia'], format='%d/%m/%Y')

# Ordenar o DataFrame pela data
df_fundo = df_fundo.sort_values(by='Dia', ascending=False)

# Calcular os retornos percentuais diários
df_fundo['Retorno Diário'] = df_fundo['Quota'].pct_change() * 100

# Calcular o desvio padrão dos retornos diários
desvio_padrao = df_fundo['Retorno Diário'].std()

# Plotar o gráfico de linha da volatilidade
plt.figure(figsize=(7, 4))
plt.plot(df_fundo['Dia'], df_fundo['Retorno Diário'], color='blue', marker='', linestyle='-')
plt.title('Volatilidade da Quota')
plt.xlabel('Data')
plt.ylabel('Volatilidade Diária (%)')
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()

# Salvar o gráfico de volatilidade em um segundo buffer BytesIO
buffer_volatility = BytesIO()
plt.savefig(buffer_volatility, format='png')
plt.close()

# Adicionar gráfico de volatilidade ao PDF
buffer_volatility.seek(0)
img_volatility = Image(buffer_volatility)
elements.append(img_volatility)

# Adicionar um Spacer para alinhar verticalmente os gráficos
elements.append(Spacer(1, inch))

# Adicionar DataFrame ao PDF
data = [df_rentabilidade.columns] + df_rentabilidade.values.tolist()

# Definir estilos para a tabela
style = [('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
         ('FONTSIZE', (0, 0), (-1, -1), 8)]
table = Table(data, style=style)

# Adicionar título "Rentabilidade histórica"
elements.append(Spacer(1, 12))  # Espaço de 12 pontos
elements.append(Paragraph('<para align=center><b>Rentabilidade histórica do fundo' + nome + '</b></para>', styles["Normal"]))
elements.append(Spacer(1, 12))  # Espaço de 12 pontos
elements.append(table)

# Renomear as colunas para evitar conflitos após o merge
df_fundo = df_fundo.rename(columns={'Rentabilidade': 'Quota_fundo'})
df_ibov = df_ibov.rename(columns={'Rentabilidade': 'Quota_Ibov'})

# Combinar os DataFrames com base na coluna "Dia"
merged_df = pd.merge(df_fundo, df_ibov, on='Dia')

# Calcular a diferença entre as quotas
merged_df['Diferenca'] = merged_df['Quota_fundo'] - merged_df['Quota_Ibov']

# Calcular a quantidade de dias em que o Itaú ficou abaixo ou acima do Ibovespa
dias_abaixo = (merged_df['Diferenca'] < 0).sum()
dias_acima = (merged_df['Diferenca'] > 0).sum()

# Criar uma tabela com os resultados
tabela = pd.DataFrame({
    'Fundo abaixo do Ibovespa': [dias_abaixo],
    'Fundo acima do Ibovespa': [dias_acima]
})

# Adicionar DataFrame ao PDF
elements.append(Spacer(1, 12))  # Espaço de 12 pontos
elements.append(Paragraph('<para align=center><b>Rentabilidade histórica do Ibovespa</b></para>', styles["Normal"]))
elements.append(Spacer(1, 12))  # Espaço de 12 pontos
data_ib = [df_rentabilidade_ib.columns] + df_rentabilidade_ib.values.tolist()
table_ib = Table(data_ib, style=style)
elements.append(table_ib)

# Calcular a matriz de correlação
correlation_matrix = df_fundo.corr()

# Plotar o gráfico de matriz de correlação
plt.figure(figsize=(7, 7))
sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', fmt=".2f", square=True)
plt.title('Matriz de Correlação')

# Salvar a imagem em um buffer BytesIO
buffer_corr = BytesIO()
plt.savefig(buffer_corr, format='png')
plt.close()

# Adicionar imagem ao PDF
buffer_corr.seek(0)
img_corr = Image(buffer_corr)
elements.append(img_corr)

# Adicionar tabela de dias positivos e negativos
elements.append(Spacer(1, 12))  # Espaço de 12 pontos
elements.append(Paragraph('<para align=center><b>Tabela de Dias Positivos e Negativos</b></para>', styles["Normal"]))
elements.append(Spacer(1, 6))  # Espaço de 6 pontos
data = [tabela.columns] + tabela.values.tolist()
table = Table(data, style=style)
elements.append(table)

# Construir o PDF
pdf.build(elements)

print("PDF criado com sucesso.")





















