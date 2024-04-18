import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
import matplotlib.pyplot as plt

# Criando os dados do estudo de capabilidade (substitua com seus próprios dados)
dados = {
    'Amostra': range(1, 11),
    'Medidas': [9.2, 9.5, 9.1, 9.6, 9.3, 9.4, 9.7, 9.2, 9.5, 9.6]
}

# Calculando a média e o desvio padrão
media = sum(dados['Medidas']) / len(dados['Medidas'])
desvio_padrao = (sum((x - media) ** 2 for x in dados['Medidas']) / len(dados['Medidas'])) ** 0.5

# Criando um DataFrame com os dados
df = pd.DataFrame(dados)

# Adicionando colunas para cálculos
df['Media'] = media
df['Desvio_Padrao'] = desvio_padrao
df['Limite_Superior'] = media + 3 * desvio_padrao
df['Limite_Inferior'] = media - 3 * desvio_padrao
df['Cp'] = (df['Limite_Superior'] - df['Limite_Inferior']) / (6 * desvio_padrao)
df['CpK'] = min(df['Limite_Superior'] - df['Medidas'], df['Medidas'] - df['Limite_Inferior']) / (3 * desvio_padrao)

# Criando um arquivo Excel e escrevendo os dados
wb = Workbook()
ws = wb.active

ws.append(['Amostra', 'Medidas', 'Media', 'Desvio_Padrao', 'Limite_Superior', 'Limite_Inferior', 'Cp', 'CpK'])
for _, row in df.iterrows():
    ws.append(row.tolist())

# Adicionando gráfico
chart = LineChart()
chart.title = "Estudo de Capabilidade"
chart.y_axis.title = 'Medidas'
data = Reference(ws, min_col=2, min_row=2, max_col=2, max_row=len(dados['Medidas']) + 1)
cats = Reference(ws, min_col=1, min_row=2, max_row=len(dados['Amostra']) + 1)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "I2")

# Adicionando descrições
ws['A1'] = "Descrição:"
ws['A2'] = "1. Amostra: Número da amostra"
ws['A3'] = "2. Medidas: Valores medidos"
ws['A4'] = f"3. Média: {media}"
ws['A5'] = f"4. Desvio Padrão: {desvio_padrao}"
ws['A6'] = f"5. Limite Superior: {media + 3 * desvio_padrao}"
ws['A7'] = f"6. Limite Inferior: {media - 3 * desvio_padrao}"
ws['A8'] = "7. Cp: Índice de Capacidade do Processo"
ws['A9'] = "8. CpK: Índice de Capacidade do Processo Ajustado"

# Salvando o arquivo Excel
wb.save("Estudo_de_Capabilidade.xlsx")

# Salvar gráfico como imagem
plt.plot(df['Amostra'], df['Medidas'])
plt.title("Estudo de Capabilidade")
plt.xlabel("Amostra")
plt.ylabel("Medidas")
plt.savefig("grafico_estudo_capabilidade.png")
