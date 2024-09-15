import pandas as pd
import matplotlib.pyplot as plt

# Carregar a planilha Excel
file_path = 'C:/Users/Gabrielly Alves/Documents/rd_locacoes_dados.xlsx'
data = pd.read_excel(file_path)


# Agrupar os dados por tipo de equipamento e calcular o total de vendas e faturamento
grouped_data = data.groupby('Tipo de Equipamento').agg(
    Total_Vendas=('Valor Total', 'size'),
    Faturamento_Total=('Valor Total', 'sum')
).reset_index()

# Exibir o resumo dos dados agrupados
print(grouped_data)

# Criar gráfico de barras para total de vendas e faturamento
fig1, ax1 = plt.subplots(figsize=(12, 6))

# Gráfico de barras para total de vendas
ax1.bar(grouped_data['Tipo de Equipamento'], grouped_data['Total_Vendas'], color='b', alpha=0.6, label='Total de Vendas')
ax1.set_xlabel('Tipo de Equipamento')
ax1.set_ylabel('Total de Vendas', color='b')

# Gráfico de linha para faturamento total
ax2 = ax1.twinx()
ax2.plot(grouped_data['Tipo de Equipamento'], grouped_data['Faturamento_Total'], color='r', marker='o', label='Faturamento Total')
ax2.set_ylabel('Faturamento Total (R$)', color='r')

# Título e legendas
plt.title('Total de Vendas e Faturamento por Tipo de Equipamento')
ax1.legend(loc='upper left')
ax2.legend(loc='upper right')

# Rotacionar rótulos do eixo x para melhor visualização
plt.xticks(rotation=45)

# Mostrar o gráfico
plt.tight_layout()
plt.show()

# Criar gráfico de pizza para faturamento total por tipo de equipamento
fig2, ax = plt.subplots(figsize=(8, 8))
ax.pie(grouped_data['Faturamento_Total'], labels=grouped_data['Tipo de Equipamento'], autopct='%1.1f%%', startangle=140)
ax.set_title('Faturamento Total por Tipo de Equipamento')

# Mostrar o gráfico
plt.tight_layout()
plt.show()
