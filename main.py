import pandas as pd
from fpdf import FPDF
import os

# Carregar os dados do Excel
tabela_vendas = pd.read_excel("Vendas.xlsx")

# Calcular métricas
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
prod_vendidos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
ticket_medio = (faturamento['Valor Final'] / prod_vendidos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={'Valor Final': 'Ticket Médio'})

# Criar PDF
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Arial", style='B', size=16)
pdf.cell(200, 10, "Relatório de Vendas por Loja", ln=True, align='C')

# Adicionar Faturamento
pdf.set_font("Arial", style='B', size=12)
pdf.cell(200, 10, "Faturamento por Loja", ln=True, align='L')
pdf.set_font("Arial", size=10)
for loja, valores in faturamento.iterrows():
    pdf.cell(200, 8, f"{loja}: R$ {valores['Valor Final']:,.2f}", ln=True)

pdf.ln(5)

# Adicionar Quantidade Vendida
pdf.set_font("Arial", style='B', size=12)
pdf.cell(200, 10, "Quantidade de Produtos Vendidos", ln=True, align='L')
pdf.set_font("Arial", size=10)
for loja, valores in prod_vendidos.iterrows():
    pdf.cell(200, 8, f"{loja}: {valores['Quantidade']} produtos", ln=True)

pdf.ln(5)

# Adicionar Ticket Médio
pdf.set_font("Arial", style='B', size=12)
pdf.cell(200, 10, "Ticket Médio por Loja", ln=True, align='L')
pdf.set_font("Arial", size=10)
for loja, valores in ticket_medio.iterrows():
    pdf.cell(200, 8, f"{loja}: R$ {valores[0]:,.2f}", ln=True)

# Salvar PDF
pdf_path = "Relatorio_Vendas.pdf"
pdf.output(pdf_path)

# Abrir automaticamente o arquivo no Windows
os.startfile(pdf_path)

print(f"O PDF foi gerado e será aberto automaticamente: {pdf_path}")
