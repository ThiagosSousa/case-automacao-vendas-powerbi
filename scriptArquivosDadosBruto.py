import pandas as pd
import random
from datetime import datetime
from openpyxl import Workbook
import os


#Configurações de dados
produtos = ["Chá Verde", "Granola", "Suco Detox", "Barra de Proteína", "Água de Coco", "Mel Orgânico", "Shampoo Natural", "Whey Vegano", "Pasta de Amendoim", "Sabonete Vegano"]

categorias = {
    "Chá Verde": "Bebidas",
    "Granola": "Alimentos",
    "Suco Detox": "Bebidas",
    "Barra de Proteína": "Suplementos",
    "Água de Coco": "Bebidas",
    "Mel Orgânico": "Alimentos",
    "Shampoo Natural": "Higiene",
    "Whey Vegano": "Suplementos",
    "Pasta de Amendoim": "Alimentos",
    "Sabonete Vegano": "Hiegiene" 
}

vendedores = ["Ana", "Carlos", "Juliana", "Bruno"]

formasPagamento = ["Dinheiro", "Crédito", "Débito", "PIX"]

#Função para gerar as vendas aleatórias de um mês
def gerarVendasMes(mes: int, ano: int, qtdLinhas: int = random.randint(50,100)):
    dados = []
    for _ in range(qtdLinhas):
        produto = random.choice(produtos)
        if mes != 2:
            dados.append({
                "Data": datetime(ano, mes, random.randint(1,30)),
                "Produto": produto,
                "Categoria": categorias[produto],
                "Valor": round(random.uniform(10, 50), 2),
                "Vendedor": random.choice(vendedores),
                "Forma de Pagamento": random.choice(formasPagamento)
            })
        else:
            dados.append({
                "Data": datetime(ano, mes, random.randint(1,28)),
                "Produto": produto,
                "Categoria": categorias[produto],
                "Valor": round(random.uniform(10, 50), 2),
                "Vendedor": random.choice(vendedores),
                "Forma de Pagamento": random.choice(formasPagamento)
            })
    return pd.DataFrame(dados)


#Pasta de Saída
pastaSaida = r"G:\BACKUP PARA FORMATAÇÃO\Thiago\Projeto Python\automacao_vendas\dados_brutos"
os.makedirs(pastaSaida, exist_ok=True)

#Gerando os arquivos para 6 meses
arquivosGerados = []
for mes in [1, 2, 3, 4, 5, 6]:
    df = gerarVendasMes(mes=mes, ano=2025, qtdLinhas=random.randint(50,100))
    nomeArquivo = f"{pastaSaida}/vendas_{mes:02d}_2025.xlsx"
    df.to_excel(nomeArquivo, index=False, sheet_name="Vendas")
    arquivosGerados.append(nomeArquivo)

arquivosGerados