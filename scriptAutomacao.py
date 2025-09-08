"""
Automação de Relatório de Vendas
--------------------------------
Este script lê todos os arquivos .xlsx, consolida os dados de vendas, gera relatórios por categoria, produtos e vendedor e exporta os resultados para um novo arquivo Excel.
"""

import os
import shutil
import pandas as pd
from glob import glob
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.styles import Font

# 1. Configurações
pastaEntrada = "dados_brutos"
horarioDeSalvamento = datetime.now().strftime("%d-%m-%Y-%H-%M")
arquivoSaida = fr'saida\relatorioFinal_{horarioDeSalvamento}.xlsx'

# 2. Função: Carregar arquivos Excel da pasta dados_brutos
def carregarArquivos(pasta):
    #Carrega todos arquivos do formato ".xlsx"
    caminhos = glob(os.path.join(pasta, "*.xlsx"))
    print(f"📂 {len(caminhos)} arquivos encontrados na pasta '{pasta}'.")
    
    dfList = []
    colunasEsperadas = ["Data", "Produto", "Categoria", "Valor", "Vendedor", "Forma de Pagamento"]

    for caminho in caminhos:
        #Carregando o 
        try:
            df = pd.read_excel(caminho, sheet_name="Vendas")
        except Exception as e:
            print(f"⚠ Erro ao ler '{caminho}': {e}")
            continue
        
        #Validando as colunas
        colunasArquivo = df.columns.tolist()
        if not all(col in colunasArquivo for col in colunasEsperadas):
            print(f"❌ Arquivo '{os.path.basename(caminho)}' ignorado: colunas ausentes.")
            print(f"    -> Esperadas: {colunasEsperadas}")
            print(f"    -> Encontradas: {colunasArquivo}")
            continue
        print(f"✔ Arquivo '{caminho}' carregado com sucesso.")

        #Criando a coluna com o nome do arquivo de origem
        df["Arquivo Origem"] = os.path.basename(caminho) #rastreabilidade

        #Criando a coluna com o mês que ocorreu a transação
        df["Mês"] = pd.to_datetime(df["Data"]).dt.month_name(locale="pt_BR")

        #Adicionando o dataframe na lista
        dfList.append(df)

    #Caso nenhum arquivo válido seja carregado
    if not dfList:
        print("Nenhum arquivo válido foi carregado")
        return pd.DataFrame() #Dataframe vazio para evitar quebra   
    
    return pd.concat(dfList, ignore_index=True)

# 3. Função: Gerar relatórios resumidos
def gerarRelatorios(df):
    resumoProduto = df.groupby("Produto")["Valor"].sum().reset_index()
    resumoVendedor = df.groupby("Vendedor")["Valor"].sum().reset_index()
    resumoPagamento = df.groupby("Forma de Pagamento")["Valor"].sum().reset_index()
    resumoMes = df.groupby("Mês", sort=False)["Valor"].sum().reset_index()
    resumoCategoriaProduto = df.groupby(["Categoria", "Produto"])["Valor"].sum().reset_index()
    ticketMedioVendedor = df.groupby("Vendedor").agg(
        Total_Vendido=("Valor", "sum"),
        Quantidade_Vendas=("Valor", "count"),
        Ticket_Médio=("Valor", "mean")    
    ).reset_index()

    #Ordenando por ordem decrescente de Valor
    resumoVendedor = resumoVendedor.sort_values(by="Valor", ascending=False)
    resumoProduto = resumoProduto.sort_values(by="Valor", ascending=False)
    resumoPagamento = resumoPagamento.sort_values(by="Valor", ascending=False)
    #resumoMes = resumoMes.sort_values(by="Valor", ascending=False) #Mantendo o resumo do mês por ordem cronológica.
    resumoCategoriaProduto = resumoCategoriaProduto.sort_values(by="Valor", ascending=False)
    ticketMedioVendedor = ticketMedioVendedor.sort_values(by="Ticket_Médio", ascending=False)
    
    return resumoProduto, resumoVendedor, resumoPagamento, resumoMes, resumoCategoriaProduto, ticketMedioVendedor

# 4. Função: Exportar para o Excel com Múltiplas abas
def exportarRelatorio(dfOriginal, prod, vend, pagamento, mes, catProd, ticketMedio, caminhoSaida):
    with pd.ExcelWriter(caminhoSaida, engine="openpyxl") as writer:
        dfOriginal.to_excel(writer, index=False, sheet_name="Vendas Consolidadas")
        prod.to_excel(writer, index=False, sheet_name="Resumo Produto")
        vend.to_excel(writer, index=False, sheet_name="Resumo Vendedor")
        pagamento.to_excel(writer, index=False, sheet_name="Resumo Pagamento")
        mes.to_excel(writer, index=False, sheet_name="Resumo Mês")
        catProd.to_excel(writer, index=False, sheet_name="Resumo Categoria x Produto")
        ticketMedio.to_excel(writer, index=False, sheet_name="Ticket Médio por Vendedor")
    print(f"Relatório salvo em: {caminhoSaida}")

    wb = load_workbook(caminhoSaida)
    abasComValores = [
        "Resumo Produto",
        "Resumo Vendedor",
        "Resumo Pagamento",
        "Resumo Mês",
        "Resumo Categoria x Produto",
        "Ticket Médio por Vendedor"
    ]

    for abaNome in abasComValores:
        aba = wb[abaNome]
        if abaNome == "Resumo Categoria x Produto":
            for cell in aba["C"][1:]:
                cell.number_format = '"R$" #,##0.00'
        elif abaNome == "Ticket Médio por Vendedor":
            for cell in aba["B"][1:] + aba["D"][1:]:
                cell.number_format = '"R$" #,##0.00'    
        else:
            for cell in aba["B"][1:]:
                cell.number_format = '"R$" #,##0.00'
    
    #Inserindo data/hora que foi gerado o relatório
    abaConsolidada = wb["Vendas Consolidadas"]
    agora = datetime.now().strftime("Relatório gerado em: %d/%m/%Y - %H:%M")
    abaConsolidada[f"A{(len(dfOriginal)+3)}"] = agora
    abaConsolidada[f"A{(len(dfOriginal)+3)}" ].font = Font(italic=True, size=12, bold=True)

    wb.save(caminhoSaida)

def copiaArquivo():
    caminhoFixo = r"saida\relatorioDashboard.xlsx"
    shutil.copy(arquivoSaida, caminhoFixo)
    print(f"📤 Cópia salva em: '{caminhoFixo}' para uso no Power BI.")


# 5. Função Principal
def main():
    df = carregarArquivos(pastaEntrada)
    print(f"{len(df)} registros carregados.")
    resumoProduto, resumoVendedor, resumoPagamento, resumoMes, resumoCatProd, ticketMedioVendedor = gerarRelatorios(df)
    exportarRelatorio(df,resumoProduto, resumoVendedor, resumoPagamento, resumoMes, resumoCatProd, ticketMedioVendedor, arquivoSaida)
    copiaArquivo()
    #Criando cópia 
    

#Execução
if __name__ == "__main__":
    main()