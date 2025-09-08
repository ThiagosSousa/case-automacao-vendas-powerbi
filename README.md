# 🧾 Automação de Vendas com Python + Dashboard em Power BI

Este projeto simula um fluxo real de automação de relatórios de vendas para pequenas empresas (PMEs), unindo Python para geração e consolidação de dados, e Power BI para visualização e análise.

---

## 📌 Objetivos

- Automatizar relatórios mensais de vendas
- Consolidar dados de múltiplos arquivos em um único relatório com análises agrupadas
- Criar dashboard profissional e dinâmico no Power BI
- Eliminar tarefas manuais repetitivas com geração e formatação de arquivos Excel via Python

---

## 🔧 Tecnologias usadas

- Python 3.x
- Pandas
- OpenPyXL
- Power BI Desktop

---

## 📁 Estrutura do projeto

```
.
├── scriptAutomacao.py       # Código Python para automação
├── dados_brutos/            # Arquivos de entrada (.xlsx mensais)
├── saida/                   # Relatório final consolidado
├── powerbi/                 # Dashboard .pbix
├── assets/                  # Imagens e prints
├── docs/                    # Documentação adicional
├── requirements.txt
└── README.md
```

---

## 🧪 Funcionalidades do script

- Gera arquivos de vendas mensais simulados (dados mockados)
- Valida colunas obrigatórias em cada arquivo
- Consolida todos os arquivos válidos em um único DataFrame
- Gera relatórios agrupados por produto, vendedor, mês, forma de pagamento e ticket médio
- Exporta todos os resumos para um único Excel com múltiplas abas
- Formata os valores com estilo monetário (R$)
- Adiciona data/hora de geração
- Cria cópia com nome fixo para integração com Power BI

---

## 📊 Dashboard (Power BI)

Inclui:

- Total de vendas
- Número total de transações
- Ticket médio
- Mês com maior volume de vendas
- Gráficos por produto, vendedor, mês, forma de pagamento
- Segmentações por período, vendedor, produto e forma de pagamento
- Ordenação cronológica dos meses com tabela auxiliar

![Dashboard Vendas](assets/dashboard.png)

---

## 🚀 Como executar

1. Clone este repositório
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Adicione seus arquivos `.xlsx` na pasta `dados_brutos/`
4. Execute `scriptAutomacao.py`
5. Abra o arquivo gerado `relatorioFinal.xlsx` com o modelo `dashboard_vendas.pbix` no Power BI

---

## 👤 Autor

Thiago Sousa  
[LinkedIn](https://www.linkedin.com/in/thiagosiqueirasousa/)

---

## 🪪 Licença

Este projeto está licenciado sob a licença MIT. Consulte o arquivo LICENSE para mais detalhes.
