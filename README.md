# 📊 Análise de Carteira de Crédito -  2026

Um script em Python focado em inteligência de dados e automação de relatórios para análise de performance de carteiras de crédito. O sistema processa dados pivotados do Excel, limpa inconsistências de formatação numérica e calcula **20 métricas financeiras e estatísticas avançadas**, gerando relatórios corporativos em PDF e uma base consolidada em Excel.

## 📝 Sobre o Projeto

Este projeto foi desenvolvido para automatizar a extração de *insights* de uma carteira de crédito ao longo do tempo. Ele lê uma visão histórica (meses como colunas), realiza o tratamento inteligente dos dados e entrega um diagnóstico detalhado para cada linha de crédito ou filial.

### 🛠️ Atualização Recente (23/03/2026)
* **Correção de Parsing Numérico (Problema da Vírgula):** Implementada uma função de limpeza inteligente (`limpa_numero`) que identifica strings com padrão brasileiro (ex: `1.500,75`), remove os pontos de milhar e converte as vírgulas para pontos decimais. Isso garante que o `pandas` converta os valores para `float` corretamente, sem perder casas decimais ou gerar valores nulos.

## 🚀 Principais Funcionalidades

* **Limpeza Automática de Dados:** Tratamento de valores nulos, remoção de linhas de cabeçalho sujas (Visão PA e Consolidado) e conversão segura de *strings* financeiras.
* **Cálculo de 20 Métricas Complexas:**
  * *Crescimento & Share:* Crescimento Nominal e Percentual, CAGR (Anualizado), Market Share (Inicial/Final e Variação).
  * *Risco & Concorrência:* Volatilidade, Índice HHI (Concentração de Mercado), Elasticidade.
  * *Estatística Descritiva:* Sazonalidade, Tendência (Regressão Linear), Amplitude, Coeficiente de Variação (CV).
  * *Resultado:* Contribuição para a rede e Índice de Performance Ajustado ao Risco.
* **Geração de PDF Visual e Corporativo (via FPDF):** * Layout moderno com paleta de cores definida.
  * *Cards* de métricas com formatação condicional (verde para positivo, vermelho para negativo).
  * Tabelas zebradas para leitura amigável.
  * Caixa de diagnóstico dinâmico gerada via texto.
* **Exportação para Excel:** Gera uma base de dados limpa com as principais métricas calculadas para uso em Dashboards (Power BI, Tableau, etc).

## 🗂️ Estrutura de Arquivos

* **Entrada:**
  * `CRED.xlsx`: Arquivo base contendo os dados pivotados (Linhas de crédito/Filiais nas linhas e Meses nas colunas).
* **Saídas Geradas Automaticamente:**
  * `Relatorio_Analise_Carteira_CRED2026.pdf`: Relatório visual detalhado com resumo executivo e páginas individuais por linha de crédito.
  * `Relatorio_20_Metricas_Carteira_CRED2026.xlsx`: Base tratada e consolidada.

## ⚙️ Tecnologias e Dependências

Este projeto foi construído com Python 3. Certifique-se de ter as seguintes bibliotecas instaladas:

```bash
pip install pandas numpy fpdf openpyxl
