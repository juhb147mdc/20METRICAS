#Script corrigido em questao da virgula e podendo ser usado 23/03/2026

import pandas as pd
import numpy as np
import warnings
from fpdf import FPDF
import os

warnings.filterwarnings('ignore')

# ============================================================================
# CONFIGURAÇÕES GERAIS
# ============================================================================
ARQUIVO_ENTRADA = 'CRED.xlsx'
ARQUIVO_SAIDA_EXCEL = 'Relatorio_20_Metricas_Carteira_CRED2026.xlsx'
ARQUIVO_SAIDA_PDF = 'Relatorio_Analise_Carteira_CRED2026.pdf'

# ============================================================================
# 1. CARGA E TRANSFORMAÇÃO DE DADOS
# ============================================================================

print("\n" + "=" * 80)
print("INICIANDO PROCESSAMENTO - 20 MÉTRICAS COM PDF VISUAL")
print("=" * 80)

if not os.path.exists(ARQUIVO_ENTRADA):
    print(f"[ERRO] Arquivo '{ARQUIVO_ENTRADA}' não encontrado na pasta.")
    exit()

print("1. Lendo e limpando arquivo pivotado...")
try:
    # Lemos o arquivo
    df = pd.read_excel(ARQUIVO_ENTRADA) 
    
    # Renomeia a primeira coluna
    if 'Linha_Credito' not in df.columns:
        primeiro_col = df.columns[0]
        df.rename(columns={primeiro_col: 'Linha_Credito'}, inplace=True)
    
    # 1. Remove a linha de "Visao PA e Consolidado" (linha 0 no Pandas)
    if df.iloc[0].astype(str).str.contains('Visao').any():
        df = df.iloc[1:].reset_index(drop=True)
        
    # Identificar colunas de meses
    meses_cols = df.columns[1:].tolist()
    
    # 2. Limpeza INTELIGENTE de números (CORREÇÃO DA VÍRGULA)
    for col in meses_cols:
        def limpa_numero(valor):
            # Só faz o replace se o valor for uma string (texto) E tiver vírgula
            if isinstance(valor, str) and ',' in valor:
                valor = valor.strip()
                valor = valor.replace('.', '')  # Remove ponto de milhar
                valor = valor.replace(',', '.') # Troca vírgula decimal por ponto
            return valor

        # Aplica a função de limpeza
        df[col] = df[col].apply(limpa_numero)
        # Converte para número (floats que já estavam corretos passam direto sem perder a casa decimal)
        df[col] = pd.to_numeric(df[col], errors='coerce')
        
    # Preenche possíveis valores vazios (NaN) com 0
    df[meses_cols] = df[meses_cols].fillna(0)

    print(f"   -> Arquivo lido e dados convertidos com sucesso!")
    
except Exception as e:
    print(f"[ERRO CRÍTICO NA LEITURA]: {e}")
    print("Verifique se seu Excel está no formato pivotado.")
    exit()

primeiro_mes = meses_cols[0]
ultimo_mes = meses_cols[-1]

print(f"   Período identificado: {primeiro_mes} a {ultimo_mes}")
print(f"   Total de Linhas válidas: {len(df)}")


# ============================================================================
# 2. CÁLCULO DAS 20 MÉTRICAS
# ============================================================================
print("2. Calculando métricas complexas...")

try:
    # M1 & M2: Crescimento
    df['M1_Crescimento_Nominal'] = df[ultimo_mes] - df[primeiro_mes]
    df['M2_Crescimento_%'] = ((df[ultimo_mes] - df[primeiro_mes]) / df[primeiro_mes].replace(0, np.nan) * 100)

    # M3: CAGR
    n_anos = (len(meses_cols) - 1) / 12
    if n_anos == 0: n_anos = 1 
    df['M3_CAGR'] = ((df[ultimo_mes] / df[primeiro_mes].replace(0, np.nan)) ** (1/n_anos) - 1) * 100

    # M4 & M5: Market Share
    total_primeiro = df[primeiro_mes].sum()
    total_ultimo = df[ultimo_mes].sum()
    df['M4_Market_Share_Inicial'] = (df[primeiro_mes] / total_primeiro * 100)
    df['M5_Market_Share_Final'] = (df[ultimo_mes] / total_ultimo * 100)

    # M6: Variação Share
    df['M6_Variacao_Market_Share'] = df['M5_Market_Share_Final'] - df['M4_Market_Share_Inicial']

    # M7: Volatilidade
    colunas_taxa = []
    for i in range(len(meses_cols) - 1):
        c_atual, c_prox = meses_cols[i], meses_cols[i+1]
        nome_col = f'Taxa_{i}'
        df[nome_col] = ((df[c_prox] - df[c_atual]) / df[c_atual].replace(0, np.nan) * 100)
        colunas_taxa.append(nome_col)

    df['M7_Volatilidade'] = df[colunas_taxa].std(axis=1, ddof=1)

    # M8 & M9: HHI e Diversificação
    hhi_inicial = ((df['M4_Market_Share_Inicial']/100)**2).sum() * 10000
    hhi_final = ((df['M5_Market_Share_Final']/100)**2).sum() * 10000
    indice_div_fin = 10000 / hhi_final if hhi_final > 0 else 0

    # M10: Elasticidade
    crescimento_carteira = ((total_ultimo - total_primeiro) / total_primeiro * 100)
    df['M10_Elasticidade'] = (df['M2_Crescimento_%'] / crescimento_carteira * df['M4_Market_Share_Inicial'])

    # M11: Sazonalidade
    df['M11_Media_Periodo'] = df[meses_cols].mean(axis=1)
    df['M11_Indice_Sazonalidade'] = (df[ultimo_mes] / df['M11_Media_Periodo'] * 100)

    # M12: Tendência (Regressão)
    meses_num = np.arange(len(meses_cols))
    for idx, row in df.iterrows():
        vals = [row[m] for m in meses_cols]
        vals_validos = [v for v in vals if pd.notnull(v)]
        
        if len(vals_validos) > 1:
            try:
                coeffs = np.polyfit(range(len(vals_validos)), vals_validos, 1)
                df.loc[idx, 'M12_Inclinacao'] = coeffs[0]
            except:
                df.loc[idx, 'M12_Inclinacao'] = 0
        else:
            df.loc[idx, 'M12_Inclinacao'] = 0

    # M13 a M18: Estatísticas Descritivas
    df['M13_Acumulado'] = ((df[ultimo_mes] / df[primeiro_mes].replace(0, np.nan) - 1) * 100)
    df['M14_Taxa_Media'] = df[colunas_taxa].mean(axis=1)
    df['M15_Min'] = df[meses_cols].min(axis=1)
    df['M16_Max'] = df[meses_cols].max(axis=1)
    df['M17_Amplitude'] = df['M16_Max'] - df['M15_Min']
    df['M18_CV'] = (df['M7_Volatilidade'] / df['M11_Media_Periodo'] * 100)

    # M19 & M20
    cresc_nom_total = total_ultimo - total_primeiro
    if cresc_nom_total != 0:
        df['M19_Contribuicao'] = (df['M1_Crescimento_Nominal'] / cresc_nom_total * 100)
    else:
        df['M19_Contribuicao'] = 0
        
    df['M20_Performance'] = df['M2_Crescimento_%'] / df['M7_Volatilidade'].replace(0, np.nan)

    print("   Cálculos concluídos com sucesso.")

except Exception as e:
    print(f"\n[ERRO NOS CÁLCULOS]: {e}")
    exit()

# ============================================================================
# CLASSE PDF REDESENHADA (Visual Moderno e Corporativo)
# ============================================================================
class PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
        # Paleta de Cores Corporativa
        self.primary_col = (41, 128, 185)    # Azul Belieze (Títulos)
        self.secondary_col = (52, 73, 94)    # Azul Escuro (Textos Fortes)
        self.accent_col = (236, 240, 241)    # Cinza Claro (Fundos)
        self.text_col = (44, 62, 80)         # Cinza Escuro (Texto Geral)
        self.positive_col = (39, 174, 96)    # Verde
        self.negative_col = (192, 57, 43)    # Vermelho

    def header(self):
        # Faixa superior colorida
        self.set_fill_color(*self.primary_col)
        self.rect(0, 0, 210, 20, 'F')
        
        self.set_y(5)
        self.set_font('Arial', 'B', 16)
        self.set_text_color(255, 255, 255)
        self.cell(0, 10, 'Relatório de Performance - Carteira de Crédito', 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Página {self.page_no()} | Gerado via Python Analytics', 0, 0, 'C')

    def section_title(self, title):
        """Cria um título de seção bonito com uma linha abaixo"""
        self.set_font('Arial', 'B', 14)
        self.set_text_color(*self.primary_col)
        self.cell(0, 8, title, 0, 1, 'L')
        # Linha fina abaixo do título
        self.set_draw_color(*self.primary_col)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(5)

    def card_metric(self, x, y, w, title, value, is_currency=False, subtitle=""):
        """Desenha um 'Card' de métrica (Caixa com título e número grande)"""
        h = 25
        self.set_xy(x, y)
        
        # Fundo do card
        self.set_fill_color(245, 247, 250)
        self.set_draw_color(200, 200, 200)
        self.rect(x, y, w, h, 'DF')
        
        # Título da métrica
        self.set_xy(x, y + 2)
        self.set_font('Arial', '', 9)
        self.set_text_color(100, 100, 100)
        self.cell(w, 5, title, 0, 1, 'C')
        
        # Valor principal
        self.set_xy(x, y + 8)
        self.set_font('Arial', 'B', 14)
        
        # Lógica de cor para números (Verde se positivo, Vermelho se negativo)
        val_str = str(value)
        if "-" in val_str and "R$" not in val_str: 
             self.set_text_color(*self.negative_col)
        elif "%" in val_str and "-" not in val_str:
             self.set_text_color(*self.positive_col)
        else:
             self.set_text_color(*self.secondary_col)
             
        self.cell(w, 8, val_str, 0, 1, 'C')

        # Subtítulo (opcional)
        if subtitle:
            self.set_xy(x, y + 17)
            self.set_font('Arial', '', 7)
            self.set_text_color(150, 150, 150)
            self.cell(w, 4, subtitle, 0, 1, 'C')

    def striped_table(self, data_dict):
        """Cria uma tabela zebrada (linhas alternadas)"""
        self.set_font('Arial', '', 10)
        line_height = 8
        fill = False
        
        # Título das colunas
        self.set_fill_color(*self.secondary_col)
        self.set_text_color(255, 255, 255)
        self.set_font('Arial', 'B', 10)
        self.cell(100, line_height, "Indicador", 0, 0, 'L', True)
        self.cell(90, line_height, "Resultado", 0, 1, 'R', True)

        # Dados
        self.set_fill_color(240, 240, 240)
        self.set_text_color(*self.text_col)
        
        for k, v in data_dict.items():
            self.set_font('Arial', '', 10)
            self.cell(100, line_height, f"  {str(k)}", 0, 0, 'L', fill)
            
            # Negrito para valores
            self.set_font('Arial', 'B', 10)
            
            # Cor condicional simples
            if "-" in str(v) and "R$" not in str(v):
                self.set_text_color(*self.negative_col)
            elif "%" in str(v) and "-" not in str(v) and "0.00" not in str(v):
                self.set_text_color(*self.positive_col)
            else:
                self.set_text_color(*self.text_col)
                
            self.cell(90, line_height, str(v), 0, 1, 'R', fill)
            
            # Reset e alternar cor
            self.set_text_color(*self.text_col)
            fill = not fill
            
        self.ln(5)

    def diagnosis_box(self, text):
        """Caixa de texto destacada para o diagnóstico"""
        self.ln(5)
        self.set_fill_color(255, 250, 240)
        self.set_draw_color(243, 156, 18)  
        
        # Salva posição Y
        y_start = self.get_y()
        
        self.set_font('Arial', 'B', 11)
        self.set_text_color(211, 84, 0)
        self.cell(0, 8, "  Diagnóstico da Carteira", 0, 1)
        
        self.set_font('Arial', '', 10)
        self.set_text_color(50, 50, 50)
        self.multi_cell(0, 6, text, border=0)
        
        # Desenha retângulo ao redor do que foi escrito
        y_end = self.get_y()
        self.rect(10, y_start, 190, y_end - y_start + 2)
        self.ln(5)

# ============================================================================
# 3. GERAÇÃO DO RELATÓRIO PDF (Melhorado)
# ============================================================================
print("3. Gerando PDF Analítico (Visual Profissional)...")
pdf = PDF()

# Função auxiliar de formatação
def fmt(val, tipo='num'):
    if pd.isna(val) or val == np.inf or val == -np.inf: return "N/A"
    if tipo == 'moeda': return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if tipo == 'pct': return f"{val:.2f}%".replace(".", ",")
    return f"{val:.2f}".replace(".", ",")

# --- PÁGINA 1: RESUMO EXECUTIVO (DASHBOARD) ---
pdf.add_page()
pdf.section_title("Resumo Executivo da Carteira")

# Cards Principais (Layout Grid)
y_cards = pdf.get_y()
pdf.card_metric(10, y_cards, 45, "Volume Inicial", fmt(total_primeiro, 'moeda'))
pdf.card_metric(58, y_cards, 45, "Volume Final", fmt(total_ultimo, 'moeda'))
pdf.card_metric(106, y_cards, 45, "Crescimento (%)", fmt(crescimento_carteira, 'pct'))
pdf.card_metric(154, y_cards, 45, "Novo Capital", fmt(cresc_nom_total, 'moeda'))

pdf.ln(35) # Espaço após os cards

# Seção de Risco e Destaques
pdf.section_title("Indicadores de Concentração e Destaques")

interp_hhi = "Alta Concentração" if hhi_final > 2500 else "Moderada" if hhi_final > 1500 else "Competitiva"
linha_top_cresc = df.loc[df['M2_Crescimento_%'].idxmax(), 'Linha_Credito']
linha_top_share = df.loc[df['M5_Market_Share_Final'].idxmax(), 'Linha_Credito']

# Tabela Resumo Manual
resumo_dados = {
    "Índice HHI (Concentração)": f"{hhi_final:.0f} pts ({interp_hhi})",
    "Nº de Linhas Equivalentes": f"{indice_div_fin:.2f}",
    "Maior Crescimento (%)": f"{linha_top_cresc} ({fmt(df['M2_Crescimento_%'].max(), 'pct')})",
    "Maior Market Share": f"{linha_top_share} ({fmt(df['M5_Market_Share_Final'].max(), 'pct')})",
    "Melhor Performance Ajustada": f"{df.loc[df['M20_Performance'].idxmax(), 'Linha_Credito']}"
}
pdf.striped_table(resumo_dados)

# Texto descritivo
texto_resumo = (
    f"A carteira apresentou uma variação total de {fmt(crescimento_carteira, 'pct')} no período de {primeiro_mes} a {ultimo_mes}. "
    f"O volume financeiro movimentou {fmt(cresc_nom_total, 'moeda')}. "
    f"O nível de concorrência interna medido pelo HHI indica uma estrutura {interp_hhi.lower()}."
)
pdf.diagnosis_box(texto_resumo)

# --- PÁGINAS INDIVIDUAIS ---
for idx, row in df.iterrows():
    nome_linha = row['Linha_Credito']
    
    pdf.add_page()
    
    # Cabeçalho da página com nome da filial grande
    pdf.set_font('Arial', 'B', 16)
    pdf.set_text_color(*pdf.secondary_col)
    pdf.cell(0, 10, f"Análise: {nome_linha}", 0, 1)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)
    
    # 3 Cards Superiores
    y_pos = pdf.get_y()
    pdf.card_metric(10, y_pos, 60, "Volume Atual", fmt(row[ultimo_mes], 'moeda'))
    pdf.card_metric(75, y_pos, 60, "Crescimento", fmt(row['M2_Crescimento_%'], 'pct'))
    pdf.card_metric(140, y_pos, 60, "Market Share", fmt(row['M5_Market_Share_Final'], 'pct'))
    
    pdf.ln(35)
    
    # Tabela de Métricas Detalhadas
    pdf.section_title("Detalhamento Técnico")
    
    dados_linha = {
        'Crescimento Nominal': fmt(row['M1_Crescimento_Nominal'], 'moeda'),
        'CAGR (Anualizado)': fmt(row['M3_CAGR'], 'pct'),
        'Variação de Share': f"{row['M6_Variacao_Market_Share']:+.2f} pp",
        'Volatilidade (Risco)': fmt(row['M7_Volatilidade'], 'pct'),
        'Índice de Performance': fmt(row['M20_Performance'], 'num'),
        'Tendência Linear (R$)': f"{row['M12_Inclinacao']:,.2f} /mês",
        'Pico (Valor Máximo)': fmt(row['M16_Max'], 'moeda'),
        'Sazonalidade': fmt(row['M11_Indice_Sazonalidade'], 'pct')
    }
    
    # Usa a nova função de tabela bonita
    pdf.striped_table(dados_linha)
    
    # Diagnóstico
    status = "crescimento" if row['M2_Crescimento_%'] > 0 else "retração"
    risco = "Alto" if row['M7_Volatilidade'] > 5 else "Baixo"
    
    diagnostico_txt = (
        f"A unidade {nome_linha} registrou {status} de {fmt(row['M2_Crescimento_%'], 'pct')} no período analisado. "
        f"Atualmente representa {fmt(row['M5_Market_Share_Final'], 'pct')} do total da carteira.\n"
        f"O comportamento da linha apresenta volatilidade classificada como {risco} ({fmt(row['M7_Volatilidade'], 'pct')}), "
        f"com uma contribuição direta de {fmt(row['M19_Contribuicao'], 'pct')} para o resultado consolidado da rede."
    )
    
    pdf.diagnosis_box(diagnostico_txt)

try:
    pdf.output(ARQUIVO_SAIDA_PDF)
    print(f"✓ PDF Visual Gerado com sucesso: {ARQUIVO_SAIDA_PDF}")
except Exception as e:
    print(f"[ERRO] PDF não salvo. Feche o arquivo se estiver aberto. Erro: {e}")

# ============================================================================
# 4. EXPORTAÇÃO EXCEL
# ============================================================================
print("4. Gerando Excel Detalhado...")

colunas_finais = [
    'Linha_Credito', primeiro_mes, ultimo_mes,
    'M1_Crescimento_Nominal', 'M2_Crescimento_%', 'M3_CAGR',
    'M4_Market_Share_Inicial', 'M5_Market_Share_Final', 'M6_Variacao_Market_Share',
    'M7_Volatilidade', 'M10_Elasticidade', 'M12_Inclinacao', 'M18_CV', 'M20_Performance'
]

cols_existentes = [c for c in colunas_finais if c in df.columns]
df_export = df[cols_existentes].copy()

df_export.to_excel(ARQUIVO_SAIDA_EXCEL, index=False)
print(f"✓ Excel Gerado com sucesso: {ARQUIVO_SAIDA_EXCEL}")

print("\n" + "=" * 80)
print("PROCESSO CONCLUÍDO")
print("=" * 80)
