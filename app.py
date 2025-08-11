import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
from io import BytesIO

# Caminho fixo para o arquivo DesVend salvo localmente
FALLBACK_DESVEND = os.path.join("data", "DesVend AUDITORIA_AUTOMATICA.xlsx")

st.set_page_config(page_title="Sistema de Premiação", layout="wide")

@st.cache_data
def read_file(uploaded_file=None, fallback_path=None):
    """
    Lê arquivo Excel da aba DesVend e retorna DataFrame.
    """
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name="DesVend")
    elif fallback_path is not None and os.path.exists(fallback_path):
        df = pd.read_excel(fallback_path, sheet_name="DesVend")
    else:
        df = pd.DataFrame()
    return df

def processar_dados(df):
    # Filtrar colunas relevantes
    df = df[['LOJA', 'TOTAL VENDAS']].dropna()
    df = df[df['TOTAL VENDAS'] > 0]
    
    # Agrupar por loja
    df_grouped = df.groupby('LOJA', as_index=False)['TOTAL VENDAS'].sum()
    
    # Calcular percentual
    total = df_grouped['TOTAL VENDAS'].sum()
    df_grouped['%'] = (df_grouped['TOTAL VENDAS'] / total) * 100
    
    # Ordenar
    df_grouped = df_grouped.sort_values('TOTAL VENDAS', ascending=False)
    return df_grouped

def gerar_grafico(df_grouped):
    fig, ax = plt.subplots(figsize=(10,6))
    bars = ax.bar(df_grouped['LOJA'].astype(str), df_grouped['TOTAL VENDAS'], color='royalblue')

    # Adicionar rótulos
    for bar, val, pct in zip(bars, df_grouped['TOTAL VENDAS'], df_grouped['%']):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                f"R$ {val:,.0f}\n({pct:.1f}%)",
                ha='center', va='bottom', fontsize=9)

    ax.set_title('Faturamento por Loja', fontsize=16, fontweight='bold')
    ax.set_ylabel('Faturamento (R$)', fontsize=12)
    ax.set_xlabel('Loja', fontsize=12)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    return fig

def gerar_excel_download(df_grouped):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_grouped.to_excel(writer, index=False, sheet_name="Faturamento")
    return output.getvalue()

# Interface Streamlit
st.title("📊 Sistema de Premiação - Faturamento por Loja")

uploaded_file = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

df = read_file(uploaded_file, fallback_path=FALLBACK_DESVEND)

if not df.empty:
    df_grouped = processar_dados(df)
    
    # Mostrar tabela
    st.subheader("Tabela Consolidada")
    st.dataframe(df_grouped, use_container_width=True)
    
    # Mostrar gráfico
    st.subheader("Gráfico de Faturamento")
    fig = gerar_grafico(df_grouped)
    st.pyplot(fig)
    
    # Botão de download
    st.download_button(
        label="📥 Baixar Excel Consolidado",
        data=gerar_excel_download(df_grouped),
        file_name="faturamento_por_loja.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Nenhum dado encontrado. Envie um arquivo válido.")
