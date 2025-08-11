import streamlit as st
import pandas as pd
import plotly.express as px
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
    
    # Criar gráfico com Plotly
    st.subheader("Gráfico de Faturamento")
    fig = px.bar(df_grouped, 
                 x=df_grouped['LOJA'].astype(str), 
                 y="TOTAL VENDAS", 
                 text=df_grouped.apply(lambda row: f"R$ {row['TOTAL VENDAS']:,.0f} ({row['%']:.1f}%)", axis=1),
                 labels={"LOJA": "Loja", "TOTAL VENDAS": "Faturamento (R$)"},
                 title="Faturamento por Loja")
    fig.update_traces(textposition='outside', marker_color='royalblue')
    fig.update_layout(yaxis_title="Faturamento (R$)", xaxis_title="Loja", uniformtext_minsize=8, uniformtext_mode='hide')
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Botão de download
    st.download_button(
        label="📥 Baixar Excel Consolidado",
        data=gerar_excel_download(df_grouped),
        file_name="faturamento_por_loja.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Nenhum dado encontrado. Envie um arquivo válido.")
