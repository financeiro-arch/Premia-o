import streamlit as st
import pandas as pd
import plotly.express as px
import os
from io import BytesIO

# Caminho fixo para o arquivo DesVend salvo localmente
FALLBACK_DESVEND = os.path.join("data", "DesVend AUDITORIA_AUTOMATICA.xlsx")

st.set_page_config(page_title="Sistema de Premiação", layout="wide")

@st.cache_data
def read_auditoria(uploaded_file=None, fallback_path=None):
    """
    Lê a aba AUDITORIA do Excel, remove linhas extras e retorna apenas as colunas necessárias.
    """
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name="AUDITORIA")
    elif fallback_path is not None and os.path.exists(fallback_path):
        df = pd.read_excel(fallback_path, sheet_name="AUDITORIA")
    else:
        return pd.DataFrame()

    # Pular cabeçalhos
    df = df.iloc[2:, [0,1,2,3,5,6]]
    df.columns = ["LOJA", "COTA", "VENDAS", "% VENDAS", "VENDAS ATUALIZADAS", "% COTA ATUAL"]
    df = df.dropna(how="all")

    # Converter numéricos
    for col in ["COTA", "VENDAS", "% VENDAS", "VENDAS ATUALIZADAS", "% COTA ATUAL"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Remover linhas totalmente vazias
    df = df.dropna(subset=["LOJA", "COTA", "VENDAS"], how="all")
    return df

def gerar_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Relatório")
    return output.getvalue()

# Interface Streamlit
st.title("📊 Sistema de Premiação - Relatório AUDITORIA")

uploaded_file = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

df = read_auditoria(uploaded_file, fallback_path=FALLBACK_DESVEND)

if not df.empty:
    # Mostrar tabela
    st.subheader("Tabela Consolidada")
    st.dataframe(df, use_container_width=True)
    
    # Criar gráfico de VENDAS
    st.subheader("Gráfico de Vendas por Loja")
    fig = px.bar(df.sort_values("VENDAS", ascending=False),
                 x="LOJA", y="VENDAS",
                 text=df["VENDAS"].apply(lambda x: f"R$ {x:,.0f}"),
                 labels={"LOJA": "Loja", "VENDAS": "Vendas (R$)"},
                 title="Vendas por Loja")
    fig.update_traces(textposition='outside', marker_color='royalblue')
    st.plotly_chart(fig, use_container_width=True)

    # Botão de download
    st.download_button(
        label="📥 Baixar Excel Consolidado",
        data=gerar_excel_download(df),
        file_name="relatorio_auditoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Nenhum dado encontrado. Envie um arquivo válido.")
