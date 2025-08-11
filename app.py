import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="Relatório de Faturamento & Premiações", layout="wide")

st.title("Relatório de Faturamento & Premiações")

aba = st.tabs(["📊 Faturamento", "🏆 Premiações"])

def gerar_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatório')
        workbook = writer.book
        worksheet = writer.sheets['Relatório']
        # Ajustar largura automática das colunas
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, max_len)
    return output.getvalue()

# --- Função para processar DesVend ---
def processar_faturamento(df):
    consolidado = df.groupby("LOJA").agg({
        "COTA TOTAL": "sum",
        "TOTAL VENDAS": "sum",
        "QUANT VENDAS": "sum",
        "SALDO COTA TOTAL": "sum"
    }).reset_index()

    consolidado["% VENDAS"] = consolidado["TOTAL VENDAS"] / consolidado["COTA TOTAL"]
    consolidado["% SALDO COTA"] = consolidado["SALDO COTA TOTAL"] / consolidado["COTA TOTAL"]
    consolidado["TICK MEDIO"] = consolidado["TOTAL VENDAS"] / consolidado["QUANT VENDAS"]

    # Reordena colunas
    consolidado = consolidado[[
        "LOJA", "COTA TOTAL", "TOTAL VENDAS", "QUANT VENDAS",
        "% VENDAS", "TICK MEDIO", "SALDO COTA TOTAL", "% SALDO COTA"
    ]]
    return consolidado

# --- Função para processar Premiações ---
def calcular_premiacoes(df_fat, df_taloes, pct_meta, valor_premio):
    # Merge apenas de vendedores que existem em ambas as planilhas
    df_merge = pd.merge(df_fat, df_taloes, on="VENDEDOR", how="inner")

    # Calcular % vendas atualizadas
    df_merge["% VENDAS ATUALIZADAS"] = df_merge["VENDAS ATUALIZADAS"] / df_merge["COTA TOTAL"]

    # Definir premiados
    df_merge["PREMIADO"] = df_merge["% VENDAS ATUALIZADAS"].apply(lambda x: "SIM" if x >= pct_meta else "NÃO")
    df_merge["VALOR"] = df_merge["PREMIADO"].apply(lambda x: valor_premio if x == "SIM" else 0)

    # Consolidar por loja
    consolidado = df_merge.groupby("LOJA").agg({
        "COTA TOTAL": "sum",
        "TOTAL VENDAS": "sum",
        "% VENDAS": "mean",
        "SALDO COTA TOTAL": "sum",
        "% SALDO COTA": "mean",
        "VENDAS FORA DA POLÍTICA": "sum",
        "VENDAS ATUALIZADAS": "sum",
        "% VENDAS ATUALIZADAS": "mean",
        "VALOR": "sum"
    }).reset_index()

    consolidado["TOTAL LOJA"] = consolidado["VALOR"]

    return consolidado

# --- Aba Faturamento ---
with aba[0]:
    arquivo_desvend = st.file_uploader("Carregar planilha DesVend", type=["xlsx"], key="desvend")
    if arquivo_desvend:
        df_desvend = pd.read_excel(arquivo_desvend)
        df_fat = processar_faturamento(df_desvend)

        st.subheader("Tabela Consolidada - Faturamento")
        st.dataframe(df_fat)

        # Gráfico ajustado
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.bar(df_fat["LOJA"], df_fat["TOTAL VENDAS"], color="skyblue")
        ax.set_title("Faturamento por Loja")
        ax.set_ylabel("Total Vendas")
        ax.set_xticklabels(df_fat["LOJA"], rotation=45, ha="right")
        st.pyplot(fig)

        st.download_button(
            label="📥 Baixar Excel Faturamento",
            data=gerar_excel_download(df_fat),
            file_name="faturamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- Aba Premiações ---
with aba[1]:
    arquivo_taloes = st.file_uploader("Carregar planilha Talões Pendentes", type=["xlsx"], key="taloes")
    pct_meta = st.number_input("Informe % para premiação (ex: 0.45 para 45%)", value=0.45, step=0.01)
    valor_premio = st.number_input("Informe o valor da premiação (R$)", value=100.0, step=10.0)

    if arquivo_desvend and arquivo_taloes:
        df_desvend = pd.read_excel(arquivo_desvend)
        df_taloes = pd.read_excel(arquivo_taloes)
        df_prem = calcular_premiacoes(df_desvend, df_taloes, pct_meta, valor_premio)

        st.subheader("Tabela Consolidada - Premiações")
        st.dataframe(df_prem)

        st.download_button(
            label="📥 Baixar Excel Premiações",
            data=gerar_excel_download(df_prem),
            file_name="premiacoes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
