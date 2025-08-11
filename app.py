import streamlit as st
import pandas as pd
import os

# Caminho fixo para o DesVend no projeto
FALLBACK_DESVEND = os.path.join("data", "DesVend_AUDITORIA_AUTOMATICA.xlsx")

st.set_page_config(page_title="Sistema de Premiação", layout="wide")

@st.cache_data
def read_excel_file(file_path):
    """Lê planilha Excel e padroniza colunas"""
    try:
        df = pd.read_excel(file_path, dtype=str)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado: {file_path}")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao ler {file_path}: {e}")
        return pd.DataFrame()

@st.cache_data
def read_uploaded_file(uploaded_file):
    """Lê arquivo enviado (.csv, .xls, .xlsx) e padroniza colunas"""
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file, sep=None, engine="python", encoding="latin1", dtype=str)
        else:
            df = pd.read_excel(uploaded_file, dtype=str)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo enviado: {e}")
        return pd.DataFrame()

def faturamento_tab(df):
    st.subheader("📈 Relatório de Faturamento")

    colunas_necessarias = [
        "LOJA", "VENDEDOR", "COTA TOTAL", "TOTAL VENDAS", "SALDO COTA TOTAL", "TICK MEDIO"
    ]

    # Verifica se todas as colunas existem
    faltantes = [c for c in colunas_necessarias if c not in df.columns]
    if faltantes:
        st.error(f"Colunas ausentes no DesVend: {', '.join(faltantes)}")
        return

    # Conversão para numérico
    for col in ["COTA TOTAL", "TOTAL VENDAS", "SALDO COTA TOTAL", "TICK MEDIO"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Cálculos de % Total Vendas e % Saldo Cota
    df["% TOTAL VENDAS"] = (df["TOTAL VENDAS"] / df["COTA TOTAL"]) * 100
    df["% SALDO COTA"] = (df["SALDO COTA TOTAL"] / df["COTA TOTAL"]) * 100

    # Seleção e ordenação das colunas
    df_resultado = df[[
        "LOJA", "VENDEDOR", "COTA TOTAL", "TOTAL VENDAS", "% TOTAL VENDAS",
        "SALDO COTA TOTAL", "% SALDO COTA", "TICK MEDIO"
    ]]

    # Formatação
    st.dataframe(
        df_resultado.style.format({
            "COTA TOTAL": "R$ {:,.2f}",
            "TOTAL VENDAS": "R$ {:,.2f}",
            "% TOTAL VENDAS": "{:.2f}%",
            "SALDO COTA TOTAL": "R$ {:,.2f}",
            "% SALDO COTA": "{:.2f}%",
            "TICK MEDIO": "R$ {:,.2f}"
        }),
        use_container_width=True
    )

def main():
    st.title("📊 Sistema de Premiação")

    # Leitura automática do DesVend
    df_desvend = read_excel_file(FALLBACK_DESVEND)
    if not df_desvend.empty:
        st.success(f"✅ DesVend carregado com sucesso! ({df_desvend.shape[0]} linhas)")
    else:
        st.error("❌ Não foi possível carregar o DesVend.")
        return

    # Abas
    tab1, tab2 = st.tabs(["📈 Faturamento", "🏆 Premiação"])

    with tab1:
        faturamento_tab(df_desvend)

    with tab2:
        st.info("Aba de premiação ainda será configurada.")

if __name__ == "__main__":
    main()
