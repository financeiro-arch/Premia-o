import streamlit as st
import pandas as pd
import os

# Caminho fixo para o arquivo DesVend salvo localmente
FALLBACK_DESVEND = os.path.join("data", "DesVend AUDITORIA_AUTOMATICA.xlsx")

st.set_page_config(page_title="Sistema de Premiação", layout="wide")

@st.cache_data
def read_file(uploaded_file=None, fallback_path=None):
    """
    Lê arquivo CSV, XLS ou XLSX do upload ou fallback path.
    Retorna DataFrame com colunas normalizadas.
    """
    df = pd.DataFrame()
    try:
        if uploaded_file is not None:
            file_name = uploaded_file.name.lower()
            if file_name.endswith(".csv"):
                # Tentar ler CSV com diferentes codificações
                try:
                    df = pd.read_csv(uploaded_file, sep=None, engine="python", encoding="utf-8", dtype=str)
                except UnicodeDecodeError:
                    # fallback para latin1
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, sep=None, engine="python", encoding="latin1", dtype=str)
            elif file_name.endswith((".xls", ".xlsx")):
                df = pd.read_excel(uploaded_file, dtype=str)
            else:
                st.error("Formato de arquivo não suportado.")
                return pd.DataFrame()
        elif fallback_path is not None:
            if fallback_path.lower().endswith(".csv"):
                df = pd.read_csv(fallback_path, sep=None, engine="python", encoding="latin1", dtype=str)
            elif fallback_path.lower().endswith((".xls", ".xlsx")):
                df = pd.read_excel(fallback_path, dtype=str)
            else:
                st.error("Arquivo de fallback com formato não suportado.")
                return pd.DataFrame()
        else:
            st.error("Nenhum arquivo fornecido para leitura.")
            return pd.DataFrame()

        # Normaliza colunas
        df.columns = df.columns.str.strip().str.upper()

    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
        return pd.DataFrame()

    return df

def faturamento_tab(df_desvend):
    st.subheader("📈 Relatório de Faturamento")

    colunas_necessarias = [
        "LOJA", "VENDEDOR", "COTA TOTAL", "TOTAL VENDAS", "SALDO COTA TOTAL", "TICK MEDIO"
    ]

    faltantes = [c for c in colunas_necessarias if c not in df_desvend.columns]
    if faltantes:
        st.error(f"Colunas ausentes no DesVend: {', '.join(faltantes)}")
        return

    # Conversão para numérico
    for col in ["COTA TOTAL", "TOTAL VENDAS", "SALDO COTA TOTAL", "TICK MEDIO"]:
        df_desvend[col] = pd.to_numeric(df_desvend[col], errors="coerce")

    # Cálculos de percentual
    df_desvend["% TOTAL VENDAS"] = (df_desvend["TOTAL VENDAS"] / df_desvend["COTA TOTAL"]) * 100
    df_desvend["% SALDO COTA"] = (df_desvend["SALDO COTA TOTAL"] / df_desvend["COTA TOTAL"]) * 100

    df_resultado = df_desvend[[
        "LOJA", "VENDEDOR", "COTA TOTAL", "TOTAL VENDAS", "% TOTAL VENDAS",
        "SALDO COTA TOTAL", "% SALDO COTA", "TICK MEDIO"
    ]]

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

def premiação_tab():
    st.subheader("🏆 Aba Premiação")
    st.info("Em construção — aguardando definição das regras exatas de cálculo.")

def main():
    st.title("📊 Sistema de Faturamento e Premiação")

    # Carregar DesVend automaticamente
    df_desvend = read_file(fallback_path=FALLBACK_DESVEND)
    if df_desvend.empty:
        st.error(f"Não foi possível carregar o arquivo {FALLBACK_DESVEND}.")
        st.stop()
    else:
        st.success(f"✅ DesVend carregado com sucesso! ({df_desvend.shape[0]} linhas)")

    # Upload Talões Pendentes
    taloes_file = st.file_uploader("Carregar arquivo Talões Pendentes (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"])
    df_taloes = pd.DataFrame()
    if taloes_file:
        df_taloes = read_file(uploaded_file=taloes_file)
        if df_taloes.empty:
            st.error("Arquivo Talões Pendentes inválido ou vazio.")
            st.stop()
        else:
            st.success(f"✅ Talões Pendentes carregado com sucesso! ({df_taloes.shape[0]} linhas)")
            if "CODFIL" not in df_taloes.columns:
                st.error("Arquivo Talões Pendentes não contém a coluna obrigatória 'CODFIL'.")
                st.stop()
            else:
                st.info("Coluna 'CODFIL' encontrada.")

    tab1, tab2 = st.tabs(["📈 Faturamento", "🏆 Premiação"])

    with tab1:
        faturamento_tab(df_desvend)

    with tab2:
        premiação_tab()

if __name__ == "__main__":
    main()
