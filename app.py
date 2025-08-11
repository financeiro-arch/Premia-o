import streamlit as st
import pandas as pd
import os

# Caminho fixo para o arquivo DesVend salvo no projeto
FALLBACK_DESVEND = os.path.join("data", "DesVend_AUDITORIA_AUTOMATICA.xlsx")

st.set_page_config(page_title="Sistema de Premiação", layout="wide")

@st.cache_data
def read_excel_file(file_path):
    """Lê planilha Excel e padroniza colunas para maiúsculo e sem espaços extras"""
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

def main():
    st.title("📊 Sistema de Premiação")

    # Leitura do DesVend automático (sem upload)
    df_desvend = read_excel_file(FALLBACK_DESVEND)
    if not df_desvend.empty:
        st.success(f"✅ DesVend carregado com sucesso! ({df_desvend.shape[0]} linhas)")
        with st.expander("Ver colunas do DesVend"):
            st.write(list(df_desvend.columns))
    else:
        st.error("❌ Não foi possível carregar o DesVend.")

    # Upload de Talões Pendentes
    taloes_file = st.file_uploader("Carregar arquivo Talões Pendentes (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"])
    if taloes_file:
        df_taloes = read_uploaded_file(taloes_file)

        if not df_taloes.empty:
            st.success(f"✅ Talões Pendentes carregado com sucesso! ({df_taloes.shape[0]} linhas)")
            with st.expander("Ver colunas do Talões Pendentes"):
                st.write(list(df_taloes.columns))

            # Verificação da coluna CODFIL
            if "CODFIL" not in df_taloes.columns:
                st.error("❌ Arquivo Talões Pendentes não contém a coluna obrigatória 'CODFIL'.")
            else:
                st.info("🔍 Coluna 'CODFIL' encontrada. Pronto para filtrar lojas.")
        else:
            st.error("❌ Arquivo Talões Pendentes inválido ou vazio.")

if __name__ == "__main__":
    main()
