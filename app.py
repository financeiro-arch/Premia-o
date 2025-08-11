import streamlit as st
import pandas as pd
import os

FALLBACK_DESVEND = '/mnt/data/DesVend.CSV'
FALLBACK_TALOES = '/mnt/data/TALÕES PENDENTES.xlsx'
FALLBACK_AUDITORIA = '/mnt/data/DesVend AUDITORIA_AUTOMATICA.xlsx'

@st.cache_data(show_spinner=False)
def read_desvend(file):
    if file is None:
        # fallback
        ext = os.path.splitext(FALLBACK_DESVEND)[1].lower()
        if ext in ['.xls', '.xlsx']:
            return pd.read_excel(FALLBACK_DESVEND, dtype=str)
        else:
            return pd.read_csv(FALLBACK_DESVEND, dtype=str, sep=None, engine='python', encoding='latin1')
    else:
        fname = file.name.lower()
        try:
            if fname.endswith('.csv'):
                # tenta ler CSV com alguns encodings
                try:
                    return pd.read_csv(file, dtype=str, sep=None, engine='python', encoding='utf-8')
                except UnicodeDecodeError:
                    try:
                        return pd.read_csv(file, dtype=str, sep=None, engine='python', encoding='latin1')
                    except UnicodeDecodeError:
                        return pd.read_csv(file, dtype=str, sep=None, engine='python', encoding='cp1252')
            elif fname.endswith(('.xls', '.xlsx')):
                return pd.read_excel(file, dtype=str)
            else:
                st.error("Formato de arquivo DesVend não suportado. Use CSV, XLS ou XLSX.")
                return None
        except Exception as e:
            st.error(f"Erro ao ler arquivo DesVend: {e}")
            return None

@st.cache_data(show_spinner=False)
def read_taloes(file):
    if file is None:
        ext = os.path.splitext(FALLBACK_TALOES)[1].lower()
        if ext in ['.xls', '.xlsx']:
            return pd.read_excel(FALLBACK_TALOES, dtype=str)
        else:
            return pd.read_csv(FALLBACK_TALOES, dtype=str, sep=None, engine='python', encoding='latin1')
    else:
        fname = file.name.lower()
        try:
            if fname.endswith('.csv'):
                try:
                    return pd.read_csv(file, dtype=str, sep=None, engine='python', encoding='utf-8')
                except UnicodeDecodeError:
                    try:
                        return pd.read_csv(file, dtype=str, sep=None, engine='python', encoding='latin1')
                    except UnicodeDecodeError:
                        return pd.read_csv(file, dtype=str, sep=None, engine='python', encoding='cp1252')
            elif fname.endswith(('.xls', '.xlsx')):
                return pd.read_excel(file, dtype=str)
            else:
                st.error("Formato de arquivo Talões Pendentes não suportado. Use CSV, XLS ou XLSX.")
                return None
        except Exception as e:
            st.error(f"Erro ao ler arquivo Talões Pendentes: {e}")
            return None

@st.cache_data(show_spinner=False)
def read_auditoria(file):
    if file is None:
        try:
            return pd.read_excel(FALLBACK_AUDITORIA, dtype=str)
        except Exception:
            return pd.DataFrame()
    else:
        try:
            return pd.read_excel(file, dtype=str)
        except Exception as e:
            st.error(f"Erro ao ler arquivo Auditoria: {e}")
            return pd.DataFrame()

def filtrar_faturamento(df_desvend, df_taloes):
    if 'CodFil' not in df_taloes.columns:
        st.error("Arquivo Talões Pendentes não contém a coluna 'CodFil'")
        return pd.DataFrame()
    if 'loja' not in df_desvend.columns:
        st.error("Arquivo DesVend não contém a coluna 'loja'")
        return pd.DataFrame()

    lojas_validas = df_taloes['CodFil'].dropna().unique()
    df_filtrado = df_desvend[df_desvend['loja'].isin(lojas_validas)]
    return df_filtrado

def main():
    st.set_page_config(page_title="Sistema de Faturamento e Premiação", layout="wide")
    st.title("Sistema de Faturamento e Premiação")

    with st.sidebar.expander("Upload dos arquivos"):
        desvend_file = st.file_uploader("Upload DesVend (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"])
        taloes_file = st.file_uploader("Upload Talões Pendentes (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"])
        auditoria_file = st.file_uploader("Upload DesVend AUDITORIA_AUTOMATICA.xlsx", type=["xls", "xlsx"])

    df_desvend = read_desvend(desvend_file)
    df_taloes = read_taloes(taloes_file)
    df_auditoria = read_auditoria(auditoria_file)

    if df_desvend is None or df_desvend.empty:
        st.error("Arquivo DesVend inválido ou não carregado.")
        return
    if df_taloes is None or df_taloes.empty:
        st.error("Arquivo Talões Pendentes inválido ou não carregado.")
        return

    tabs = st.tabs(["Faturamento", "Premiação"])

    with tabs[0]:
        st.header("Faturamento")
        df_filtrado = filtrar_faturamento(df_desvend, df_taloes)

        if df_filtrado.empty:
            st.warning("Nenhum dado encontrado após filtro de lojas.")
        else:
            if 'consultor' in df_filtrado.columns:
                consultores = df_filtrado['consultor'].dropna().unique()
                consultor_selec = st.multiselect("Selecione Consultor(es)", options=consultores, default=consultores)
                df_filtrado = df_filtrado[df_filtrado['consultor'].isin(consultor_selec)]

            st.dataframe(df_filtrado.reset_index(drop=True))

            csv_fat = df_filtrado.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV Faturamento", data=csv_fat, file_name="faturamento_filtrado.csv", mime="text/csv")

    with tabs[1]:
        st.header("Premiação")

        df_filtrado = filtrar_faturamento(df_desvend, df_taloes)

        st.subheader("Configuração de premiação")
        prem_config_txt = st.text_area(
            "Informe as faixas no formato: Nome,Percentual(%),ValorFixo\nExemplo:\nFaixa 1,5,100\nFaixa 2,10,200",
            height=150,
            value="Faixa 1,5,100\nFaixa 2,10,200"
        )

        prem_config = []
        for linha in prem_config_txt.strip().split('\n'):
            partes = [p.strip() for p in linha.split(',')]
            if len(partes) == 3:
                try:
                    nome = partes[0]
                    percentual = float(partes[1])
                    valor_fixo = float(partes[2])
                    prem_config.append({'nome': nome, 'percentual': percentual, 'valor_fixo': valor_fixo})
                except:
                    st.error(f"Erro ao ler linha de premiação: {linha}")

        if 'consultor' not in df_filtrado.columns or 'valor' not in df_filtrado.columns:
            st.warning("Arquivo DesVend deve conter colunas 'consultor' e 'valor' para cálculo de premiação.")
        else:
            resumo = df_filtrado.groupby('consultor')['valor'].apply(lambda x: x.astype(float).sum()).reset_index()
            resumo['Premiação Calculada'] = 0.0

            for faixa in prem_config:
                resumo['Premiação Calculada'] += resumo['valor'] * faixa['percentual'] / 100 + faixa['valor_fixo']

            st.dataframe(resumo)

            csv_prem = resumo.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV Premiação", data=csv_prem, file_name="premiacao_calculada.csv", mime="text/csv")

if __name__ == "__main__":
    main()
