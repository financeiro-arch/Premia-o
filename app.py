import streamlit as st
import pandas as pd
from io import BytesIO

FALLBACK_DESVEND = '/mnt/data/DesVend.CSV'
FALLBACK_TALOES = '/mnt/data/TALÕES PENDENTES.xlsx'
FALLBACK_AUDITORIA = '/mnt/data/DesVend AUDITORIA_AUTOMATICA.xlsx'

@st.cache_data(show_spinner=False)
def read_csv_either(uploaded, fallback_path):
    if uploaded is not None:
        try:
            return pd.read_csv(uploaded, sep=None, engine='python', dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                return pd.read_csv(uploaded, sep=None, engine='python', dtype=str, encoding='latin1')
            except UnicodeDecodeError:
                return pd.read_csv(uploaded, sep=None, engine='python', dtype=str, encoding='cp1252')
        except Exception as e:
            st.error(f"Erro ao ler CSV: {e}")
            return None
    else:
        return pd.read_csv(fallback_path, sep=None, engine='python', dtype=str)

@st.cache_data(show_spinner=False)
def read_excel_either(uploaded, fallback_path):
    if uploaded is not None:
        try:
            return pd.read_excel(uploaded, dtype=str)
        except Exception as e:
            st.error(f"Erro ao ler Excel: {e}")
            return None
    else:
        return pd.read_excel(fallback_path, dtype=str)

def filtrar_faturamento(df_desvend, df_taloes):
    # Filtra DesVend apenas com lojas presentes em CodFil dos taloes
    lojas_validas = df_taloes['CodFil'].dropna().unique()
    df_filtrado = df_desvend[df_desvend['loja'].isin(lojas_validas)]
    return df_filtrado

def main():
    st.set_page_config(page_title="Sistema de Faturamento e Premiação", layout="wide")
    st.title("Sistema de Faturamento e Premiação")

    # Upload dos arquivos
    with st.sidebar.expander("Upload dos arquivos"):
        csv_upload = st.file_uploader("Upload DesVend.CSV", type=["csv"])
        xlsx_taloes_upload = st.file_uploader("Upload Talões Pendentes.xlsx", type=["xlsx", "xls"])
        xlsx_auditoria_upload = st.file_uploader("Upload DesVend AUDITORIA_AUTOMATICA.xlsx", type=["xlsx", "xls"])

    # Ler arquivos
    df_desvend = read_csv_either(csv_upload, FALLBACK_DESVEND)
    df_taloes = read_excel_either(xlsx_taloes_upload, FALLBACK_TALOES)
    df_auditoria = read_excel_either(xlsx_auditoria_upload, FALLBACK_AUDITORIA)

    if df_desvend is None or df_taloes is None:
        st.error("Por favor, faça o upload dos arquivos DesVend.CSV e Talões Pendentes.xlsx corretamente.")
        return

    tabs = st.tabs(["Faturamento", "Premiação"])

    with tabs[0]:
        st.header("Faturamento")
        df_filtrado = filtrar_faturamento(df_desvend, df_taloes)

        # Filtro por consultor
        if 'consultor' in df_filtrado.columns:
            consultores = df_filtrado['consultor'].dropna().unique()
            consultor_selec = st.multiselect("Selecione Consultor(es)", options=consultores, default=consultores)
            df_filtrado = df_filtrado[df_filtrado['consultor'].isin(consultor_selec)]

        st.dataframe(df_filtrado.reset_index(drop=True))

        csv_fat = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV Faturamento", data=csv_fat, file_name="faturamento_filtrado.csv", mime="text/csv")

    with tabs[1]:
        st.header("Premiação")

        if df_auditoria is None:
            st.warning("Upload do arquivo DesVend AUDITORIA_AUTOMATICA.xlsx não realizado. Será usado modelo vazio.")
            df_auditoria = pd.DataFrame()  # ou crie uma estrutura vazia padrão

        # Reaplicar filtro de lojas no faturamento para base da premiação
        df_filtrado = filtrar_faturamento(df_desvend, df_taloes)

        # Campo para inserir percentuais de premiação (faixas)
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

        # Aqui você pode implementar a lógica de cálculo conforme seu modelo.
        # Exemplo simples: somar percentual * valor de venda + valor fixo por consultor.

        if 'consultor' not in df_filtrado.columns or 'valor' not in df_filtrado.columns:
            st.warning("Arquivo DesVend.CSV deve conter colunas 'consultor' e 'valor' para calcular premiação.")
        else:
            resumo = df_filtrado.groupby('consultor')['valor'].apply(lambda x: x.astype(float).sum()).reset_index()
            resumo['Premiação Calculada'] = 0.0

            for faixa in prem_config:
                # Exemplo: aplica percentual e soma valor fixo para todos consultores
                resumo['Premiação Calculada'] += resumo['valor'] * faixa['percentual'] / 100 + faixa['valor_fixo']

            st.dataframe(resumo)

            csv_prem = resumo.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV Premiação", data=csv_prem, file_name="premiacao_calculada.csv", mime="text/csv")

if __name__ == "__main__":
    main()
