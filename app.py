import streamlit as st
import pandas as pd
import os
import io
import csv

FALLBACK_DESVEND = '/mnt/data/DesVend.CSV'
FALLBACK_TALOES = '/mnt/data/TALÕES PENDENTES.xlsx'

# Modelo padrão embutido da planilha DesVend AUDITORIA_AUTOMATICA
MODELO_PREMIACAO_PADRAO = pd.DataFrame({
    'Faixa': ['Faixa 1', 'Faixa 2'],
    'Percentual': [5, 10],
    'ValorFixo': [100, 200]
})

@st.cache_data(show_spinner=False)
def read_file(file, fallback_path):
    if file is None:
        if os.path.exists(fallback_path):
            try:
                ext = os.path.splitext(fallback_path)[1].lower()
                if ext in ['.xls', '.xlsx']:
                    df = pd.read_excel(fallback_path, dtype=str)
                else:
                    with open(fallback_path, 'r', encoding='latin1') as f:
                        sample = f.read(1024)
                    delimiter = ','
                    try:
                        delimiter = csv.Sniffer().sniff(sample).delimiter
                    except Exception:
                        pass
                    df = pd.read_csv(fallback_path, sep=delimiter, dtype=str, encoding='latin1')
                return df
            except Exception as e:
                st.error(f"Erro ao ler arquivo padrão: {fallback_path}\n{e}")
                return pd.DataFrame()
        else:
            st.warning(f"Arquivo padrão não encontrado: {fallback_path}. Por favor, faça upload.")
            return pd.DataFrame()

    fname = file.name.lower()
    try:
        if fname.endswith('.csv'):
            content = file.getvalue() if hasattr(file, 'getvalue') else file.read()
            if isinstance(content, bytes):
                try:
                    content_str = content.decode('utf-8')
                except UnicodeDecodeError:
                    content_str = content.decode('latin1')
            else:
                content_str = content
            delimiter = ','
            try:
                delimiter = csv.Sniffer().sniff(content_str.splitlines()[0]).delimiter
            except Exception:
                pass
            df = pd.read_csv(io.StringIO(content_str), sep=delimiter, dtype=str)
            return df
        elif fname.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file, dtype=str)
            return df
        else:
            st.error("Formato do arquivo não suportado. Use CSV, XLS ou XLSX.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao ler arquivo {file.name}: {e}")
        return pd.DataFrame()

def validar_colunas(df, colunas_necessarias, nome_arquivo):
    faltantes = [c for c in colunas_necessarias if c not in df.columns]
    if faltantes:
        st.error(f"Arquivo '{nome_arquivo}' está faltando colunas necessárias: {faltantes}")
        return False
    return True

def filtrar_faturamento(df_desvend, df_taloes):
    lojas_validas = df_taloes['CODFIL'].dropna().unique()
    return df_desvend[df_desvend['LOJA'].isin(lojas_validas)]

def calcular_premiacao(df_filtrado, premiacoes):
    resumo = df_filtrado.groupby('VENDEDOR')['TOTAL VENDAS'].apply(lambda x: pd.to_numeric(x, errors='coerce').fillna(0).sum()).reset_index()
    resumo['Premiação Calculada'] = 0.0

    for faixa in premiacoes:
        resumo['Premiação Calculada'] += resumo['TOTAL VENDAS'] * (faixa['percentual'] / 100) + faixa['valor_fixo']

    return resumo

def destaque_premiacao(val):
    if val > 500:
        color = 'background-color: #9AE69A'  # verde claro
    elif val > 200:
        color = 'background-color: #FFF59D'  # amarelo claro
    else:
        color = ''
    return color

def exportar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
        writer.save()
    return output.getvalue()

def main():
    st.set_page_config(page_title="Sistema Faturamento e Premiação", layout="wide")
    st.title("Sistema de Faturamento e Premiação")

    with st.sidebar.expander("Upload dos arquivos"):
        desvend_file = st.file_uploader("Upload DesVend (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"], help="Arquivo com dados de vendas")
        taloes_file = st.file_uploader("Upload Talões Pendentes (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"], help="Arquivo com lojas válidas")
        # Upload DesVend AUDITORIA_AUTOMATICA desabilitado (modelo embutido)

    with st.spinner("Lendo arquivos..."):
        df_desvend = read_file(desvend_file, FALLBACK_DESVEND)
        df_taloes = read_file(taloes_file, FALLBACK_TALOES)
        df_auditoria = MODELO_PREMIACAO_PADRAO.copy()

    # Mostrar colunas do Talões Pendentes para debug
    st.write("Colunas encontradas no arquivo Talões Pendentes:", df_taloes.columns.tolist())

    if not validar_colunas(df_desvend, ['LOJA', 'VENDEDOR', 'TOTAL VENDAS'], 'DesVend'):
        st.stop()

    if not validar_colunas(df_taloes, ['CODFIL'], 'Talões Pendentes'):
        st.error("Arquivo Talões Pendentes não contém a coluna 'CODFIL'.")
        st.stop()

    tabs = st.tabs(["Faturamento", "Premiação"])

    with tabs[0]:
        st.header("Faturamento")

        df_filtrado = filtrar_faturamento(df_desvend, df_taloes)
        if df_filtrado.empty:
            st.warning("Nenhum dado após filtro de lojas.")
        else:
            vendedores = sorted(df_filtrado['VENDEDOR'].dropna().unique())
            with st.sidebar.form(key="filtro_vendedores"):
                vendedor_selec = st.multiselect("Filtrar por Vendedores", options=vendedores, default=vendedores)
                btn_filtro = st.form_submit_button("Aplicar filtro")
            if btn_filtro:
                df_filtrado = df_filtrado[df_filtrado['VENDEDOR'].isin(vendedor_selec)]

            st.dataframe(df_filtrado.reset_index(drop=True))

            csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
            xlsx_data = exportar_excel(df_filtrado)

            st.download_button("Download CSV Faturamento", data=csv_data, file_name="faturamento.csv", mime="text/csv")
            st.download_button("Download Excel Faturamento", data=xlsx_data, file_name="faturamento.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tabs[1]:
        st.header("Premiação")

        df_filtrado = filtrar_faturamento(df_desvend, df_taloes)
        if df_filtrado.empty:
            st.warning("Nenhum dado após filtro de lojas.")
            st.stop()

        st.markdown("### Configuração de Premiação")
        with st.form(key="form_premiacao"):
            prem_config_txt = st.text_area(
                "Informe as faixas no formato: Nome,Percentual(%),ValorFixo\nExemplo:\nFaixa 1,5,100\nFaixa 2,10,200",
                height=160,
                value="\n".join([f"{row.Faixa},{row.Percentual},{row.ValorFixo}" for _, row in df_auditoria.iterrows()])
            )
            btn_prem = st.form_submit_button("Calcular Premiação")

        if btn_prem:
            premiacoes = []
            linhas = prem_config_txt.strip().split('\n')
            erro_prem = False
            for linha in linhas:
                partes = [p.strip() for p in linha.split(',')]
                if len(partes) != 3:
                    st.error(f"Linha inválida na configuração de premiação: '{linha}'")
                    erro_prem = True
                    break
                try:
                    nome = partes[0]
                    percentual = float(partes[1])
                    valor_fixo = float(partes[2])
                    premiacoes.append({'nome': nome, 'percentual': percentual, 'valor_fixo': valor_fixo})
                except:
                    st.error(f"Erro ao converter valores na linha: '{linha}'")
                    erro_prem = True
                    break
            if not erro_prem:
                df_prem = calcular_premiacao(df_filtrado, premiacoes)
                st.markdown("### Resultado da Premiação")
                st.dataframe(df_prem.style.applymap(destaque_premiacao, subset=['Premiação Calculada']))

                csv_prem = df_prem.to_csv(index=False).encode('utf-8')
                xlsx_prem = exportar_excel(df_prem)

                st.download_button("Download CSV Premiação", data=csv_prem, file_name="premiacao.csv", mime="text/csv")
                st.download_button("Download Excel Premiação", data=xlsx_prem, file_name="premiacao.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
