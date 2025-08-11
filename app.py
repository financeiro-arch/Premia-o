import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Relatório de Faturamento & Premiações", layout="wide")

# -----------------------------
# Função para formatar Excel
# -----------------------------
def gerar_excel_download(df, nome_aba):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=nome_aba)

        workbook = writer.book
        worksheet = writer.sheets[nome_aba]

        # Formatos
        moeda_fmt = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        perc_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        normal_fmt = workbook.add_format({'border': 1})

        # Formatar cabeçalho
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Formatar colunas por tipo
        for idx, col in enumerate(df.columns):
            if "COTA" in col or "VENDAS" in col or "VALOR" in col or "TOTAL" in col:
                worksheet.set_column(idx, idx, 15, moeda_fmt)
            elif "%" in col:
                worksheet.set_column(idx, idx, 12, perc_fmt)
            else:
                worksheet.set_column(idx, idx, 20, normal_fmt)

        # Auto ajuste de largura mínima
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

    return output.getvalue()

# -----------------------------
# Aba principal com tabs
# -----------------------------
st.title("Relatório de Faturamento & Premiações")
tab1, tab2 = st.tabs(["📊 Faturamento", "🏆 Premiações"])

# -----------------------------
# Aba Faturamento
# -----------------------------
with tab1:
    st.subheader("Faturamento")
    arquivo_faturamento = st.file_uploader("Envie o arquivo de Faturamento", type=["xlsx"], key="fat")

    if arquivo_faturamento:
        df_fat = pd.read_excel(arquivo_faturamento)

        # Aqui você pode aplicar o tratamento atual da sua tabela Faturamento
        st.dataframe(df_fat)

        st.download_button(
            "⬇️ Baixar Faturamento",
            data=gerar_excel_download(df_fat, "Faturamento"),
            file_name="Faturamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -----------------------------
# Aba Premiações
# -----------------------------
with tab2:
    st.subheader("Premiações")
    arquivo_fat = st.file_uploader("Envie o arquivo de Faturamento", type=["xlsx"], key="prem_fat")
    arquivo_taloes = st.file_uploader("Envie o arquivo de TALÕES PENDENTES", type=["xlsx"], key="prem_taloes")

    if arquivo_fat and arquivo_taloes:
        df_fat = pd.read_excel(arquivo_fat)
        df_taloes = pd.read_excel(arquivo_taloes)

        # Seleciona colunas da Faturamento
        prem = df_fat[["LOJA", "COTA TOTAL", "TOTAL VENDAS", "% VENDAS", "SALDO COTA", "% SALDO COTA"]].copy()

        # Junta com TALÕES PENDENTES
        if "LOJA" in df_taloes.columns:
            prem = prem.merge(df_taloes, on="LOJA", how="left")

        # Calcula colunas novas
        prem["VENDAS ATUALIZADAS"] = prem["TOTAL VENDAS"] - prem["VENDAS FORA DA POLÍTICA"]
        prem["% VENDAS ATUALIZADAS"] = prem["VENDAS ATUALIZADAS"] / prem["COTA TOTAL"]

        # Inputs manuais
        premiado_list = []
        valor_list = []
        for i, row in prem.iterrows():
            col1, col2, col3 = st.columns([3, 2, 2])
            with col1:
                premiado = st.selectbox(f"Loja {row['LOJA']} premiada?", ["NÃO", "SIM"], key=f"prem_{i}")
            with col2:
                valor = st.number_input(f"Valor premiação loja {row['LOJA']}", min_value=0.0, step=50.0, key=f"val_{i}")
            premiado_list.append(premiado)
            valor_list.append(valor)

        prem["PREMIADO"] = premiado_list
        prem["VALOR"] = valor_list
        prem["TOTAL LOJA"] = prem["TOTAL VENDAS"] + prem["VALOR"]

        st.dataframe(prem)

        st.download_button(
            "⬇️ Baixar Premiações",
            data=gerar_excel_download(prem, "Premiações"),
            file_name="Premiacoes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
