import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt

st.set_page_config(page_title="Relatório de Faturamento & Premiações", layout="wide")

# -------------------- Função para exportar Excel --------------------
def gerar_excel_download(df, nome_aba):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=nome_aba, index=False)

        workbook = writer.book
        worksheet = writer.sheets[nome_aba]

        # Formatos
        formato_moeda = workbook.add_format({"num_format": "R$ #,##0.00"})
        formato_percent = workbook.add_format({"num_format": "0.0%"})
        formato_cabecalho = workbook.add_format({"bold": True, "bg_color": "#D9D9D9"})

        # Cabeçalho
        worksheet.set_row(0, None, formato_cabecalho)

        # Ajustar largura das colunas e aplicar formatos
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, max_len)
            if "COTA" in col or "VENDAS" in col or "VALOR" in col or "TOTAL" in col:
                worksheet.set_column(idx, idx, max_len, formato_moeda)
            if "%" in col:
                worksheet.set_column(idx, idx, max_len, formato_percent)

    output.seek(0)
    return output

# -------------------- Aba FATURAMENTO --------------------
st.title("Relatório de Faturamento & Premiações")
tab1, tab2 = st.tabs(["📊 Faturamento", "🏆 Premiações"])

with tab1:
    st.subheader("Faturamento")
    arquivo = st.file_uploader("Envie a planilha de Faturamento (DesVend)", type=["xlsx"], key="fat_file")

    if arquivo:
        df = pd.read_excel(arquivo)

        # Consolidar por LOJA
        consolidado = df.groupby("LOJA").agg({
            "COTA TOTAL": "sum",
            "TOTAL VENDAS": "sum",
            "SALDO COTA TOTAL": "sum",
            "QUANT VENDAS": "sum"
        }).reset_index()

        # Calcular colunas
        consolidado["% VENDAS"] = consolidado["TOTAL VENDAS"] / consolidado["COTA TOTAL"]
        consolidado["% SALDO COTA"] = consolidado["SALDO COTA TOTAL"] / consolidado["COTA TOTAL"]
        consolidado["TICK MEDIO"] = consolidado["TOTAL VENDAS"] / consolidado["QUANT VENDAS"]

        # Exibir
        st.dataframe(consolidado.style.format({
            "COTA TOTAL": "R$ {:,.2f}",
            "TOTAL VENDAS": "R$ {:,.2f}",
            "% VENDAS": "{:.1%}",
            "SALDO COTA TOTAL": "R$ {:,.2f}",
            "% SALDO COTA": "{:.1%}",
            "TICK MEDIO": "R$ {:,.2f}"
        }))

        # Gráfico
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.bar(consolidado["LOJA"], consolidado["TOTAL VENDAS"], color="royalblue")
        ax.set_ylabel("Total Vendas (R$)")
        ax.set_title("Faturamento por Loja")
        plt.xticks(rotation=45)
        st.pyplot(fig)

        # Download Excel
        st.download_button(
            "⬇️ Baixar Faturamento",
            data=gerar_excel_download(consolidado, "Faturamento"),
            file_name="Faturamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -------------------- Aba PREMIAÇÕES --------------------
with tab2:
    st.subheader("Premiações")

    arquivo_fat = st.file_uploader("Envie o arquivo de Faturamento (processado)", type=["xlsx"], key="prem_fat")
    arquivo_taloes = st.file_uploader("Envie o arquivo de TALÕES PENDENTES", type=["xlsx"], key="prem_taloes")

    if arquivo_fat and arquivo_taloes:
        df_fat = pd.read_excel(arquivo_fat)
        df_taloes = pd.read_excel(arquivo_taloes)

        # Base inicial
        prem = df_fat[["LOJA", "COTA TOTAL", "TOTAL VENDAS", "% VENDAS", "SALDO COTA TOTAL", "% SALDO COTA"]].copy()

        # Junta com talões pendentes
        if "VENDAS FORA DA POLÍTICA" in df_taloes.columns:
            prem = prem.merge(df_taloes[["LOJA", "VENDAS FORA DA POLÍTICA"]], on="LOJA", how="left")
        else:
            prem["VENDAS FORA DA POLÍTICA"] = 0

        # Calcula vendas atualizadas
        prem["VENDAS ATUALIZADAS"] = prem["TOTAL VENDAS"] - prem["VENDAS FORA DA POLÍTICA"]
        prem["% VENDAS ATUALIZADAS"] = prem["VENDAS ATUALIZADAS"] / prem["COTA TOTAL"]

        # Escolha de modo
        modo = st.radio("Escolha o modo de preenchimento de premiação:", ["Manual", "Automático"])

        if modo == "Manual":
            premiado_list = []
            valor_list = []
            for i, row in prem.iterrows():
                col1, col2 = st.columns([2, 2])
                with col1:
                    premiado = st.selectbox(f"Loja {row['LOJA']} premiada?", ["NÃO", "SIM"], key=f"prem_{i}")
                with col2:
                    valor = st.number_input(f"Valor premiação loja {row['LOJA']}", min_value=0.0, step=50.0, key=f"val_{i}")
                premiado_list.append(premiado)
                valor_list.append(valor)

        else:
            perc_min = st.number_input("Informe o % mínimo para premiar", min_value=0.0, max_value=100.0, step=1.0, value=45.0)
            valor_fix = st.number_input("Informe o valor fixo da premiação", min_value=0.0, step=50.0, value=100.0)
            premiado_list = ["SIM" if x >= perc_min/100 else "NÃO" for x in prem["% VENDAS ATUALIZADAS"]]
            valor_list = [valor_fix if p == "SIM" else 0 for p in premiado_list]

        prem["PREMIADO"] = premiado_list
        prem["VALOR"] = valor_list
        prem["TOTAL LOJA"] = prem["TOTAL VENDAS"] + prem["VALOR"]

        # Exibir tabela formatada
        st.dataframe(prem.style.format({
            "COTA TOTAL": "R$ {:,.2f}",
            "TOTAL VENDAS": "R$ {:,.2f}",
            "% VENDAS": "{:.1%}",
            "SALDO COTA TOTAL": "R$ {:,.2f}",
            "% SALDO COTA": "{:.1%}",
            "VENDAS FORA DA POLÍTICA": "R$ {:,.2f}",
            "VENDAS ATUALIZADAS": "R$ {:,.2f}",
            "% VENDAS ATUALIZADAS": "{:.1%}",
            "VALOR": "R$ {:,.2f}",
            "TOTAL LOJA": "R$ {:,.2f}"
        }))

        # Download Excel
        st.download_button(
            "⬇️ Baixar Premiações",
            data=gerar_excel_download(prem, "Premiações"),
            file_name="Premiacoes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
