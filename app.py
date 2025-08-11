import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Relatório de Faturamento & Premiações", layout="wide")

st.title("📊 Relatório de Faturamento & Premiações")

# Criar abas
aba_faturamento, aba_premiacoes = st.tabs(["Faturamento", "Premiações"])

def gerar_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Faturamento', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Faturamento']

        # Ajuste automático da largura
        for idx, col in enumerate(df.columns):
            col_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, col_width)

        # Formatos
        formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})
        formato_percentual = workbook.add_format({'num_format': '0.0%'})
        formato_header = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9'})

        # Cabeçalho cinza
        worksheet.set_row(0, None, formato_header)

        # Aplicar formatos
        colunas_moeda = ["COTA TOTAL", "TOTAL VENDAS", "SALDO COTA", "TICK MEDIO"]
        colunas_percentual = ["% VENDAS", "% SALDO COTA"]

        for idx, col in enumerate(df.columns):
            if col in colunas_moeda:
                worksheet.set_column(idx, idx, None, formato_moeda)
            elif col in colunas_percentual:
                worksheet.set_column(idx, idx, None, formato_percentual)

    output.seek(0)
    return output

with aba_faturamento:
    st.subheader("📈 Faturamento")

    # Upload do arquivo DesVend
    arquivo = st.file_uploader("Envie a planilha DesVend", type=["xlsx", "xls"])

    if arquivo is not None:
        df = pd.read_excel(arquivo)

        colunas_necessarias = ["LOJA", "COTA TOTAL", "TOTAL VENDAS", "SALDO COTA TOTAL", "QUANT VENDAS"]
        if all(col in df.columns for col in colunas_necessarias):
            consolidado = df.groupby("LOJA", as_index=False).agg({
                "COTA TOTAL": "sum",
                "TOTAL VENDAS": "sum",
                "SALDO COTA TOTAL": "sum",
                "QUANT VENDAS": "sum"
            })

            consolidado["% VENDAS"] = consolidado["TOTAL VENDAS"] / consolidado["COTA TOTAL"]
            consolidado["% SALDO COTA"] = consolidado["SALDO COTA TOTAL"] / consolidado["COTA TOTAL"]
            consolidado["TICK MEDIO"] = consolidado["TOTAL VENDAS"] / consolidado["QUANT VENDAS"]

            consolidado.rename(columns={"SALDO COTA TOTAL": "SALDO COTA"}, inplace=True)

            # Formatar para exibição no Streamlit
            def moeda(v): return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            def perc(v): return f"{v:.1%}"

            tabela_formatada = consolidado.copy()
            tabela_formatada["COTA TOTAL"] = tabela_formatada["COTA TOTAL"].apply(moeda)
            tabela_formatada["TOTAL VENDAS"] = tabela_formatada["TOTAL VENDAS"].apply(moeda)
            tabela_formatada["SALDO COTA"] = tabela_formatada["SALDO COTA"].apply(moeda)
            tabela_formatada["% VENDAS"] = tabela_formatada["% VENDAS"].apply(perc)
            tabela_formatada["% SALDO COTA"] = tabela_formatada["% SALDO COTA"].apply(perc)
            tabela_formatada["TICK MEDIO"] = tabela_formatada["TICK MEDIO"].apply(moeda)

            st.dataframe(tabela_formatada, use_container_width=True)

            # Gráfico de barras
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(consolidado["LOJA"], consolidado["% VENDAS"] * 100, color="#4CAF50")
            ax.set_ylabel("% Vendas")
            ax.set_xlabel("Loja")
            ax.set_title("Percentual de Vendas por Loja")
            ax.set_ylim(0, 120)

            for i, v in enumerate(consolidado["% VENDAS"] * 100):
                ax.text(i, v + 1, f"{v:.1f}%", ha="center")

            st.pyplot(fig)

            # Botão para download do Excel
            st.download_button(
                label="📥 Baixar Excel Consolidado",
                data=gerar_excel_download(consolidado),
                file_name="Faturamento_Consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.error("⚠️ A planilha não contém todas as colunas necessárias.")

    else:
        st.info("📂 Envie o arquivo para visualizar o relatório.")

with aba_premiacoes:
    st.subheader("🏆 Premiações")
    st.info("Área de premiações ainda em desenvolvimento.")
