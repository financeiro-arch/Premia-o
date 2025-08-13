import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Relatório de Faturamento & Premiações", layout="wide")

# ==========================
# FUNÇÃO PARA FORMATAR EXCEL
# ==========================
def exportar_excel_com_formatacao(df_faturamento, df_premiacoes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        moeda_fmt = workbook.add_format({"num_format": "R$ #,##0.00"})
        perc_fmt = workbook.add_format({"num_format": "0.0%"})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9D9D9"})

        # Aba Faturamento
        df_faturamento.to_excel(writer, sheet_name="Faturamento", index=False)
        ws1 = writer.sheets["Faturamento"]
        for col_num, value in enumerate(df_faturamento.columns.values):
            ws1.write(0, col_num, value, header_fmt)
            col_width = max(df_faturamento[value].astype(str).map(len).max(), len(value)) + 2
            ws1.set_column(col_num, col_num, col_width, moeda_fmt if "COTA" in value or "VENDAS" in value else perc_fmt if "%" in value else None)

        # Aba Premiações
        df_premiacoes.to_excel(writer, sheet_name="Premiações", index=False)
        ws2 = writer.sheets["Premiações"]
        for col_num, value in enumerate(df_premiacoes.columns.values):
            ws2.write(0, col_num, value, header_fmt)
            col_width = max(df_premiacoes[value].astype(str).map(len).max(), len(value)) + 2
            ws2.set_column(col_num, col_num, col_width, moeda_fmt if "COTA" in value or "VENDAS" in value or "VALOR" in value else perc_fmt if "%" in value else None)

    return output.getvalue()

# ==========================
# UPLOAD DE PLANILHAS
# ==========================
st.title("📊 Relatório de Faturamento & Premiações")

uploaded_desvend = st.file_uploader("📂 Envie a planilha DesVend", type=["xlsx"])
uploaded_taloes = st.file_uploader("📂 Envie a planilha TALÕES PENDENTES", type=["xlsx"])

if uploaded_desvend and uploaded_taloes:
    df_desvend = pd.read_excel(uploaded_desvend)
    df_taloes = pd.read_excel(uploaded_taloes)

    # Filtro de LOJA e VENDEDOR
    lojas = sorted(df_desvend["LOJA"].unique())
    vendedores = sorted(df_desvend["VENDEDOR"].unique())
    loja_sel = st.multiselect("Filtrar por LOJA", lojas)
    vendedor_sel = st.multiselect("Filtrar por VENDEDOR", vendedores)

    df_filtrado = df_desvend.copy()
    if loja_sel:
        df_filtrado = df_filtrado[df_filtrado["LOJA"].isin(loja_sel)]
    if vendedor_sel:
        df_filtrado = df_filtrado[df_filtrado["VENDEDOR"].isin(vendedor_sel)]

    # Agrupamento por LOJA e VENDEDOR
    df_faturamento = df_filtrado.groupby(["LOJA", "VENDEDOR"], as_index=False).agg({
        "COTA TOTAL": "sum",
        "TOTAL VENDAS": "sum",
        "QUANT VENDAS": "sum",
        "SALDO COTA TOTAL": "sum"
    })

    df_faturamento["% VENDAS"] = df_faturamento["TOTAL VENDAS"] / df_faturamento["COTA TOTAL"]
    df_faturamento["TICK MEDIO"] = df_faturamento["TOTAL VENDAS"] / df_faturamento["QUANT VENDAS"]
    df_faturamento["% SALDO COTA"] = df_faturamento["SALDO COTA TOTAL"] / df_faturamento["COTA TOTAL"]

    # Totais por loja
    df_totais_loja = df_faturamento.groupby("LOJA").sum(numeric_only=True).reset_index()
    df_totais_loja["% VENDAS"] = df_totais_loja["TOTAL VENDAS"] / df_totais_loja["COTA TOTAL"]
    df_totais_loja["TICK MEDIO"] = df_totais_loja["TOTAL VENDAS"] / df_totais_loja["QUANT VENDAS"]
    df_totais_loja["% SALDO COTA"] = df_totais_loja["SALDO COTA TOTAL"] / df_totais_loja["COTA TOTAL"]

    mostrar_totais = st.toggle("Mostrar/ocultar totais por loja", value=True)
    if mostrar_totais:
        df_faturamento = pd.concat([df_faturamento, df_totais_loja.assign(VENDEDOR="TOTAL LOJA")])

    # ==========================
    # GRÁFICO HORIZONTAL
    # ==========================
    st.subheader("📈 Gráfico de Faturamento")
    metricas_disp = ["TOTAL VENDAS", "COTA TOTAL", "QUANT VENDAS", "TICK MEDIO"]
    metrica_sel = st.selectbox("Escolha a métrica para o gráfico:", metricas_disp, index=0)

    df_grafico = df_totais_loja.sort_values(by=metrica_sel, ascending=False)
    fig, ax = plt.subplots(figsize=(12, 6))
    barras = ax.barh(df_grafico["LOJA"], df_grafico[metrica_sel], color=plt.cm.Blues(range(len(df_grafico))))
    ax.invert_yaxis()
    ax.set_xlabel(metrica_sel)
    ax.set_ylabel("LOJA")
    ax.bar_label(barras, fmt="%.2f")
    st.pyplot(fig)

    # ==========================
    # TABELA CONSOLIDADA
    # ==========================
    st.subheader("📄 Tabela Consolidada - Faturamento")
    st.dataframe(df_faturamento.style.format({
        "COTA TOTAL": "R$ {:,.2f}",
        "TOTAL VENDAS": "R$ {:,.2f}",
        "SALDO COTA TOTAL": "R$ {:,.2f}",
        "TICK MEDIO": "R$ {:,.2f}",
        "% VENDAS": "{:.1%}",
        "% SALDO COTA": "{:.1%}"
    }))

    # ==========================
    # EXPORTAÇÃO
    # ==========================
    st.download_button(
        label="💾 Baixar Relatório Excel",
        data=exportar_excel_com_formatacao(df_faturamento, df_faturamento),  # Aqui ainda precisa integrar df_premiacoes real
        file_name="Relatorio_Faturamento_Premiacoes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("📌 Envie as duas planilhas para visualizar o relatório.")
