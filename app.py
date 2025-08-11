import streamlit as st
import pandas as pd
import plotly.express as px
import os
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# Caminho fixo para o arquivo DesVend salvo localmente
FALLBACK_DESVEND = os.path.join("data", "DesVend AUDITORIA_AUTOMATICA.xlsx")

st.set_page_config(page_title="Sistema de Premiação", layout="wide")

@st.cache_data
def read_auditoria(uploaded_file=None, fallback_path=None):
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name="AUDITORIA")
    elif fallback_path is not None and os.path.exists(fallback_path):
        df = pd.read_excel(fallback_path, sheet_name="AUDITORIA")
    else:
        return pd.DataFrame()

    # Seleciona apenas colunas necessárias
    df = df.iloc[2:, [0,1,2,3,5,6]]
    df.columns = ["LOJA", "COTA", "VENDAS", "% VENDAS", "VENDAS ATUALIZADAS", "% COTA ATUAL"]
    df = df.dropna(how="all")

    # Converte colunas numéricas
    for col in ["COTA", "VENDAS", "% VENDAS", "VENDAS ATUALIZADAS", "% COTA ATUAL"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["LOJA", "COTA", "VENDAS"], how="all")
    return df

def gerar_excel_download_formatado(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Relatório")
    output.seek(0)
    
    wb = load_workbook(output)
    ws = wb.active

    # Ajustar largura das colunas
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    # Ajustar altura das linhas
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 18

    # Sombreamento na coluna LOJA
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    for row in range(2, ws.max_row + 1):
        ws[f"A{row}"].fill = gray_fill

    # Formatar colunas como moeda
    moeda_cols = ["COTA", "VENDAS", "VENDAS ATUALIZADAS"]
    for col_idx, col_name in enumerate(df.columns, start=1):
        if col_name in moeda_cols:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_idx).number_format = 'R$ #,##0.00'

    # Formatar colunas como percentual
    perc_cols = ["% VENDAS", "% COTA ATUAL"]
    for col_idx, col_name in enumerate(df.columns, start=1):
        if col_name in perc_cols:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_idx).number_format = '0.0%'

    output_final = BytesIO()
    wb.save(output_final)
    return output_final.getvalue()

# Interface Streamlit
st.title("📊 Sistema de Premiação - Relatório AUDITORIA")

uploaded_file = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

df = read_auditoria(uploaded_file, fallback_path=FALLBACK_DESVEND)

if not df.empty:
    st.subheader("Tabela Consolidada")
    st.dataframe(df, use_container_width=True)

    # Agrupar apenas por LOJA para o gráfico
    df_loja = df.groupby("LOJA", as_index=False).agg({"VENDAS": "sum"})

    st.subheader("Gráfico de Vendas por Loja")
    fig = px.bar(df_loja.sort_values("VENDAS", ascending=False),
                 x="LOJA", y="VENDAS",
                 text=df_loja["VENDAS"].apply(lambda x: f"R$ {x:,.0f}"),
                 labels={"LOJA": "Loja", "VENDAS": "Vendas (R$)"},
                 title="Vendas por Loja")
    fig.update_traces(textposition='outside', marker_color='royalblue')
    st.plotly_chart(fig, use_container_width=True)

    st.download_button(
        label="📥 Baixar Excel Consolidado",
        data=gerar_excel_download_formatado(df),
        file_name="relatorio_auditoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Nenhum dado encontrado. Envie um arquivo válido.")
