import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import os

# Configuração da página
st.set_page_config(page_title="Sistema de Premiação", layout="wide")

# Caminho para fallback local
FALLBACK_PATH = os.path.join("data", "DesVend AUDITORIA_AUTOMATICA.xlsx")

@st.cache_data
def read_auditoria(uploaded_file=None, fallback_path=None):
    """Lê a planilha AUDITORIA e retorna DataFrame limpo."""
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name="AUDITORIA")
    elif fallback_path and os.path.exists(fallback_path):
        df = pd.read_excel(fallback_path, sheet_name="AUDITORIA")
    else:
        return pd.DataFrame()

    # Seleciona apenas as colunas relevantes
    df = df.iloc[2:, [0,1,2,3,5,6]]
    df.columns = ["LOJA", "COTA", "VENDAS", "% VENDAS", "VENDAS ATUALIZADAS", "% COTA ATUAL"]

    # Remove linhas totalmente vazias
    df = df.dropna(how="all")

    # Converte colunas numéricas
    for col in ["COTA", "VENDAS", "% VENDAS", "VENDAS ATUALIZADAS", "% COTA ATUAL"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Remove linhas onde LOJA esteja vazia
    df = df.dropna(subset=["LOJA"], how="all")

    return df

def gerar_excel_download_formatado(df):
    """Gera o arquivo Excel com formatação automática."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Relatório")

    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    # Ajustar largura das colunas
    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

    # Ajustar altura das linhas
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 18

    # Sombreamento cinza na coluna LOJA
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    for row in range(2, ws.max_row + 1):
        ws[f"A{row}"].fill = gray_fill

    # Formatar como moeda
    col_moeda = ["COTA", "VENDAS", "VENDAS ATUALIZADAS"]
    for idx, col_name in enumerate(df.columns, start=1):
        if col_name in col_moeda:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=idx).number_format = 'R$ #,##0.00'

    # Formatar como percentual
    col_perc = ["% VENDAS", "% COTA ATUAL"]
    for idx, col_name in enumerate(df.columns, start=1):
        if col_name in col_perc:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=idx).number_format = '0.0%'

    output_final = BytesIO()
    wb.save(output_final)
    return output_final.getvalue()

def formatar_dataframe_para_exibicao(df):
    """Retorna DataFrame formatado visualmente para exibição no Streamlit."""
    df_formatado = df.copy()
    # Formata moedas
    for col in ["COTA", "VENDAS", "VENDAS ATUALIZADAS"]:
        df_formatado[col] = df_formatado[col].apply(lambda x: f"R$ {x:,.2f}" if pd.notnull(x) else "")
    # Formata percentuais
    for col in ["% VENDAS", "% COTA ATUAL"]:
        df_formatado[col] = df_formatado[col].apply(lambda x: f"{x*100:.1f}%" if pd.notnull(x) else "")
    return df_formatado

# Interface principal
st.title("📊 Sistema de Premiação - Relatório AUDITORIA")

uploaded_file = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])
df = read_auditoria(uploaded_file, fallback_path=FALLBACK_PATH)

if not df.empty:
    st.subheader("Tabela Consolidada")
    st.dataframe(formatar_dataframe_para_exibicao(df), use_container_width=True)

    # Agrupar por LOJA para o gráfico
    df_loja = df.groupby("LOJA", as_index=False).agg({"VENDAS": "sum"})

    st.subheader("Gráfico de Vendas por Loja")
    fig = px.bar(df_loja.sort_values("VENDAS", ascending=False),
                 x="LOJA", y="VENDAS",
                 text=df_loja["VENDAS"].apply(lambda x: f"R$ {x:,.0f}"),
                 labels={"LOJA": "Loja", "VENDAS": "Vendas (R$)"},
                 title="Vendas por Loja")
    fig.update_traces(textposition='outside', marker_color='royalblue')
    st.plotly_chart(fig, use_container_width=True)

    # Botão para download do Excel formatado
    st.download_button(
        label="📥 Baixar Excel Consolidado",
        data=gerar_excel_download_formatado(df),
        file_name="relatorio_auditoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Nenhum dado encontrado. Envie um arquivo válido.")
