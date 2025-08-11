import streamlit as st
import pandas as pd
import plotly.express as px
import io

# ==============================
# Função para gerar Excel formatado
# ==============================
def gerar_excel_formatado(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Faturamento')
        workbook = writer.book
        worksheet = writer.sheets['Faturamento']

        # Ajustar largura das colunas
        for i, col in enumerate(df.columns):
            col_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, col_width)

        # Formatos
        formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})
        formato_percent = workbook.add_format({'num_format': '0.0%'})
        formato_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})

        # Cabeçalho com cinza claro
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, formato_header)

        # Aplicar formatação
        moeda_cols = ['COTA TOTAL', 'TOTAL VENDAS', 'TICK MEDIO', 'SALDO COTA TOTAL']
        perc_cols = ['% VENDAS', '% SALDO COTA']

        for row_num in range(1, len(df) + 1):
            for col_num, col_name in enumerate(df.columns):
                if col_name in moeda_cols:
                    worksheet.write(row_num, col_num, df[col_name].iloc[row_num-1], formato_moeda)
                elif col_name in perc_cols:
                    worksheet.write(row_num, col_num, df[col_name].iloc[row_num-1], formato_percent)

    output.seek(0)
    return output

# ==============================
# Função para consolidar dados
# ==============================
def consolidar_faturamento(df):
    df_group = df.groupby('LOJA').agg({
        'COTA TOTAL': 'sum',
        'TOTAL VENDAS': 'sum',
        'QUANT VENDAS': 'sum',
        'SALDO COTA TOTAL': 'sum'
    }).reset_index()

    df_group['% VENDAS'] = df_group['TOTAL VENDAS'] / df_group['COTA TOTAL']
    df_group['TICK MEDIO'] = df_group['TOTAL VENDAS'] / df_group['QUANT VENDAS']
    df_group['% SALDO COTA'] = df_group['SALDO COTA TOTAL'] / df_group['COTA TOTAL']

    # Reordenar colunas
    colunas_ordem = [
        'LOJA', 'COTA TOTAL', 'TOTAL VENDAS', 'QUANT VENDAS',
        '% VENDAS', 'TICK MEDIO', 'SALDO COTA TOTAL', '% SALDO COTA'
    ]
    df_group = df_group[colunas_ordem]

    return df_group

# ==============================
# Interface Streamlit
# ==============================
st.set_page_config(page_title="Relatório de Faturamento & Premiações", layout="wide")
st.title("📊 Relatório de Faturamento & Premiações")

aba = st.tabs(["Faturamento", "Premiações"])

with aba[0]:
    st.header("📈 Faturamento")
    uploaded_file = st.file_uploader("Envie o arquivo Excel com os dados de faturamento", type=['xlsx'])

    if uploaded_file:
        df = pd.read_excel(uploaded_file, sheet_name="DesVend")
        consolidado = consolidar_faturamento(df)

        # Exibir tabela formatada
        st.dataframe(
            consolidado.style
            .format({
                'COTA TOTAL': "R$ {:,.2f}",
                'TOTAL VENDAS': "R$ {:,.2f}",
                'TICK MEDIO': "R$ {:,.2f}",
                'SALDO COTA TOTAL': "R$ {:,.2f}",
                '% VENDAS': "{:.1%}",
                '% SALDO COTA': "{:.1%}"
            }),
            use_container_width=True
        )

        # Gráfico dinâmico
        fig = px.bar(
            consolidado,
            x='LOJA',
            y='TOTAL VENDAS',
            text='TOTAL VENDAS',
            title="Faturamento por Loja",
            color='TOTAL VENDAS',
            color_continuous_scale='Blues'
        )
        fig.update_traces(texttemplate='R$ %{y:,.2f}', textposition='outside')
        fig.update_layout(xaxis_tickangle=-45, height=500, uniformtext_minsize=8, uniformtext_mode='hide')
        st.plotly_chart(fig, use_container_width=True)

        # Download Excel
        excel_bytes = gerar_excel_formatado(consolidado)
        st.download_button(
            "📥 Baixar Excel Consolidado",
            data=excel_bytes,
            file_name="faturamento_consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
