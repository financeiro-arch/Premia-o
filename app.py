import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, numbers
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Relatório de Faturamento & Premiações", layout="wide")

# ---------------- Funções auxiliares ----------------

def format_moeda(valor):
    return f"R${valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def format_percentual(valor):
    return f"{valor*100:.1f}%"

def exportar_excel(df, nome_aba):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nome_aba)
        ws = writer.sheets[nome_aba]

        # Formatações
        header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        for col_num, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            col_letter = get_column_letter(col_num)
            ws.column_dimensions[col_letter].width = max(len(str(col_name)) + 2, 15)

        # Formatos de moeda e percentual
        for row in range(2, ws.max_row + 1):
            for col_num, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=row, column=col_num)
                if "COTA" in col_name or "VENDAS" in col_name or "VALOR" in col_name or "TICK" in col_name:
                    if "%" in col_name:
                        cell.number_format = "0.0%"
                    else:
                        cell.number_format = 'R$ #,##0.00'
                elif "%" in col_name:
                    cell.number_format = "0.0%"

    return output.getvalue()

def gerar_grafico(df):
    fig, ax = plt.subplots(figsize=(max(8, len(df['LOJA'].unique()) * 1.2), 6))
    ax.bar(df['LOJA'], df['TOTAL VENDAS'], color="skyblue")
    ax.set_ylabel("Total Vendas (R$)")
    ax.set_xlabel("Loja")
    ax.set_title("Faturamento por Loja")
    for i, v in enumerate(df['TOTAL VENDAS']):
        ax.text(i, v, format_moeda(v), ha='center', va='bottom', fontsize=8)
    st.pyplot(fig)

# ---------------- Aba Faturamento ----------------
st.title("Relatório de Faturamento & Premiações")
aba = st.tabs(["📊 Faturamento", "🏆 Premiações"])

with aba[0]:
    st.subheader("📊 Faturamento")
    arquivo_fat = st.file_uploader("Upload do arquivo DesVend.xlsx", type=["xlsx"], key="fat")
    if arquivo_fat:
        df = pd.read_excel(arquivo_fat)

        # Consolidação
        df_group = df.groupby(['LOJA', 'VENDEDOR'], as_index=False).agg({
            'COTA TOTAL': 'sum',
            'TOTAL VENDAS': 'sum',
            'QUANT VENDAS': 'sum',
            'SALDO COTA TOTAL': 'sum',
            'TICK MEDIO': 'mean'
        })
        df_group['% VENDAS'] = df_group['TOTAL VENDAS'] / df_group['COTA TOTAL']
        df_group['% SALDO COTA'] = df_group['SALDO COTA TOTAL'] / df_group['COTA TOTAL']

        # Ordena e exibe
        colunas_ordem = ['LOJA', 'COTA TOTAL', 'TOTAL VENDAS', 'QUANT VENDAS', '% VENDAS', 'TICK MEDIO', 'SALDO COTA TOTAL', '% SALDO COTA']
        df_group = df_group[colunas_ordem]

        # Formata visualmente
        df_fmt = df_group.copy()
        for col in ['COTA TOTAL', 'TOTAL VENDAS', 'TICK MEDIO', 'SALDO COTA TOTAL']:
            df_fmt[col] = df_fmt[col].apply(format_moeda)
        for col in ['% VENDAS', '% SALDO COTA']:
            df_fmt[col] = df_fmt[col].apply(format_percentual)

        st.dataframe(df_fmt, use_container_width=True)

        gerar_grafico(df_group)

        st.download_button(
            label="📥 Baixar Excel",
            data=exportar_excel(df_group, "Faturamento"),
            file_name="Faturamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------------- Aba Premiações ----------------
with aba[1]:
    st.subheader("🏆 Premiações")
    arquivo_desvend = st.file_uploader("Upload do arquivo DesVend.xlsx", type=["xlsx"], key="prem_fat")
    arquivo_taloes = st.file_uploader("Upload do arquivo TALÕES PENDENTES.xlsx", type=["xlsx"], key="prem_tal")
    
    if arquivo_desvend and arquivo_taloes:
        df_fat = pd.read_excel(arquivo_desvend)
        df_tal = pd.read_excel(arquivo_taloes)

        # Consolidação por LOJA
        df_group = df_fat.groupby('LOJA', as_index=False).agg({
            'COTA TOTAL': 'sum',
            'TOTAL VENDAS': 'sum',
            'SALDO COTA TOTAL': 'sum'
        })
        df_group['% VENDAS'] = df_group['TOTAL VENDAS'] / df_group['COTA TOTAL']
        df_group['% SALDO COTA'] = df_group['SALDO COTA TOTAL'] / df_group['COTA TOTAL']

        # Vincula talões pendentes pelo VENDEDOR
        df_merge = pd.merge(df_fat, df_tal, on='VENDEDOR', how='left')
        df_vendas_forapol = df_merge.groupby('LOJA', as_index=False)['VENDAS FORA DA POLÍTICA'].sum()
        df_group = pd.merge(df_group, df_vendas_forapol, on='LOJA', how='left')

        # Calcula vendas atualizadas
        df_group['VENDAS ATUALIZADAS'] = df_group['TOTAL VENDAS'] - df_group['VENDAS FORA DA POLÍTICA']
        df_group['% VENDAS ATUALIZADAS'] = df_group['VENDAS ATUALIZADAS'] / df_group['COTA TOTAL']

        # Configuração da premiação
        pct_premio = st.number_input("Informe o % mínimo para premiação", min_value=0.0, max_value=1.0, value=0.45, step=0.01)
        valor_premio = st.number_input("Informe o valor da premiação", min_value=0.0, value=100.0, step=10.0)

        df_group['PREMIADO'] = df_group['% VENDAS ATUALIZADAS'] >= pct_premio
        df_group['VALOR'] = df_group['PREMIADO'].apply(lambda x: valor_premio if x else 0)
        df_group['TOTAL LOJA'] = df_group.groupby('LOJA')['VALOR'].transform('sum')

        # Exibe formatado
        df_fmt = df_group.copy()
        for col in ['COTA TOTAL', 'TOTAL VENDAS', 'SALDO COTA TOTAL', 'VENDAS FORA DA POLÍTICA', 'VENDAS ATUALIZADAS', 'VALOR', 'TOTAL LOJA']:
            df_fmt[col] = df_fmt[col].apply(format_moeda)
        for col in ['% VENDAS', '% SALDO COTA', '% VENDAS ATUALIZADAS']:
            df_fmt[col] = df_fmt[col].apply(format_percentual)
        df_fmt['PREMIADO'] = df_fmt['PREMIADO'].apply(lambda x: "SIM" if x else "NÃO")

        st.dataframe(df_fmt, use_container_width=True)

        st.download_button(
            label="📥 Baixar Excel Premiações",
            data=exportar_excel(df_group, "Premiações"),
            file_name="Premiacoes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
