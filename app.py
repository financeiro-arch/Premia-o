import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

st.set_page_config(page_title="Sistema Faturamento & Premiação", layout="wide")

# ---------- Helpers -------------------------------------------------
FALLBACK_DESVEND = "/mnt/data/DesVend.CSV"
FALLBACK_TALOES = "/mnt/data/TALÕES PENDENTES.xlsx"

@st.cache_data
def read_csv_either(uploaded, fallback_path):
    if uploaded is not None:
        return pd.read_csv(uploaded, sep=None, engine='python', dtype=str)
    p = Path(fallback_path)
    if p.exists():
        return pd.read_csv(p, sep=None, engine='python', dtype=str)
    return pd.DataFrame()

@st.cache_data
def read_excel_either(uploaded, fallback_path):
    if uploaded is not None:
        return pd.read_excel(uploaded, engine='openpyxl', dtype=str)
    p = Path(fallback_path)
    if p.exists():
        return pd.read_excel(p, engine='openpyxl', dtype=str)
    return pd.DataFrame()

def find_column(df, candidates):
    if df is None or df.empty:
        return None
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    for col in df.columns:
        low = col.lower()
        for cand in candidates:
            if cand.lower() in low:
                return col
    return None

def to_numeric_safe(s):
    if pd.isna(s):
        return 0.0
    if isinstance(s, (int, float)):
        return float(s)
    ss = str(s).replace('.', '').replace(',', '.')
    try:
        return float(ss)
    except:
        return 0.0

# ---------- Sidebar --------------------------------------------------
st.sidebar.title("Arquivos e configurações")
csv_upload = st.sidebar.file_uploader("DesVend.CSV", type=['csv'])
pend_upload = st.sidebar.file_uploader("Talões Pendentes (xlsx)", type=['xlsx'])

st.sidebar.markdown("---")
st.sidebar.markdown("**Configurações da premiação**")
prm_text = st.sidebar.text_area(
    "Tabelas de premiação (CSV - Nome,Percentual,ValorFixo)",
    value="Faixa A,5,100\nFaixa B,3,50",
    height=120
)

apply_consultant_filter = st.sidebar.checkbox("Ativar filtro por consultor na aba Faturamento", value=True)

# ---------- Load files ------------------------------------------------
with st.spinner("Carregando arquivos..."):
    df_desvend = read_csv_either(csv_upload, FALLBACK_DESVEND)
    df_taloes = read_excel_either(pend_upload, FALLBACK_TALOES)

if df_desvend.empty:
    st.warning("Arquivo DesVend não encontrado.")
if df_taloes.empty:
    st.warning("Arquivo Talões Pendentes não encontrado.")

# ---------- Detect key columns ---------------------------------------
loja_col = find_column(df_desvend, ['loja', 'codfil', 'filial', 'fil']) if not df_desvend.empty else None
codfil_col = find_column(df_taloes, ['codfil', 'cod_fil', 'filial', 'loja']) if not df_taloes.empty else None

# ---------- Main Layout ----------------------------------------------
st.title("Sistema de Faturamento e Premiação")
tabs = st.tabs(["Faturamento", "Premiação"])

# ------------------- Aba Faturamento ---------------------------------
with tabs[0]:
    st.header("Faturamento")
    if df_desvend.empty or df_taloes.empty:
        st.info("Aguarde upload dos arquivos.")
    else:
        st.markdown(f"**Coluna chave em DesVend:** `{loja_col}` — **Coluna chave em Talões Pendentes:** `{codfil_col}`")
        df_desvend = df_desvend.astype(str)
        df_taloes = df_taloes.astype(str)
        if codfil_col in df_taloes.columns and loja_col in df_desvend.columns:
            valid_lojas = set(df_taloes[codfil_col].dropna().astype(str).unique())
            df_filtered = df_desvend[df_desvend[loja_col].astype(str).isin(valid_lojas)].copy()
        else:
            df_filtered = df_desvend.copy()

        consultant_col = find_column(df_filtered, ['consultor', 'vendedor', 'nome', 'representante'])
        value_col = find_column(df_filtered, ['valor', 'vlr', 'venda', 'total', 'valor_total', 'valor bruto'])

        if apply_consultant_filter and consultant_col:
            consultants = sorted(df_filtered[consultant_col].fillna('N/D').unique())
            selected_consultants = st.multiselect("Filtrar por Consultor", options=consultants, default=consultants)
            df_filtered = df_filtered[df_filtered[consultant_col].isin(selected_consultants)]

        st.subheader("Tabela filtrada — Faturamento")
        st.dataframe(df_filtered, use_container_width=True)
        csv_bytes = df_filtered.to_csv(index=False).encode('utf-8')
        st.download_button("Baixar Faturamento filtrado (CSV)", data=csv_bytes, file_name="faturamento_filtrado.csv", mime='text/csv')

# ------------------- Aba Premiação -----------------------------------
with tabs[1]:
    st.header("Premiação")
    if df_desvend.empty or df_taloes.empty:
        st.info("Aguarde upload dos arquivos.")
    else:
        df_desvend = df_desvend.astype(str)
        df_taloes = df_taloes.astype(str)
        if codfil_col in df_taloes.columns and loja_col in df_desvend.columns:
            valid_lojas = set(df_taloes[codfil_col].dropna().astype(str).unique())
            df_prem = df_desvend[df_desvend[loja_col].astype(str).isin(valid_lojas)].copy()
        else:
            df_prem = df_desvend.copy()

        consultant_col = find_column(df_prem, ['consultor', 'vendedor', 'nome', 'representante'])
        value_col = find_column(df_prem, ['valor', 'vlr', 'venda', 'total', 'valor_total', 'valor bruto'])

        df_prem['__valor_num__'] = df_prem[value_col].apply(to_numeric_safe)
        if consultant_col in df_prem.columns:
            agg = df_prem.groupby(consultant_col)['__valor_num__'].sum().reset_index().rename(columns={'__valor_num__': 'Faturamento_Total'})
        else:
            agg = pd.DataFrame({'Faturamento_Total': [df_prem['__valor_num__'].sum()]})

        st.subheader("Faturamento por consultor")
        st.dataframe(agg, use_container_width=True)

        prm_lines = [r.strip() for r in prm_text.splitlines() if r.strip()]
        prm_rows = []
        for ln in prm_lines:
            parts = [p.strip() for p in ln.split(',')]
            if len(parts) == 3:
                name = parts[0]
                pct = float(parts[1]) if parts[1].replace('.', '', 1).isdigit() else 0.0
                fixed = float(parts[2]) if parts[2].replace('.', '', 1).isdigit() else 0.0
                prm_rows.append({'Faixa': name, 'Percentual': pct, 'ValorFixo': fixed})

        prm_df = pd.DataFrame(prm_rows)
        edited_prm = st.experimental_data_editor(prm_df, num_rows="dynamic")

        result = agg.copy()
        for idx, row in edited_prm.iterrows():
            pct = float(row.get('Percentual', 0.0))
            fixed = float(row.get('ValorFixo', 0.0))
            col_name = f"Bônus_{row.get('Faixa', idx)}"
            result[col_name] = result['Faturamento_Total'] * (pct / 100.0) + fixed
        result['Bônus_Total'] = result.filter(like='Bônus_').sum(axis=1)

        st.subheader("Resultado da premiação")
        st.dataframe(result, use_container_width=True)
        csv_out = result.to_csv(index=False).encode('utf-8')
        st.download_button("Baixar relatório de premiação (CSV)", data=csv_out, file_name="premiacao_relatorio.csv", mime='text/csv')
