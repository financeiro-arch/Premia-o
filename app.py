import streamlit as st
import pandas as pd
import numpy as np
import io
from pathlib import Path

st.set_page_config(page_title="Sistema Faturamento & Premiação", layout="wide")

# ---------- Helpers -------------------------------------------------

FALLBACK_DESVEND = "/mnt/data/DesVend.CSV"
FALLBACK_TALOES = "/mnt/data/TALÕES PENDENTES.xlsx"
FALLBACK_AUD = "/mnt/data/DesVend AUDITORIA_AUTOMATICA.xlsx"

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
    # try fuzzy contains
    for col in df.columns:
        low = col.lower()
        for cand in candidates:
            if cand.lower() in low:
                return col
    return None


def to_numeric_safe(s):
    # remove common thousands separators and convert
    if pd.isna(s):
        return 0.0
    if isinstance(s, (int, float)):
        return float(s)
    ss = str(s).replace('.', '').replace(',', '.')
    try:
        return float(ss)
    except:
        return 0.0

# ---------- Layout --------------------------------------------------

st.sidebar.title("Arquivos e configurações")
st.sidebar.markdown("Faça upload dos arquivos ou deixe que o sistema tente carregar os padrões em `/mnt/data`.")
csv_upload = st.sidebar.file_uploader("DesVend.CSV", type=['csv'])
pend_upload = st.sidebar.file_uploader("Talões Pendentes (xlsx)", type=['xlsx'])
aud_upload = st.sidebar.file_uploader("DesVend AUDITORIA_AUTOMATICA (xlsx)", type=['xlsx'])

st.sidebar.markdown("---")
st.sidebar.markdown("**Configurações da premiação**")
st.sidebar.info("No editor abaixo informe linhas no formato: Nome,Percentual(%),ValorFixo\nEx: Faixa A,5,100\nIsso será aplicado como: bônus = faturamento_total * (Percentual/100) + ValorFixo")
prm_text = st.sidebar.text_area("Tabelas de premiação (CSV - Nome,Percentual,ValorFixo)", value="Faixa A,5,100\nFaixa B,3,50", height=120)

apply_consultant_filter = st.sidebar.checkbox("Ativar filtro por consultor na aba Faturamento", value=True)

# Load files
with st.spinner("Carregando arquivos..."):
    df_desvend = read_csv_either(csv_upload, FALLBACK_DESVEND)
    df_taloes = read_excel_either(pend_upload, FALLBACK_TALOES)
    df_aud = read_excel_either(aud_upload, FALLBACK_AUD)

# Validate basic
if df_desvend.empty:
    st.warning("Arquivo DesVend não encontrado — faça upload ou coloque em /mnt/data/DesVend.CSV")

if df_taloes.empty:
    st.warning("Arquivo Talões Pendentes não encontrado — faça upload ou coloque em /mnt/data/TALÕES PENDENTES.xlsx")

# Find key columns
loja_col = find_column(df_desvend, ['loja', 'codfil', 'filial', 'fil']) if not df_desvend.empty else None
codfil_col = find_column(df_taloes, ['codfil', 'cod_fil', 'filial', 'loja']) if not df_taloes.empty else None

# Normalize column names check
st.title("Sistema de Faturamento e Premiação")
st.markdown("Sistema criado em Streamlit — abas: Faturamento | Premiação")

tabs = st.tabs(["Faturamento", "Premiação"])

# ------------------- Aba Faturamento ---------------------------------
with tabs[0]:
    st.header("Faturamento")

    if df_desvend.empty or df_taloes.empty:
        st.info("Aguarde upload dos arquivos ou confira os caminhos em /mnt/data.")
    else:
        st.markdown(f"**Coluna chave em DesVend:** `{loja_col}`  —  **Coluna chave em Talões Pendentes:** `{codfil_col}`")

        # Ensure columns are strings
        df_desvend = df_desvend.astype(str)
        df_taloes = df_taloes.astype(str)

        # Build set of valid lojas
        if codfil_col in df_taloes.columns and loja_col in df_desvend.columns:
            valid_lojas = set(df_taloes[codfil_col].dropna().astype(str).unique())
            df_filtered = df_desvend[df_desvend[loja_col].astype(str).isin(valid_lojas)].copy()
        else:
            df_filtered = df_desvend.copy()

        # Detect consultant and value columns
        consultant_col = find_column(df_filtered, ['consultor', 'vendedor', 'nome', 'representante'])
        value_col = find_column(df_filtered, ['valor', 'vlr', 'venda', 'total', 'valor_total', 'valor bruto'])

        st.markdown("**Colunas detectadas**")
        st.write({
            'consultant_col': consultant_col,
            'value_col': value_col
        })

        # Filters
        if apply_consultant_filter and consultant_col:
            consultants = sorted(df_filtered[consultant_col].fillna('N/D').unique())
            selected_consultants = st.multiselect("Filtrar por Consultor", options=consultants, default=consultants)
            df_filtered = df_filtered[df_filtered[consultant_col].isin(selected_consultants)]

        st.subheader("Tabela filtrada — Faturamento")
        st.dataframe(df_filtered, use_container_width=True)

        # Allow download of filtered faturamento
        csv_bytes = df_filtered.to_csv(index=False).encode('utf-8')
        st.download_button("Baixar Faturamento filtrado (CSV)", data=csv_bytes, file_name="faturamento_filtrado.csv", mime='text/csv')

# ------------------- Aba Premiação ----------------------------------
with tabs[1]:
    st.header("Premiação")

    if df_desvend.empty or df_taloes.empty:
        st.info("Aguarde upload dos arquivos ou confira os caminhos em /mnt/data.")
    else:
        # use same filtering by loja
        df_desvend = df_desvend.astype(str)
        df_taloes = df_taloes.astype(str)
        if codfil_col in df_taloes.columns and loja_col in df_desvend.columns:
            valid_lojas = set(df_taloes[codfil_col].dropna().astype(str).unique())
            df_prem = df_desvend[df_desvend[loja_col].astype(str).isin(valid_lojas)].copy()
        else:
            df_prem = df_desvend.copy()

        # If AUD model provided, try to extract any mapping or important columns
        st.markdown("Você pode fornecer uma planilha `DesVend AUDITORIA_AUTOMATICA.xlsx` para que possamos tentar replicar fórmulas automaticamente. Se não houver, o sistema usará uma regra padrão (percentual + valor fixo).")
        if not df_aud.empty:
            st.success("Planilha de auditoria detectada — usaremos como referência quando possível.")

        consultant_col = find_column(df_prem, ['consultor', 'vendedor', 'nome', 'representante'])
        value_col = find_column(df_prem, ['valor', 'vlr', 'venda', 'total', 'valor_total', 'valor bruto'])

        if consultant_col is None or value_col is None:
            st.error("Não foi possível detectar automaticamente as colunas de consultor ou valor. Verifique os nomes das colunas nos arquivos ou informe manualmente abaixo.")
            consultant_col = st.text_input("Nome da coluna de consultor (insira exatamente como no arquivo)")
            value_col = st.text_input("Nome da coluna de valor")

        # Convert values
        df_prem['__valor_num__'] = df_prem[value_col].apply(to_numeric_safe)

        # Aggregate faturamento por consultor
        if consultant_col in df_prem.columns:
            agg = df_prem.groupby(consultant_col)['__valor_num__'].sum().reset_index().rename(columns={'__valor_num__': 'Faturamento_Total'})
        else:
            agg = pd.DataFrame({'Faturamento_Total': [df_prem['__valor_num__'].sum()]})

        st.subheader("Faturamento por consultor (base para premiação)")
        st.dataframe(agg, use_container_width=True)

        # Parse premiação input
        prm_lines = [r.strip() for r in prm_text.splitlines() if r.strip()]
        prm_rows = []
        for ln in prm_lines:
            parts = [p.strip() for p in ln.split(',')]
            if len(parts) == 3:
                name = parts[0]
                try:
                    pct = float(parts[1])
                except:
                    pct = 0.0
                try:
                    fixed = float(parts[2])
                except:
                    fixed = 0.0
                prm_rows.append({'Faixa': name, 'Percentual': pct, 'ValorFixo': fixed})
        if len(prm_rows) == 0:
            st.warning("Nenhuma linha de premiação válida detectada no editor. Use o formato: Nome,Percentual(%),ValorFixo")

        prm_df = pd.DataFrame(prm_rows)
        st.subheader("Regras de premiação (revise/edite)")
        edited_prm = st.experimental_data_editor(prm_df, num_rows="dynamic") if not prm_df.empty else st.experimental_data_editor(pd.DataFrame(columns=['Faixa','Percentual','ValorFixo']))

        # Choose rule application: apply single rule to all? or allow split by ranking?
        st.markdown("**Como aplicar as regras?**")
        mode = st.radio("Aplicação das regras:", options=["Aplicar todas as regras a todos os consultores (soma de bônus)", "Atribuir uma única faixa por consultor via regra de ranking"], index=0)

        result = agg.copy()
        if 'Faturamento_Total' not in result.columns:
            result['Faturamento_Total'] = 0.0

        if mode.startswith("Aplicar todas"):
            # For each rule compute bonus and sum
            total_bonus_cols = []
            for idx, row in edited_prm.iterrows():
                pct = float(row.get('Percentual', 0.0)) if not pd.isna(row.get('Percentual', 0.0)) else 0.0
                fixed = float(row.get('ValorFixo', 0.0)) if not pd.isna(row.get('ValorFixo', 0.0)) else 0.0
                col_name = f"Bônus_{row.get('Faixa', idx)}"
                result[col_name] = result['Faturamento_Total'] * (pct/100.0) + fixed
                total_bonus_cols.append(col_name)
            result['Bônus_Total'] = result[total_bonus_cols].sum(axis=1) if len(total_bonus_cols)>0 else 0.0
        else:
            # Ranking mode: ask for thresholds by faturamento
            st.info("No modo ranking, defina faixas pela interface: ex: >=10000 => Faixa A. As faixas serão aplicadas por faturamento total.")
            # we'll reuse edited_prm rows as ordered mapping: if Total >= previous threshold assign rule.
            thresholds = st.text_area("Defina thresholds (uma linha por faixa): FaturamentoMin,FaixaNome\nEx:\n10000,Faixa A\n5000,Faixa B", value="10000,Faixa A\n5000,Faixa B")
            thresh_rows = [r.strip() for r in thresholds.splitlines() if r.strip()]
            thresh_map = []
            for ln in thresh_rows:
                parts = [p.strip() for p in ln.split(',')]
                if len(parts) >= 2:
                    try:
                        tv = float(parts[0])
                    except:
                        tv = 0.0
                    fname = parts[1]
                    thresh_map.append((tv, fname))
            # sort descending
            thresh_map = sorted(thresh_map, key=lambda x: -x[0])

            # map faixa to rule
            faixa_to_rule = {}
            for idx, row in edited_prm.iterrows():
                faixa_to_rule[str(row.get('Faixa'))] = (float(row.get('Percentual',0.0)), float(row.get('ValorFixo',0.0)))

            def assign_faixa(val):
                for thresh, fname in thresh_map:
                    if val >= thresh:
                        return fname
                return None

            result['Faixa'] = result['Faturamento_Total'].apply(assign_faixa)
            result['Faixa'] = result['Faixa'].fillna('Sem Faixa')
            # compute bonus based on the mapping
            def compute_bonus(row):
                f = row['Faixa']
                if f in faixa_to_rule:
                    pct, fx = faixa_to_rule[f]
                    return row['Faturamento_Total']*(pct/100.0) + fx
                return 0.0
            result['Bônus_Total'] = result.apply(compute_bonus, axis=1)

        st.subheader("Resultado da premiação")
        st.dataframe(result, use_container_width=True)

        # Download results
        csv_out = result.to_csv(index=False).encode('utf-8')
        st.download_button("Baixar relatório de premiação (CSV)", data=csv_out, file_name="premiacao_relatorio.csv", mime='text/csv')

st.sidebar.markdown("---")
st.sidebar.markdown("Desenvolvido para integração com GitHub e deploy no Streamlit Cloud. \n\nObservações: revise as colunas detectadas e as regras de premiação antes de gerar o relatório.")

# EOF
