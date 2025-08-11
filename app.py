# app.py
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Relatório de Faturamento & Premiações", layout="wide")
st.title("📊 Relatório de Faturamento & Premiações")

# -------------------------
# Helpers: format display
# -------------------------
def format_moeda_str(x):
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return ""

def format_percent_str(x):
    try:
        return f"{x:.1%}".replace(".", ",")
    except:
        return ""

def df_for_display(df, moeda_cols, percent_cols, cols_order=None):
    """Return a copy formatted as strings for display in Streamlit."""
    df_disp = df.copy()
    if cols_order:
        df_disp = df_disp[cols_order]
    for c in moeda_cols:
        if c in df_disp.columns:
            df_disp[c] = df_disp[c].apply(lambda v: format_moeda_str(v) if pd.notnull(v) else "")
    for c in percent_cols:
        if c in df_disp.columns:
            df_disp[c] = df_disp[c].apply(lambda v: format_percent_str(v) if pd.notnull(v) else "")
    return df_disp

# -------------------------
# Excel export with formatting using openpyxl
# -------------------------
def to_excel_bytes_with_formats(df, sheet_name="Sheet1", moeda_cols=None, percent_cols=None):
    """Return bytes of an Excel file where numeric columns are formatted (R$ and %) and column widths auto."""
    if moeda_cols is None:
        moeda_cols = []
    if percent_cols is None:
        percent_cols = []

    buffer = BytesIO()
    # write using pandas openpyxl engine (creates basic workbook)
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb[sheet_name]

    # Header styling: gray background, bold
    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    # Shade LOJA column (first column) light gray for data rows
    gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    max_row = ws.max_row
    # find LOJA column letter (assume first column header is LOJA)
    # We'll find header index with name 'LOJA' to be safe
    loja_col_idx = None
    for idx, cell in enumerate(ws[1], start=1):
        if str(cell.value).strip().upper() == "LOJA":
            loja_col_idx = idx
            break
    if loja_col_idx:
        loja_col_letter = get_column_letter(loja_col_idx)
        for r in range(2, max_row + 1):
            ws[f"{loja_col_letter}{r}"].fill = gray_fill

    # Apply number formats for moeda and percent columns
    col_idx_by_name = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}
    for col_name in moeda_cols:
        if col_name in col_idx_by_name:
            col_idx = col_idx_by_name[col_name]
            fmt = 'R$ #,##0.00'
            for r in range(2, max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                # keep blank cells blank
                if cell.value is not None and cell.value != "":
                    try:
                        cell.number_format = fmt
                    except:
                        pass

    for col_name in percent_cols:
        if col_name in col_idx_by_name:
            col_idx = col_idx_by_name[col_name]
            fmt = '0.0%'
            for r in range(2, max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                if cell.value is not None and cell.value != "":
                    try:
                        cell.number_format = fmt
                    except:
                        pass

    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value is not None:
                    l = len(str(cell.value))
                    if l > max_length:
                        max_length = l
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# -------------------------
# Faturamento processing
# -------------------------
def process_faturamento(df_desvend):
    # Expected columns in DesVend: at least VENDEDOR, LOJA, COTA TOTAL, TOTAL VENDAS, QUANT VENDAS, SALDO COTA TOTAL, TICK MEDIO (optional)
    required = ["VENDEDOR", "LOJA", "COTA TOTAL", "TOTAL VENDAS", "QUANT VENDAS", "SALDO COTA TOTAL"]
    for c in required:
        if c not in df_desvend.columns:
            raise ValueError(f"Coluna obrigatória ausente em DesVend: {c}")

    # Group by LOJA + VENDEDOR
    grp = df_desvend.groupby(["LOJA", "VENDEDOR"], as_index=False).agg({
        "COTA TOTAL": "sum",
        "TOTAL VENDAS": "sum",
        "QUANT VENDAS": "sum",
        "SALDO COTA TOTAL": "sum"
    })

    # % VENDAS per vendor-row
    grp["% VENDAS"] = grp["TOTAL VENDAS"] / grp["COTA TOTAL"]
    # TICK MEDIO as TOTAL VENDAS / QUANT VENDAS (avoid div by zero)
    grp["TICK MEDIO"] = grp.apply(lambda r: (r["TOTAL VENDAS"] / r["QUANT VENDAS"]) if r["QUANT VENDAS"] else 0.0, axis=1)
    # % SALDO COTA
    grp["% SALDO COTA"] = grp["SALDO COTA TOTAL"] / grp["COTA TOTAL"]

    # Reorder columns:
    ordered = ["LOJA", "VENDEDOR", "COTA TOTAL", "TOTAL VENDAS", "QUANT VENDAS",
               "% VENDAS", "TICK MEDIO", "SALDO COTA TOTAL", "% SALDO COTA"]
    grp = grp[ordered]
    return grp

# -------------------------
# Premiações processing
# -------------------------
def process_premiacoes(df_desvend, df_taloes, pct_threshold, valor_premio, match_inner=True):
    # Validate columns
    if "VENDEDOR" not in df_taloes.columns:
        raise ValueError("Coluna 'VENDEDOR' ausente em Talões Pendentes (é necessária para o match).")

    # Ensure numeric columns exist in taloes (if absent, create zeros)
    for col in ["VENDAS FORA DA POLÍTICA", "VENDAS ATUALIZADAS", "% VENDAS ATUALIZADAS"]:
        if col not in df_taloes.columns:
            df_taloes[col] = 0.0

    # Merge by VENDEDOR. Default inner to consider only matching vendors; allow outer if requested
    how = "inner" if match_inner else "left"
    merged = pd.merge(df_desvend, df_taloes[["VENDEDOR", "VENDAS FORA DA POLÍTICA", "VENDAS ATUALIZADAS", "% VENDAS ATUALIZADAS"]],
                      on="VENDEDOR", how=how, suffixes=("", "_TALAO"))

    # If VENDAS_ATUALIZADAS missing, derive as TOTAL VENDAS - VENDAS FORA...
    merged["VENDAS FORA DA POLÍTICA"] = merged["VENDAS FORA DA POLÍTICA"].fillna(0.0)
    merged["VENDAS ATUALIZADAS"] = merged.apply(
        lambda r: r["VENDAS ATUALIZADAS"] if pd.notnull(r["VENDAS ATUALIZADAS"]) and r["VENDAS ATUALIZADAS"] != 0 else r["TOTAL VENDAS"] - r["VENDAS FORA DA POLÍTICA"],
        axis=1
    )
    # Percent updated
    merged["% VENDAS ATUALIZADAS"] = merged.apply(
        lambda r: (r["VENDAS ATUALIZADAS"] / r["COTA TOTAL"]) if r["COTA TOTAL"] else 0.0,
        axis=1
    )

    # Determine premiado per VENDEDOR
    merged["PREMIADO"] = merged["% VENDAS ATUALIZADAS"].apply(lambda v: "SIM" if v >= pct_threshold else "NÃO")
    merged["VALOR"] = merged["PREMIADO"].apply(lambda s: valor_premio if s == "SIM" else 0.0)

    # Now aggregate by LOJA to produce the premiações sheet
    agg = merged.groupby("LOJA", as_index=False).agg({
        "COTA TOTAL": "sum",
        "TOTAL VENDAS": "sum",
        "VENDAS FORA DA POLÍTICA": "sum",
        "VENDAS ATUALIZADAS": "sum",
        "% VENDAS ATUALIZADAS": "mean",
        "SALDO COTA TOTAL": "sum",
        "VALOR": "sum"
    })

    # Derived fields
    agg["% VENDAS"] = agg.apply(lambda r: (r["TOTAL VENDAS"] / r["COTA TOTAL"]) if r["COTA TOTAL"] else 0.0, axis=1)
    agg["% SALDO COTA"] = agg.apply(lambda r: (r["SALDO COTA TOTAL"] / r["COTA TOTAL"]) if r["COTA TOTAL"] else 0.0, axis=1)
    # TOTAL LOJA = TOTAL VENDAS + VALOR (keeps numeric)
    agg["TOTAL LOJA"] = agg["TOTAL VENDAS"] + agg["VALOR"]

    # Reorder columns as requested
    cols_order = ["LOJA", "COTA TOTAL", "TOTAL VENDAS", "% VENDAS", "SALDO COTA TOTAL",
                  "% SALDO COTA", "VENDAS FORA DA POLÍTICA", "VENDAS ATUALIZADAS",
                  "% VENDAS ATUALIZADAS", "VALOR", "TOTAL LOJA"]
    # keep only columns that exist
    cols_order = [c for c in cols_order if c in agg.columns]
    agg = agg[cols_order]

    return merged, agg  # return both merged per-vendedor and aggregated per-loja

# -------------------------
# UI: tabs and controls
# -------------------------
tab1, tab2 = st.tabs(["Faturamento", "Premiações"])

with tab1:
    st.header("Faturamento")
    file_desvend = st.file_uploader("Envie o arquivo DesVend (xlsx) — aba 'DesVend' será usada", type=["xlsx"], key="u1")
    if file_desvend:
        try:
            df_desvend = pd.read_excel(file_desvend, sheet_name="DesVend")
        except Exception:
            # fallback: try reading first sheet
            df_desvend = pd.read_excel(file_desvend)

        try:
            df_fatur = process_faturamento(df_desvend)
        except Exception as e:
            st.error(f"Erro processando DesVend: {e}")
            st.stop()

        # Prepare display with formatting and ordering
        moeda_cols = ["COTA TOTAL", "TOTAL VENDAS", "TICK MEDIO", "SALDO COTA TOTAL"]
        percent_cols = ["% VENDAS", "% SALDO COTA"]
        display_order = ["LOJA", "VENDEDOR", "COTA TOTAL", "TOTAL VENDAS", "QUANT VENDAS",
                         "% VENDAS", "TICK MEDIO", "SALDO COTA TOTAL", "% SALDO COTA"]

        df_disp = df_for_display(df_fatur, moeda_cols, percent_cols, cols_order=display_order)

        st.subheader("Tabela consolidada (por LOJA e VENDEDOR)")
        st.dataframe(df_disp, use_container_width=True)

        # Graph: aggregate by LOJA (sum total vendas) for clearer visualization
        agg_loja = df_fatur.groupby("LOJA", as_index=False).agg({"TOTAL VENDAS": "sum"}).sort_values("TOTAL VENDAS", ascending=False)
        n_lojas = agg_loja.shape[0]
        # dynamic height:  max(400, 40 * n_lojas)
        height = max(400, int(40 * n_lojas))
        fig = px.bar(agg_loja, x="LOJA", y="TOTAL VENDAS", text="TOTAL VENDAS",
                     title="Total Vendas por Loja (agregado)",
                     labels={"TOTAL VENDAS": "Total Vendas (R$)"})
        fig.update_traces(texttemplate='R$ %{y:,.2f}', textposition='outside')
        fig.update_layout(xaxis_tickangle=-45, height=height)
        st.plotly_chart(fig, use_container_width=True)

        # Download consolidated excel (numbers preserved, formatting applied in file)
        excel_bytes = to_excel_bytes_with_formats(df_fatur, sheet_name="Faturamento",
                                                  moeda_cols=moeda_cols, percent_cols=percent_cols)
        st.download_button("📥 Baixar Excel - Faturamento", data=excel_bytes,
                           file_name="faturamento_consolidado.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab2:
    st.header("Premiações")
    st.info("Envie DesVend e Talões Pendentes. Match será feito por VENDEDOR; depois agregamos por LOJA.")
    file_desvend2 = st.file_uploader("Arquivo DesVend (xlsx)", type=["xlsx"], key="u2")
    file_taloes = st.file_uploader("Arquivo Talões Pendentes (xlsx)", type=["xlsx"], key="u3")

    mode = st.radio("Modo de premiação", options=["Automático", "Manual"], index=0)

    pct_threshold_input = st.number_input("Percentual mínimo para premiar (ex: 45 = 45%)", min_value=0.0, max_value=100.0, value=45.0, step=0.5)
    pct_threshold = pct_threshold_input / 100.0
    valor_premio = st.number_input("Valor fixo da premiação (R$)", min_value=0.0, value=100.0, step=1.0)

    if file_desvend2 and file_taloes:
        try:
            try:
                df_desvend_all = pd.read_excel(file_desvend2, sheet_name="DesVend")
            except Exception:
                df_desvend_all = pd.read_excel(file_desvend2)
            try:
                df_taloes = pd.read_excel(file_taloes)
            except Exception:
                df_taloes = pd.read_excel(file_taloes)
        except Exception as e:
            st.error(f"Erro ao ler arquivos: {e}")
            st.stop()

        # Process premiações
        merged_vendedor, df_prem_agg = process_premiacoes(df_desvend_all, df_taloes, pct_threshold, valor_premio, match_inner=True)

        # If manual mode, show per-vendedor controls to override premiado and value
        if mode == "Manual":
            st.write("Modo Manual — ajuste por vendedor abaixo (se quiser).")
            # Build DataFrame of vendors (merged_vendedor)
            manual_df = merged_vendedor[["LOJA", "VENDEDOR", "COTA TOTAL", "TOTAL VENDAS", "VENDAS FORA DA POLÍTICA", "VENDAS ATUALIZADAS", "% VENDAS ATUALIZADAS", "PREMIADO", "VALOR"]].copy()
            manual_df = manual_df.reset_index(drop=True)
            # Create editable controls per vendor (could be many; show a subset or paginated if necessary)
            override_rows = []
            for i, row in manual_df.iterrows():
                cols = st.columns([3, 2, 2, 2])
                with cols[0]:
                    st.markdown(f"**Loja:** {row['LOJA']} — **Vendedor:** {row['VENDEDOR']}")
                with cols[1]:
                    premiado = st.selectbox(f"Premiado? [{i}]", options=["SIM", "NÃO"], index=0 if row["PREMIADO"]=="SIM" else 1, key=f"p_{i}")
                with cols[2]:
                    valor = st.number_input(f"Valor [{i}]", min_value=0.0, value=float(row["VALOR"]), key=f"v_{i}")
                override_rows.append({"index": i, "PREMIADO": premiado, "VALOR": valor})
            # Apply overrides
            for r in override_rows:
                idx = r["index"]
                merged_vendedor.at[idx, "PREMIADO"] = r["PREMIADO"]
                merged_vendedor.at[idx, "VALOR"] = r["VALOR"]

            # Re-aggregate by loja after manual edits
            df_prem_agg = merged_vendedor.groupby("LOJA", as_index=False).agg({
                "COTA TOTAL": "sum",
                "TOTAL VENDAS": "sum",
                "VENDAS FORA DA POLÍTICA": "sum",
                "VENDAS ATUALIZADAS": "sum",
                "% VENDAS ATUALIZADAS": "mean",
                "SALDO COTA TOTAL": "sum",
                "VALOR": "sum"
            })
            df_prem_agg["% VENDAS"] = df_prem_agg.apply(lambda r: (r["TOTAL VENDAS"]/r["COTA TOTAL"]) if r["COTA TOTAL"] else 0.0, axis=1)
            df_prem_agg["% SALDO COTA"] = df_prem_agg.apply(lambda r: (r["SALDO COTA TOTAL"]/r["COTA TOTAL"]) if r["COTA TOTAL"] else 0.0, axis=1)
            df_prem_agg["TOTAL LOJA"] = df_prem_agg["TOTAL VENDAS"] + df_prem_agg["VALOR"]

        # Prepare display and Excel
        moeda_cols_prem = ["COTA TOTAL", "TOTAL VENDAS", "SALDO COTA TOTAL", "VENDAS FORA DA POLÍTICA", "VENDAS ATUALIZADAS", "VALOR", "TOTAL LOJA"]
        percent_cols_prem = ["% VENDAS", "% SALDO COTA", "% VENDAS ATUALIZADAS"]

        # Display premiações aggregated by LOJA
        display_order_prem = ["LOJA", "COTA TOTAL", "TOTAL VENDAS", "% VENDAS", "SALDO COTA TOTAL", "% SALDO COTA",
                              "VENDAS FORA DA POLÍTICA", "VENDAS ATUALIZADAS", "% VENDAS ATUALIZADAS", "VALOR", "TOTAL LOJA"]
        df_prem_disp = df_for_display(df_prem_agg, moeda_cols_prem, percent_cols_prem, cols_order=display_order_prem)

        st.subheader("Premiações (agregado por LOJA)")
        st.dataframe(df_prem_disp, use_container_width=True)

        # Plot optional: show TOTAL LOJA or TOTAL VENDAS per loja
        fig2 = px.bar(df_prem_agg.sort_values("TOTAL LOJA", ascending=False),
                      x="LOJA", y="TOTAL LOJA", text="TOTAL LOJA",
                      title="Total por Loja (Vendas + Premiações)")
        fig2.update_traces(texttemplate='R$ %{y:,.2f}', textposition='outside')
        fig2.update_layout(xaxis_tickangle=-45, height=max(400, 40 * df_prem_agg.shape[0]))
        st.plotly_chart(fig2, use_container_width=True)

        # Excel export for premiações
        excel_prem = to_excel_bytes_with_formats(df_prem_agg, sheet_name="Premiações",
                                                moeda_cols=moeda_cols_prem, percent_cols=percent_cols_prem)
        st.download_button("📥 Baixar Excel - Premiações", data=excel_prem,
                           file_name="premiacoes_consolidado.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Envie ambos os arquivos: DesVend e Talões Pendentes para calcular premiações.")
