"""
Energy Intelligence — Unidad Minera
Dashboard Streamlit conectado a Proyeccion_Energia_v5_FINAL.xlsx
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Energy Intelligence — Unidad Minera",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Colores por planta (consistentes en todo el dashboard)
COLORS = {
    "Sulfuros":          "#185FA5",
    "Óxidos sin Elec.":  "#ED7D31",
    "Electrodeposición": "#C55A11",
    "Óxidos Total":      "#2E75B6",
    "Infraestructura":   "#7030A0",
    "Total":             "#1F3864",
}

# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data(path: str):
    xl = pd.ExcelFile(path)

    # --- Histórico ---
    hist = pd.read_excel(xl, sheet_name="Datos Históricos", header=1)
    hist.columns = [c.replace("\n", " ").strip() for c in hist.columns]
    hist = hist.dropna(subset=["Fecha"])
    hist["Fecha"] = pd.to_datetime(hist["Fecha"])
    hist = hist[hist["Fecha"] >= "2022-01-01"].copy()

    # --- Proyección ---
    proj = pd.read_excel(xl, sheet_name="Proyección Energía", header=3)
    proj.columns = [c.replace("\n", " ").strip() for c in proj.columns]
    proj = proj.dropna(subset=["Período"])

    # Fecha de representación
    def parse_period(lbl):
        s = str(lbl)
        if "m" in s:
            yr = int("20" + s[:2]); m = int(s.split("m")[1])
            return pd.Timestamp(yr, m, 1)
        elif "q" in s.lower() or "Q" in s:
            yr = int("20" + s[:2]); q = int(s[-1])
            return pd.Timestamp(yr, 3*q - 2, 1)
        else:
            try: return pd.Timestamp(int(float(s)), 1, 1)
            except: return pd.NaT

    proj["Fecha"] = proj["Período"].apply(parse_period)
    proj = proj.dropna(subset=["Fecha"])

    # --- Resumen Anual ---
    ann = pd.read_excel(xl, sheet_name="Resumen Anual", header=1)
    ann.columns = [c.replace("\n", " ").strip() for c in ann.columns]
    ann = ann.dropna(subset=["Año"])
    ann["Año"] = ann["Año"].astype(int)

    return hist, proj, ann

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚡ Energy Intelligence")
    st.caption("Unidad Minera — LOM25 Óptimo")
    st.divider()

    EXCEL_PATH = st.text_input(
        "Ruta del archivo Excel",
        value="Proyeccion_Energia_v5_FINAL.xlsx",
        help="Sube el archivo o indica la ruta relativa",
    )

    uploaded = st.file_uploader(
        "O sube el Excel aquí", type=["xlsx"],
        label_visibility="collapsed"
    )

    st.divider()
    st.markdown("### Filtros")

    vista = st.radio(
        "Vista temporal",
        ["Histórico (2022–2026)", "Proyectado (2026–2045)", "Histórico + Proyectado"],
        index=2,
    )

    plantas_opciones = ["Todas", "Sulfuros", "Óxidos", "Infraestructura"]
    planta_sel = st.selectbox("Planta", plantas_opciones)

    st.divider()
    st.markdown("### Rango de fechas")

    MESES_ES = {1:"Ene", 2:"Feb", 3:"Mar", 4:"Abr", 5:"May", 6:"Jun",
                7:"Jul", 8:"Ago", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dic"}

    # Rango fijo 2022-2026 (años disponibles en histórico)
    anos_hist = list(range(2022, 2027))
    ano_inicio = st.selectbox("Año inicio", anos_hist, index=0)
    mes_inicio = st.selectbox(
        "Mes inicio",
        list(range(1, 13)),
        format_func=lambda m: MESES_ES[m],
        index=0
    )
    ano_fin = st.selectbox("Año fin", anos_hist, index=len(anos_hist)-1)
    mes_fin = st.selectbox(
        "Mes fin",
        list(range(1, 13)),
        format_func=lambda m: MESES_ES[m],
        index=11
    )
    fecha_inicio = pd.Timestamp(ano_inicio, mes_inicio, 1)
    fecha_fin    = pd.Timestamp(ano_fin, mes_fin, 1) + pd.offsets.MonthEnd(0)

    st.divider()
    if st.button("🔄 Recargar datos"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("### Acerca del modelo")
    st.caption(
        "Metodología: Tendencia Anual OLS + "
        "Estacionalidad Mensual | Modelo v5\n\n"
        "Correcciones aplicadas:\n"
        "- Floor Óxidos = nivel 2025-26\n"
        "- Infra: media fija (sin falsa tendencia)\n"
        "- Sulfuros: base últimos 15m"
    )

# ─────────────────────────────────────────────
# CARGAR DATOS
# ─────────────────────────────────────────────
try:
    if uploaded:
        hist, proj, ann = load_data(uploaded)
    else:
        hist, proj, ann = load_data(EXCEL_PATH)
    data_ok = True
except Exception as e:
    st.error(f"No se pudo cargar el archivo: {e}")
    st.info("Sube el archivo **Proyeccion_Energia_v5_FINAL.xlsx** usando el panel lateral.")
    data_ok = False

if not data_ok:
    st.stop()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title("⚡ Consumo Energético — Plantas de Proceso")
    st.caption(
        f"LOM25 Óptimo · Sulfuros · Óxidos · Infraestructura · "
        f"Histórico 2022–2026 + Proyección 2026–2045"
    )
with col_h2:
    ultimo_hist = hist["Fecha"].max().strftime("%b-%Y")
    st.metric("Último dato real", ultimo_hist)

st.divider()

# ─────────────────────────────────────────────
# KPIs — ÚLTIMO MES HISTÓRICO
# ─────────────────────────────────────────────
ult = hist.iloc[-1]
pen = hist.iloc[-2]

# Aplicar filtro de fechas al histórico
hist_filtered = hist[
    (hist["Fecha"] >= fecha_inicio) &
    (hist["Fecha"] <= fecha_fin)
].copy()

# Si el filtro deja menos de 2 filas, usar histórico completo
if len(hist_filtered) < 2:
    hist_filtered = hist.copy()
    st.sidebar.warning("Rango muy estrecho — mostrando histórico completo")

ult = hist_filtered.iloc[-1]
pen = hist_filtered.iloc[-2] if len(hist_filtered) >= 2 else hist_filtered.iloc[-1]

def delta_pct(a, b):
    return f"{((a/b)-1)*100:+.1f}%" if b > 0 else "—"

# ── FILA 1: Consumos GWh ──────────────────────────────────
st.markdown("#### Consumos energéticos — último mes real")
cols_gwh = st.columns(4)
gwh_data = [
    ("⚡ Total mensual",
     ult["Sulfuros Total (kWh)"] + ult["Óxidos Total (kWh)"] + ult["Infra Bombeo (kWh)"],
     pen["Sulfuros Total (kWh)"] + pen["Óxidos Total (kWh)"] + pen["Infra Bombeo (kWh)"]),
    ("🟦 Sulfuros",
     ult["Sulfuros Total (kWh)"], pen["Sulfuros Total (kWh)"]),
    ("🟧 Óxidos Total",
     ult["Óxidos Total (kWh)"],   pen["Óxidos Total (kWh)"]),
    ("🟪 Infraestructura",
     ult["Infra Bombeo (kWh)"],   pen["Infra Bombeo (kWh)"]),
]
for col, (lbl, val, val_prev) in zip(cols_gwh, gwh_data):
    col.metric(lbl, f"{val/1e6:,.3f} GWh", delta=delta_pct(val, val_prev))

st.markdown("")

# ── FILA 2: Ratios kWh/driver ─────────────────────────────
st.markdown("#### Ratios de consumo unitario — último mes real")
cols_rat = st.columns(4)

# Ratio Ox-EW — leer desde hoja Proyección Energía (col 16)
# Usar primer período proyectado (abr-2026) como referencia
r_elec_ult = float(proj["Ratio Ox-EW TMF (kWh/t)"].iloc[0]) if "Ratio Ox-EW TMF (kWh/t)" in proj.columns else 0
r_elec_pen = float(proj["Ratio Ox-EW TMF (kWh/t)"].iloc[1]) if "Ratio Ox-EW TMF (kWh/t)" in proj.columns and len(proj) > 1 else r_elec_ult

# Ratio Infra — calcular desde kWh y m3
r_infra_ult = (float(ult["Infra Bombeo (kWh)"]) / float(ult["Agua Mar (m3)"])
               if float(ult.get("Agua Mar (m3)", 0)) > 0 else 0)
r_infra_pen = (float(pen["Infra Bombeo (kWh)"]) / float(pen["Agua Mar (m3)"])
               if float(pen.get("Agua Mar (m3)", 0)) > 0 else 0)

ratio_data = [
    ("Ratio Sulf\n(kWh/t TMS)",
     ult["Ratio Sulf kWh/t TMS"], pen["Ratio Sulf kWh/t TMS"], "kWh/t"),
    ("Ratio Óx TMS\n(kWh/t)",
     ult["Ratio Óx TMS kWh/t"],   pen["Ratio Óx TMS kWh/t"],   "kWh/t"),
    ("Ratio Óx-EW\n(kWh/tmf)",
     r_elec_ult,                   r_elec_pen,                   "kWh/tmf"),
    ("Ratio Infra\n(kWh/m³)",
     r_infra_ult,                  r_infra_pen,                  "kWh/m³"),
]
for col, (lbl, val, val_prev, unit) in zip(cols_rat, ratio_data):
    col.metric(
        lbl,
        f"{val:,.2f} {unit}" if val > 0 else "—",
        delta=delta_pct(val, val_prev) if val > 0 and val_prev > 0 else None,
    )

st.divider()

# ─────────────────────────────────────────────
# GRÁFICO PRINCIPAL — CONSUMO MENSUAL
# ─────────────────────────────────────────────
st.subheader("Consumo mensual por planta (GWh)")

# Construir series según vista
hist_plot = hist_filtered.copy()
hist_plot["Sulf_GWh"]  = hist_plot["Sulfuros Total (kWh)"] / 1e6
hist_plot["Oxid_GWh"]  = hist_plot["Óxidos Total (kWh)"] / 1e6
hist_plot["Infra_GWh"] = hist_plot["Infra Bombeo (kWh)"] / 1e6

proj_plot = proj.copy()
proj_plot["Sulf_GWh"]  = proj_plot["kWh Sulfuros"] / 1e6
proj_plot["Oxid_GWh"]  = proj_plot["kWh Óxidos TOTAL"] / 1e6
proj_plot["Infra_GWh"] = proj_plot["kWh Infraestruc."] / 1e6

fig_main = go.Figure()

def add_bars(df, suffix, opacity=1.0, dash=False):
    name_suffix = " (proy.)" if dash else " (real)"
    showleg = not dash
    for col, lbl, color in [
        ("Sulf_GWh",  "Sulfuros",     COLORS["Sulfuros"]),
        ("Oxid_GWh",  "Óxidos Total", COLORS["Óxidos Total"]),
        ("Infra_GWh", "Infra",        COLORS["Infraestructura"]),
    ]:
        if planta_sel != "Todas":
            if planta_sel == "Sulfuros" and lbl != "Sulfuros": continue
            if planta_sel == "Óxidos" and lbl != "Óxidos Total": continue
            if planta_sel == "Infraestructura" and lbl != "Infra": continue

        marker_color = color if not dash else color.replace("#","")
        fig_main.add_trace(go.Bar(
            x=df["Fecha"], y=df[col],
            name=lbl + name_suffix,
            marker_color=color,
            opacity=opacity,
            legendgroup=lbl,
            showlegend=showleg,
        ))

if vista in ("Histórico (2022–2026)", "Histórico + Proyectado"):
    add_bars(hist_plot, "hist", opacity=1.0)
if vista in ("Proyectado (2026–2045)", "Histórico + Proyectado"):
    add_bars(proj_plot, "proj", opacity=0.45, dash=True)

fig_main.update_layout(
    barmode="stack",
    height=320,
    margin=dict(l=0, r=0, t=10, b=0),
    legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
    xaxis_title=None,
    yaxis_title="GWh",
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(gridcolor="rgba(128,128,128,0.15)"),
    yaxis=dict(gridcolor="rgba(128,128,128,0.15)"),
)
st.plotly_chart(fig_main, use_container_width=True)

# ─────────────────────────────────────────────
# FILA 2 — RATIOS + DISTRIBUCIÓN
# ─────────────────────────────────────────────
col_rat, col_dist = st.columns([2, 1])

with col_rat:
    st.subheader("Ratios de consumo unitario")
    tab_s, tab_o, tab_i = st.tabs(["Sulfuros (kWh/t)", "Óxidos TMS (kWh/t)", "Infra (kWh/m³)"])

    def ratio_chart(hist_col, proj_col, color, ymin, ymax):
        fig = go.Figure()
        # Histórico
        fig.add_trace(go.Scatter(
            x=hist_filtered["Fecha"], y=hist_filtered[hist_col],
            mode="lines+markers", name="Real",
            line=dict(color=color, width=2),
            marker=dict(size=3),
        ))
        # Proyectado
        fig.add_trace(go.Scatter(
            x=proj["Fecha"], y=proj[proj_col],
            mode="lines", name="Proyectado",
            line=dict(color=color, width=2, dash="dash"),
        ))
        # Media base
        mean_val = hist_filtered[hist_col].dropna().mean()
        fig.add_hline(y=mean_val, line_dash="dot",
                      line_color="gray", annotation_text=f"Media {mean_val:.2f}",
                      annotation_position="top left")
        fig.update_layout(
            height=220, margin=dict(l=0,r=0,t=10,b=0),
            legend=dict(orientation="h", yanchor="bottom", y=1.01),
            yaxis=dict(range=[ymin, ymax], gridcolor="rgba(128,128,128,0.15)"),
            xaxis=dict(gridcolor="rgba(128,128,128,0.15)"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        )
        return fig

    with tab_s:
        st.plotly_chart(
            ratio_chart("Ratio Sulf kWh/t TMS", "Ratio Sulf (kWh/t)",
                        COLORS["Sulfuros"], 18, 28),
            use_container_width=True
        )
    with tab_o:
        st.plotly_chart(
            ratio_chart("Ratio Óx TMS kWh/t", "Ratio Ox TMS (kWh/t)",
                        COLORS["Óxidos sin Elec."], 4.5, 10),
            use_container_width=True
        )
    with tab_i:
        # Infra only available from 2025
        hist_i = hist_filtered[hist_filtered["Agua Mar (m3)"] > 0].copy()
        hist_i["ratio_infra"] = hist_i["Infra Bombeo (kWh)"] / hist_i["Agua Mar (m3)"]
        fig_i = go.Figure()
        fig_i.add_trace(go.Scatter(
            x=hist_i["Fecha"], y=hist_i["ratio_infra"],
            mode="lines+markers", name="Real",
            line=dict(color=COLORS["Infraestructura"], width=2), marker=dict(size=3),
        ))
        fig_i.add_trace(go.Scatter(
            x=proj["Fecha"], y=proj["Ratio Infra (kWh/m3)"],
            mode="lines", name="Proyectado",
            line=dict(color=COLORS["Infraestructura"], width=2, dash="dash"),
        ))
        mean_i = hist_i["ratio_infra"].mean()
        fig_i.add_hline(y=mean_i, line_dash="dot", line_color="gray",
                        annotation_text=f"Media {mean_i:.3f}")
        fig_i.update_layout(
            height=220, margin=dict(l=0,r=0,t=10,b=0),
            legend=dict(orientation="h", yanchor="bottom", y=1.01),
            yaxis=dict(range=[5.5, 6.3], gridcolor="rgba(128,128,128,0.15)"),
            xaxis=dict(gridcolor="rgba(128,128,128,0.15)"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_i, use_container_width=True)

with col_dist:
    st.subheader("Distribución promedio")
    # Use latest 12 months of historical
    last12 = hist_filtered.tail(12)
    avg_s = last12["Sulfuros Total (kWh)"].mean() / 1e6
    avg_o = last12["Óxidos Total (kWh)"].mean() / 1e6
    avg_i = last12["Infra Bombeo (kWh)"].mean() / 1e6
    total_avg = avg_s + avg_o + avg_i

    fig_pie = go.Figure(go.Pie(
        labels=["Sulfuros", "Óxidos", "Infraestructura"],
        values=[avg_s, avg_o, avg_i],
        hole=0.65,
        marker_colors=[COLORS["Sulfuros"], COLORS["Óxidos Total"], COLORS["Infraestructura"]],
        textinfo="percent",
        textfont_size=12,
    ))
    fig_pie.update_layout(
        height=220, margin=dict(l=0,r=0,t=10,b=0),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.15),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        annotations=[dict(text=f"{total_avg:.1f}<br>GWh/mes", x=0.5, y=0.5,
                          font_size=13, showarrow=False)],
    )
    st.plotly_chart(fig_pie, use_container_width=True)

    # KPI ratios del último mes
    st.caption("Ratios — último mes real")
    r_s = hist.iloc[-1]["Ratio Sulf kWh/t TMS"]
    r_o = hist.iloc[-1]["Ratio Óx TMS kWh/t"]
    r_e = hist.iloc[-1]["Ratio Electrodep. kWh/tmf"]
    st.markdown(f"""
    | Planta | Ratio |
    |---|---|
    | Sulfuros | **{r_s:.2f}** kWh/t |
    | Óxidos TMS | **{r_o:.2f}** kWh/t |
    | Electrodep. | **{r_e:,.0f}** kWh/tmf |
    """)

st.divider()

# ─────────────────────────────────────────────
# PROYECCIÓN ANUAL LOM 2026-2045
# ─────────────────────────────────────────────
st.subheader("Proyección anual 2026–2045 (GWh/año)")

fig_ann = go.Figure()
for col, lbl, color in [
    ("Sulfuros (GWh)",        "Sulfuros",       COLORS["Sulfuros"]),
    ("Óxidos TOTAL (GWh)",    "Óxidos Total",   COLORS["Óxidos Total"]),
    ("Infraestructura (GWh)", "Infra",          COLORS["Infraestructura"]),
]:
    fig_ann.add_trace(go.Bar(
        x=ann["Año"], y=ann[col],
        name=lbl, marker_color=color,
    ))

fig_ann.update_layout(
    barmode="stack", height=280,
    margin=dict(l=0, r=0, t=10, b=0),
    legend=dict(orientation="h", yanchor="bottom", y=1.01),
    xaxis_title=None, yaxis_title="GWh",
    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(gridcolor="rgba(128,128,128,0.15)", type="category"),
    yaxis=dict(gridcolor="rgba(128,128,128,0.15)"),
)
st.plotly_chart(fig_ann, use_container_width=True)

# ─────────────────────────────────────────────
# TABLA DETALLE PROYECCIÓN
# ─────────────────────────────────────────────
with st.expander("Ver tabla completa de proyección período a período"):
    display_cols = {
        "Período":             "Período",
        "Frecuencia":          "Frecuencia",
        "kWh Sulfuros":        "kWh Sulfuros",
        "kWh Óxidos TOTAL":    "kWh Óxidos Total",
        "kWh Infraestruc.":    "kWh Infra",
        "kWh TOTAL":           "kWh Total",
        "GWh Total":           "GWh Total",
        "Ratio Sulf (kWh/t)":  "Ratio Sulf",
        "Ratio Ox TMS (kWh/t)":"Ratio Ox TMS",
        "Ratio Infra (kWh/m3)":"Ratio Infra",
    }
    df_show = proj[[c for c in display_cols if c in proj.columns]].copy()
    df_show.columns = [display_cols[c] for c in df_show.columns]
    st.dataframe(
        df_show.style.format({
            "kWh Sulfuros":   "{:,.0f}",
            "kWh Óxidos Total": "{:,.0f}",
            "kWh Infra":      "{:,.0f}",
            "kWh Total":      "{:,.0f}",
            "GWh Total":      "{:.3f}",
            "Ratio Sulf":     "{:.2f}",
            "Ratio Ox TMS":   "{:.2f}",
            "Ratio Infra":    "{:.3f}",
        }),
        use_container_width=True,
        height=300,
    )
    csv = df_show.to_csv(index=False).encode("utf-8")
    st.download_button("Descargar CSV", csv, "proyeccion_energia.csv", "text/csv")

# ─────────────────────────────────────────────
# TABLA RESUMEN ANUAL
# ─────────────────────────────────────────────
with st.expander("Ver resumen anual 2026–2045"):
    ann_show = ann.copy()
    ann_show["Δ% vs año ant."] = ann_show["Δ% vs año ant."].apply(
        lambda x: f"{x*100:+.1f}%" if pd.notna(x) else "—"
    )
    st.dataframe(
        ann_show.style.format({
            "Sulfuros (GWh)":         "{:.3f}",
            "Óxidos sin Elec. (GWh)": "{:.3f}",
            "Electrodeposición (GWh)":"{:.3f}",
            "Óxidos TOTAL (GWh)":     "{:.3f}",
            "Infraestructura (GWh)":  "{:.3f}",
            "TOTAL (GWh)":            "{:.3f}",
            "Sulf TMS (Mt)":          "{:.3f}",
            "Oxid TMS (Mt)":          "{:.3f}",
            "Oxid TMF (kt)":          "{:.1f}",
            "Ratio total kWh/t sulf": "{:.2f}",
        }),
        use_container_width=True,
    )

st.divider()
st.caption("Energy Intelligence v1.0 · Modelo v5 · LOM25 Óptimo · Datos: Proyeccion_Energia_v5_FINAL.xlsx")
