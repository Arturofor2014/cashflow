import streamlit as st
import pandas as pd
import plotly.express as px

RUTA = "https://docs.google.com/spreadsheets/d/1aVxN_nlhGUqh-79i5FKV-eNwMZJumZ2D/export?format=xlsx"
HOJA = "Cash Flow Budget & Executed"

ENCABEZADOS = [
    "Fecha",
    "Dic - 2025",
    "Ene - 2026", "Feb - 2026", "Mar - 2026", "Abr - 2026", "May - 2026",
    "Jun - 2026", "Jul - 2026", "Ago - 2026", "Sep - 2026", "Oct - 2026",
    "Nov - 2026", "Dic - 2026", "Total Budget 2026",
    "Ene - 2027", "Feb - 2027", "Mar - 2027", "Abr - 2027", "May - 2027",
    "Jun - 2027", "Jul - 2027", "Ago - 2027", "Sep - 2027", "Oct - 2027",
    "Nov - 2027", "Dic - 2027", "Total Budget 2027",
    "Ene - 2028", "Feb - 2028", "Mar - 2028", "Abr - 2028", "May - 2028",
    "Jun - 2028", "Jul - 2028", "Ago - 2028", "Sep - 2028", "Oct - 2028",
    "Nov - 2028", "Dic - 2028", "Total Budget 2028",
]

FILAS_SECCION = {
    "Ingreso", "Otros Ingresos",
    "Administración", "Capital Humano", "Marketing", "Operativos",
}
FILAS_TOTAL = {
    "Total Ingresos", "Total Otros Ingresos", "Total de Ingresos", "Total de Efectivo Disponible",
    "Subtotal Administración", "Subtotal Capital Humano", "Subtotal Marketing",
    "Total Structural Costs", "Subtotal Operativos", "Total Gastos Operativos",
    "Total Otros Egresos", "Total Financial Costs", "Total Taxes",
    "Total Non Cash Expenses", "Total Ajustes", "Efectivo Cierre de Mes",
}

COLS_TOTAL_BUDGET = {"Total Budget 2026", "Total Budget 2027", "Total Budget 2028"}

COLS_VERDE    = ["Dic - 2025", "Ene - 2026", "Feb - 2026", "Mar - 2026"]
COLS_AMARILLO = [
    "Abr - 2026", "May - 2026", "Jun - 2026", "Jul - 2026", "Ago - 2026",
    "Sep - 2026", "Oct - 2026", "Nov - 2026", "Dic - 2026", "Total Budget 2026",
]
COLS_CELESTE  = [
    "Ene - 2027", "Feb - 2027", "Mar - 2027", "Abr - 2027", "May - 2027",
    "Jun - 2027", "Jul - 2027", "Ago - 2027", "Sep - 2027", "Oct - 2027",
    "Nov - 2027", "Dic - 2027", "Total Budget 2027",
]
COLS_TURQUESA = [
    "Ene - 2028", "Feb - 2028", "Mar - 2028", "Abr - 2028", "May - 2028",
    "Jun - 2028", "Jul - 2028", "Ago - 2028", "Sep - 2028", "Oct - 2028",
    "Nov - 2028", "Dic - 2028", "Total Budget 2028",
]

st.set_page_config(page_title="Cash Flow Dashboard", layout="wide")

st.markdown("""
<meta name="viewport" content="width=1024">
<style>
html {
    scroll-padding-top: 11rem;
    min-width: 1024px;
}
.navbar {
    position: fixed;
    top: 3.5rem;
    left: 0;
    right: 0;
    z-index: 999;
    background-color: #f8f9fa;
    padding: 8px 16px 12px 16px;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 6px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.12);
}
.navbar-title {
    font-size: 34px;
    font-weight: 900;
    color: #0052FF;
    letter-spacing: 2px;
    text-transform: uppercase;
}
.navbar-buttons {
    display: flex;
    gap: 8px;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
    flex-wrap: nowrap;
    justify-content: center;
    width: 100%;
}
.navbar::-webkit-scrollbar {
    height: 3px;
}
.navbar::-webkit-scrollbar-thumb {
    background: #4C9BE8;
    border-radius: 3px;
}
.navbar a {
    text-decoration: none;
    padding: 8px 18px;
    background-color: #4C9BE8;
    color: white !important;
    border-radius: 6px;
    font-weight: 600;
    font-size: 14px;
    white-space: nowrap;
    flex-shrink: 0;
}
.navbar a:hover {
    background-color: #2176c7;
}
.main .block-container {
    padding-top: 11rem;
    padding-left: 1rem;
    padding-right: 1rem;
}
@media (max-width: 768px) {
    .main .block-container {
        padding-top: 6rem;
        padding-left: 0.5rem;
        padding-right: 0.5rem;
    }
    .navbar {
        top: 3rem;
        padding: 6px 8px;
    }
}
</style>

<div class="navbar">
    <div class="navbar-title">Cash Flow Dashboard</div>
    <div class="navbar-buttons">
        <a href="#inicio">Inicio</a>
        <a href="#ingreso">Ingresos</a>
        <a href="#gastos-operativo">Gastos Operativos</a>
        <a href="#otros-egresos">Otros Egresos</a>
        <a href="#financial-outflows">Financial Outflows</a>
        <a href="#taxes">Taxes</a>
        <a href="#financial-outflows-others">Financial Outflows Others</a>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div id="top"></div>', unsafe_allow_html=True)

cols_numericas = ENCABEZADOS[1:]



@st.cache_data
def cargar_datos():
    df = pd.read_excel(
        str(RUTA),
        sheet_name=HOJA,
        header=None,
        skiprows=2
    )
    df = df.iloc[:, :len(ENCABEZADOS)]
    df.columns = ENCABEZADOS
    return df


def limpiar_bloque(df_bloque):
    bloque = df_bloque.copy().reset_index(drop=True)
    bloque["Fecha"] = bloque["Fecha"].astype(str).replace(["None", "nan", "NaN"], "—")
    for col in cols_numericas:
        bloque[col] = pd.to_numeric(bloque[col], errors="coerce")
    return bloque


def fila_suma(df_bloque, label):
    fila = {"Fecha": label}
    for col in cols_numericas:
        vals = pd.to_numeric(df_bloque[col], errors="coerce")
        fila[col] = vals.sum() if vals.notna().any() else ""
    return fila


COLS_POR_AÑO = {
    "2026": COLS_VERDE + COLS_AMARILLO,
    "2027": COLS_CELESTE,
    "2028": COLS_TURQUESA,
}

def filtro_años(key):
    c1, c2, c3, _ = st.columns([1, 1, 1, 5])
    a2026 = c1.checkbox("2026", value=True,  key=f"{key}_2026")
    a2027 = c2.checkbox("2027", value=False, key=f"{key}_2027")
    a2028 = c3.checkbox("2028", value=False, key=f"{key}_2028")
    cols = []
    if a2026: cols += COLS_POR_AÑO["2026"]
    if a2027: cols += COLS_POR_AÑO["2027"]
    if a2028: cols += COLS_POR_AÑO["2028"]
    return cols


def filtrar_tabla(df, cols_sel):
    cols_presentes = [c for c in cols_sel if c in df.columns]
    return df[["Fecha"] + cols_presentes]


def formato_valores(x):
    try:
        if x is None:
            return "—"
        if str(x).strip() in ("", "None", "nan", "NaN"):
            return "—"
        if pd.isna(x):
            return "—"
        v = float(x)
        if v == 0:
            return "—"
        return f"{v:,.0f}"
    except:
        return "—"


def estilo_filas(row):
    if row["Fecha"] in FILAS_TOTAL:
        return ["background-color: #d0e8ff; font-weight: bold"] * len(row)
    if row["Fecha"] in FILAS_SECCION:
        return ["background-color: #fff3cd; font-weight: bold"] * len(row)
    return [""] * len(row)



COL_BG = {c: "#e8f5e9" for c in COLS_VERDE}
COL_BG.update({c: "#fffde7"  for c in COLS_AMARILLO})
COL_BG.update({c: "#f5f5f5"  for c in COLS_CELESTE})
COL_BG.update({c: "#fff3e0"  for c in COLS_TURQUESA})


def mostrar_tabla(df, header_color="#d0e8ff", fecha_bg=None, fecha_width=220):
    cols_num_df = [c for c in cols_numericas if c in df.columns]

    # Encabezados
    ths = []
    for i, col in enumerate(df.columns):
        if i == 0:
            bg = fecha_bg if fecha_bg else (header_color if header_color else COL_BG.get(col, "#ffffff"))
        else:
            bg = header_color if header_color else COL_BG.get(col, "#ffffff")
        base = f"padding:7px 10px;border:1px solid #ddd;font-weight:bold;white-space:nowrap;font-size:13px;background:{bg};color:#000000;"
        if i == 0:
            s = f"padding:7px 10px;border:1px solid #ddd;font-weight:bold;font-size:13px;word-wrap:break-word;white-space:normal;position:sticky;left:0;z-index:3;background:{bg};min-width:{fecha_width}px;max-width:{fecha_width}px;"
        else:
            s = base + "text-align:right;"
        ths.append(f'<th style="{s}">{col}</th>')

    # Filas
    trs = []
    for _, row in df.iterrows():
        fecha = row["Fecha"]
        if fecha in FILAS_TOTAL:
            row_bg, fw = "#e0f7fa", "bold"
        elif fecha in FILAS_SECCION:
            row_bg, fw = "#fff3cd", "bold"
        else:
            row_bg, fw = None, "normal"

        tds = []
        for i, col in enumerate(df.columns):
            bg = row_bg if row_bg else COL_BG.get(col, "#ffffff")
            val = formato_valores(row[col]) if col in cols_num_df else (str(row[col]) if str(row[col]) not in ("None","nan","NaN","") else "—")
            base = f"padding:6px 10px;border:1px solid #ddd;font-weight:{fw};white-space:nowrap;font-size:13px;color:#000000;"
            if i == 0:
                cell_bg = fecha_bg if fecha_bg else bg
                s = f"padding:6px 10px;border:1px solid #ddd;font-weight:{fw};font-size:13px;word-wrap:break-word;white-space:normal;position:sticky;left:0;z-index:1;background:{cell_bg};min-width:{fecha_width}px;max-width:{fecha_width}px;"
            else:
                s = base + f"background:{bg};text-align:right;"
            tds.append(f'<td style="{s}">{val}</td>')
        trs.append(f'<tr>{"".join(tds)}</tr>')

    html = f"""
    <div style="overflow-x:auto;border:1px solid #ddd;border-radius:6px;margin-bottom:16px;">
      <table style="border-collapse:collapse;width:100%;font-family:sans-serif;">
        <thead><tr>{"".join(ths)}</tr></thead>
        <tbody>{"".join(trs)}</tbody>
      </table>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


df_raw = cargar_datos()

# ── Fila 3 del Excel (índice 0): saldo inicial
fila_3 = limpiar_bloque(df_raw.iloc[[0]])
FILAS_TOTAL.add(fila_3.iloc[0]["Fecha"])

# ── CHART 1: Ingresos — Excel filas 5-10 → iloc[2:8]
bloque1 = limpiar_bloque(df_raw.iloc[2:8])
total_ingresos = fila_suma(bloque1, "Total Ingresos")
df_chart1 = pd.concat([bloque1, pd.DataFrame([total_ingresos])], ignore_index=True)

# ── CHART 2: Otros Ingresos — Excel filas 12-19 → iloc[9:17]
bloque2 = limpiar_bloque(df_raw.iloc[9:17])
total_otros = fila_suma(bloque2, "Total Otros Ingresos")
df_chart2 = pd.concat([bloque2, pd.DataFrame([total_otros])], ignore_index=True)

# ── CHART 3: Resumen
def sumar_filas(fila_a, fila_b, label):
    fila = {"Fecha": label}
    for col in cols_numericas:
        v1 = pd.to_numeric(fila_a.get(col, 0), errors="coerce")
        v2 = pd.to_numeric(fila_b.get(col, 0), errors="coerce")
        fila[col] = (0 if pd.isna(v1) else v1) + (0 if pd.isna(v2) else v2)
    return fila

total_de_ingresos = sumar_filas(total_ingresos, total_otros, "Total de Ingresos")

fila3_vals = fila_3.iloc[0].to_dict()
total_ef_disp = sumar_filas(fila3_vals, total_de_ingresos, "Total de Efectivo Disponible")

for col in COLS_TOTAL_BUDGET:
    total_ef_disp[col] = 0

df_chart3 = pd.concat([fila_3, pd.DataFrame([total_ingresos, total_otros, total_de_ingresos, total_ef_disp])], ignore_index=True)

# ── GASTOS OPERATIVOS ──────────────────────────────────────────────────────────

# Chart GO1: Administración — Excel 25-43 (título + items), drop Subtotal en idx 19
EXCLUIR_TITULOS = {"Administración", "Capital Humano", "Marketing", "Operativos",
                   "Subtotal", "Subtotal Administración", "Subtotal Capital Humano",
                   "Subtotal Marketing", "Subtotal Operativos", "Total Structural Costs", "—"}

go1 = limpiar_bloque(df_raw.iloc[22:42])
go1_items = go1[~go1["Fecha"].isin(EXCLUIR_TITULOS)].reset_index(drop=True)
subtotal_adm = fila_suma(go1_items, "Subtotal Administración")
df_go1 = pd.concat([go1_items, pd.DataFrame([subtotal_adm])], ignore_index=True)

go2 = limpiar_bloque(df_raw.iloc[43:59])
go2_items = go2[~go2["Fecha"].isin(EXCLUIR_TITULOS)].reset_index(drop=True)
subtotal_ch = fila_suma(go2_items, "Subtotal Capital Humano")
df_go2 = pd.concat([go2_items, pd.DataFrame([subtotal_ch])], ignore_index=True)

go3 = limpiar_bloque(df_raw.iloc[60:63])
go3_items = go3[~go3["Fecha"].isin(EXCLUIR_TITULOS)].reset_index(drop=True)
total_marketing = fila_suma(go3_items, "Subtotal Marketing")
df_go3 = pd.concat([go3_items, pd.DataFrame([total_marketing])], ignore_index=True)


go4 = limpiar_bloque(df_raw.iloc[64:82])
go4_items = go4[~go4["Fecha"].isin(EXCLUIR_TITULOS)].reset_index(drop=True)
subtotal_op = fila_suma(go4_items, "Subtotal Operativos")
df_go4 = pd.concat([go4_items, pd.DataFrame([subtotal_op])], ignore_index=True)

# Chart GO5: Resumen Gastos Operativos
total_structural = sumar_filas(sumar_filas(subtotal_adm, subtotal_ch, ""), total_marketing, "")
total_gastos_op  = sumar_filas(total_structural, subtotal_op, "Total Gastos Operativos")
df_go5 = pd.DataFrame([subtotal_adm, subtotal_ch, total_marketing, subtotal_op, total_gastos_op])

# ── OTROS EGRESOS ───────────────────────────────────────────────────────────────
# Excel 91-95 → iloc[88:93], excluir título y total del Excel por valor
oe = limpiar_bloque(df_raw.iloc[86:91])
EXCLUIR_OE = {"Otros Egresos", "Total Otros egresos", "Total Otros Egresos", "—"}
oe = oe[~oe["Fecha"].isin(EXCLUIR_OE)].reset_index(drop=True)
total_otros_egresos = fila_suma(oe, "Total Otros Egresos")
df_oe = pd.concat([oe, pd.DataFrame([total_otros_egresos])], ignore_index=True)

# ── FINANCIAL OUTFLOWS ─────────────────────────────────────────────────────────
# Excel 98-101 → iloc[95:99], excluir título y total del Excel por valor
fo = limpiar_bloque(df_raw.iloc[93:97])
EXCLUIR_FO = {"Financial Outflows", "Total Financial Costs", "Total Financial Outflows", "—"}
fo = fo[~fo["Fecha"].isin(EXCLUIR_FO)].reset_index(drop=True)
total_financial = fila_suma(fo, "Total Financial Costs")
df_fo = pd.concat([fo, pd.DataFrame([total_financial])], ignore_index=True)

# ── TAXES ──────────────────────────────────────────────────────────────────────
# Excel 102-112 → iloc[99:110]
taxes = limpiar_bloque(df_raw.iloc[99:110])
taxes = taxes[~taxes["Fecha"].isin({"Taxes", "Total Taxes", "—"})].reset_index(drop=True)
total_taxes = fila_suma(taxes, "Total Taxes")
df_taxes = pd.concat([taxes, pd.DataFrame([total_taxes])], ignore_index=True)

# ── NON CASH EXPENSES ──────────────────────────────────────────────────────────
# Excel 115-117 → iloc[112:115]
nce = limpiar_bloque(df_raw.iloc[112:115])
nce = nce[~nce["Fecha"].isin({"Non Cash Expenses - Operativo", "Total Non Cash Expenses", "—"})].reset_index(drop=True)
total_nce = fila_suma(nce, "Total Non Cash Expenses")
df_nce = pd.concat([nce, pd.DataFrame([total_nce])], ignore_index=True)

# ── FINANCIAL OUTFLOWS - OTHERS ────────────────────────────────────────────────
# Excel 122-126 → iloc[119:124]
flo = limpiar_bloque(df_raw.iloc[117:122])
flo = flo[~flo["Fecha"].isin({"Financial Outflows", "Financial Outflows - Others", "Total Ajustes", "—"})].reset_index(drop=True)
total_ajustes = fila_suma(flo, "Total Ajustes")
df_flo = pd.concat([flo, pd.DataFrame([total_ajustes])], ignore_index=True)

# ── EFECTIVO CIERRE DE MES ─────────────────────────────────────────────────────
# Fila 126 (iloc[125]) = Ingresos, Fila 127 (iloc[126]) = Egresos
ecm = limpiar_bloque(df_raw.iloc[123:125])
efectivo_cierre = {"Fecha": "Efectivo Cierre de Mes"}
for col in cols_numericas:
    v1 = pd.to_numeric(ecm.iloc[0][col], errors="coerce")
    v2 = pd.to_numeric(ecm.iloc[1][col], errors="coerce")
    efectivo_cierre[col] = (0 if pd.isna(v1) else v1) - (0 if pd.isna(v2) else v2)
df_ecm = pd.concat([ecm, pd.DataFrame([efectivo_cierre])], ignore_index=True)

# ── RENDER ─────────────────────────────────────────────────────────────────────

# ── EFECTIVO CIERRE DE MES (primera sección) ───────────────────────────────────
st.markdown('<div id="inicio" style="margin-top:5rem;"></div>', unsafe_allow_html=True)
st.markdown('<h2 style="font-weight:900;letter-spacing:1px;color:#0052FF;margin-bottom:24px;">EFECTIVO DISPONIBLE POR PERIODO Y POR MES</h2>', unsafe_allow_html=True)

ecm_vals = {col: efectivo_cierre[col] for col in cols_numericas}
df_ecm_graf = pd.DataFrame({
    "Mes":   list(ecm_vals.keys()),
    "Valor": [pd.to_numeric(v, errors="coerce") for v in ecm_vals.values()],
}).dropna(subset=["Valor"])
df_ecm_graf = df_ecm_graf[df_ecm_graf["Mes"].isin(COLS_TOTAL_BUDGET)]
rows_html = ""
for _, r in df_ecm_graf.iterrows():
    val = f"{r['Valor']:,.0f}" if pd.notna(r['Valor']) else "—"
    rows_html += f'<div style="display:flex;justify-content:space-between;align-items:center;padding:12px 20px;border-bottom:1px solid #e0e0e0;"><span style="font-size:14px;font-weight:600;color:#000000;">{r["Mes"]}</span><span style="font-size:20px;font-weight:bold;color:#000000;">{val}</span></div>'

html_box = f'<div style="background:#ffffff;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.1);overflow:hidden;max-width:500px;margin-bottom:24px;"><div style="background:#d0e8ff;padding:12px 20px;"><span style="color:#000000;font-weight:700;font-size:14px;">Efectivo al Cierre del Periodo</span></div>{rows_html}</div>'

# ── Fila 1: cuadro de texto | tabla ECM
r1_left, _, r1_right = st.columns([1, 0.1, 3])
with r1_left:
    st.markdown(html_box, unsafe_allow_html=True)
with r1_right:
    st.markdown('<div id="inicio"></div>', unsafe_allow_html=True)
    st.subheader("Efectivo Cierre de Mes")
    cols_sel = filtro_años("ecm")
    mostrar_tabla(filtrar_tabla(df_ecm.drop(columns=[c for c in COLS_TOTAL_BUDGET if c in df_ecm.columns]), cols_sel), fecha_width=140)

# ── Filtros compartidos para los gráficos
st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
ay1, ay2, ay3, _ = st.columns([1, 1, 1, 5])
pa2026 = ay1.checkbox("2026", value=True,  key="pie_2026")
pa2027 = ay2.checkbox("2027", value=False, key="pie_2027")
pa2028 = ay3.checkbox("2028", value=False, key="pie_2028")

cols_pie = []
if pa2026: cols_pie += COLS_VERDE + COLS_AMARILLO
if pa2027: cols_pie += COLS_CELESTE
if pa2028: cols_pie += COLS_TURQUESA
meses_disp = [c for c in cols_pie if c in cols_numericas]

if meses_disp:
    meses_sel = st.multiselect("Selecciona período(s):", meses_disp,
                               default=[meses_disp[0]], key="pie_mes")

    def sumar_cat(total_dict, meses):
        return sum(pd.to_numeric(total_dict.get(m, 0), errors="coerce") or 0 for m in meses)

    # ── Fila 2: pastel GO | pastel Taxes/FO
    r2_left, r2_right = st.columns([1, 1])

    with r2_left:
        st.subheader("Gastos Operativos")
        cat_go = {
            "Administración": sumar_cat(subtotal_adm,    meses_sel),
            "Capital Humano": sumar_cat(subtotal_ch,     meses_sel),
            "Marketing":      sumar_cat(total_marketing, meses_sel),
            "Operativos":     sumar_cat(subtotal_op,     meses_sel),
        }
        df_pie1 = pd.DataFrame({"Categoría": list(cat_go.keys()), "Valor": list(cat_go.values())})
        df_pie1 = df_pie1[df_pie1["Valor"] > 0]
        if not df_pie1.empty:
            fig1 = px.pie(df_pie1, names="Categoría", values="Valor",
                          color_discrete_sequence=["#E8834C","#4C9BE8","#2ECC71","#E8C34C"])
            fig1.update_traces(textinfo="label+value+percent", texttemplate="%{label}<br>%{value:,.0f}<br>%{percent}", textposition="outside", textfont=dict(size=13))
            fig1.update_layout(height=400, showlegend=False, margin=dict(t=30,b=30,l=30,r=30))
            st.plotly_chart(fig1, use_container_width=True)

    with r2_right:
        st.subheader("Taxes / Outflows / Otros")
        cat_tx = {
            "Taxes":              sumar_cat(total_taxes,         meses_sel),
            "Financial Outflows": sumar_cat(total_financial,     meses_sel),
            "FO Others":          sumar_cat(total_ajustes,       meses_sel),
            "Otros Egresos":      sumar_cat(total_otros_egresos, meses_sel),
        }
        df_pie2 = pd.DataFrame({"Categoría": list(cat_tx.keys()), "Valor": list(cat_tx.values())})
        df_pie2 = df_pie2[df_pie2["Valor"] > 0]
        if not df_pie2.empty:
            fig2 = px.pie(df_pie2, names="Categoría", values="Valor",
                          color_discrete_sequence=["#E84C9B","#9B4CE8","#4CE8C3","#FF7043"])
            fig2.update_traces(textinfo="label+value+percent", texttemplate="%{label}<br>%{value:,.0f}<br>%{percent}", textposition="outside", textfont=dict(size=13))
            fig2.update_layout(height=400, showlegend=False, margin=dict(t=30,b=30,l=30,r=30))
            st.plotly_chart(fig2, use_container_width=True)
else:
    st.info("Selecciona al menos un año.")

st.markdown('<hr id="ingreso" style="border:none;border-top:2px solid #cccccc;margin:32px 0;">', unsafe_allow_html=True)

# ── GRÁFICO: Total de Efectivo Disponible por mes ─────────────────────────────
st.markdown('<h2 style="font-weight:900;letter-spacing:1px;color:#0052FF;">INGRESOS</h2>', unsafe_allow_html=True)
st.subheader("Total de Efectivo Disponible por Mes")
cols_sel = filtro_años("ing_graf")
ef_disp_vals = {col: total_ef_disp[col] for col in cols_sel}
df_grafico = pd.DataFrame({
    "Mes": list(ef_disp_vals.keys()),
    "Valor": [pd.to_numeric(v, errors="coerce") for v in ef_disp_vals.values()],
}).dropna(subset=["Valor"])
df_grafico = df_grafico[~df_grafico["Mes"].isin(COLS_TOTAL_BUDGET)]
colores = ["#2ECC71" if "Dic" in m else "#4C9BE8" for m in df_grafico["Mes"]]
fig = px.bar(df_grafico, x="Mes", y="Valor", text="Valor", labels={"Mes": "", "Valor": ""})
fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textangle=0,
                  textfont=dict(size=14), marker_color=colores)
fig.update_layout(xaxis_tickangle=-45, yaxis=dict(tickformat=",.0f"),
                  plot_bgcolor="white", height=500)
st.plotly_chart(fig, use_container_width=True)

st.subheader("Resumen de Ingresos")
cols_sel = filtro_años("ing_res")
mostrar_tabla(filtrar_tabla(df_chart3, cols_sel), header_color=None, fecha_bg="#f5f5f5")

st.subheader("Ingresos")
cols_sel = filtro_años("ing_t1")
mostrar_tabla(filtrar_tabla(df_chart1, cols_sel))

st.subheader("Otros Ingresos")
cols_sel = filtro_años("ing_t2")
mostrar_tabla(filtrar_tabla(df_chart2, cols_sel))


st.markdown('<hr style="border:none;border-top:2px solid #cccccc;margin:32px 0;">', unsafe_allow_html=True)

st.markdown('<div id="gastos-operativo" style="margin-top:3rem;"></div>', unsafe_allow_html=True)
st.markdown('<h2 style="font-weight:900;letter-spacing:1px;color:#0052FF;">GASTOS OPERATIVOS</h2>', unsafe_allow_html=True)
st.subheader("Total Gastos Operativos por Mes")
cols_sel = filtro_años("go_graf")
go_vals = {col: total_gastos_op[col] for col in cols_sel}
df_go_graf = pd.DataFrame({
    "Mes":   list(go_vals.keys()),
    "Valor": [pd.to_numeric(v, errors="coerce") for v in go_vals.values()],
}).dropna(subset=["Valor"])
colores_go = ["#2ECC71" if "Total" in m else "#E8834C" for m in df_go_graf["Mes"]]
fig_go = px.bar(df_go_graf, x="Mes", y="Valor", text="Valor", labels={"Mes": "", "Valor": ""})
fig_go.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textangle=0,
                     textfont=dict(size=14), marker_color=colores_go)
fig_go.update_layout(xaxis_tickangle=-45, yaxis=dict(tickformat=",.0f"),
                     plot_bgcolor="white", height=500)
st.plotly_chart(fig_go, use_container_width=True)

st.subheader("Resumen Gastos Operativos")
cols_sel = filtro_años("go_res")
mostrar_tabla(filtrar_tabla(df_go5, cols_sel), header_color=None, fecha_bg="#f5f5f5")

st.subheader("Administración")
cols_sel = filtro_años("go_adm")
mostrar_tabla(filtrar_tabla(df_go1, cols_sel))

st.subheader("Capital Humano")
cols_sel = filtro_años("go_ch")
mostrar_tabla(filtrar_tabla(df_go2, cols_sel))

st.subheader("Marketing")
cols_sel = filtro_años("go_mkt")
mostrar_tabla(filtrar_tabla(df_go3, cols_sel))

st.subheader("Operativos")
cols_sel = filtro_años("go_op")
mostrar_tabla(filtrar_tabla(df_go4, cols_sel))

st.markdown('<hr style="border:none;border-top:2px solid #cccccc;margin:32px 0;">', unsafe_allow_html=True)
st.markdown('<div id="otros-egresos" style="margin-top:3rem;"></div>', unsafe_allow_html=True)
st.markdown('<h2 style="font-weight:900;letter-spacing:1px;color:#0052FF;">OTROS EGRESOS</h2>', unsafe_allow_html=True)
st.subheader("Total Otros Egresos por Mes")
cols_sel = filtro_años("oe_graf")
oe_vals = {col: total_otros_egresos[col] for col in cols_sel}
df_oe_graf = pd.DataFrame({
    "Mes":   list(oe_vals.keys()),
    "Valor": [pd.to_numeric(v, errors="coerce") for v in oe_vals.values()],
}).dropna(subset=["Valor"])
colores_oe = ["#2ECC71" if "Total" in m else "#E84C9B" for m in df_oe_graf["Mes"]]
fig_oe = px.bar(df_oe_graf, x="Mes", y="Valor", text="Valor", labels={"Mes": "", "Valor": ""})
fig_oe.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textangle=0,
                     textfont=dict(size=14), marker_color=colores_oe)
fig_oe.update_layout(xaxis_tickangle=-45, yaxis=dict(tickformat=",.0f"),
                     plot_bgcolor="white", height=500)
st.plotly_chart(fig_oe, use_container_width=True)

st.subheader("Otros Egresos")
cols_sel = filtro_años("oe_tab")
mostrar_tabla(filtrar_tabla(df_oe, cols_sel))

st.markdown('<hr style="border:none;border-top:2px solid #cccccc;margin:32px 0;">', unsafe_allow_html=True)
st.markdown('<div id="financial-outflows" style="margin-top:3rem;"></div>', unsafe_allow_html=True)
st.markdown('<h2 style="font-weight:900;letter-spacing:1px;color:#0052FF;">FINANCIAL OUTFLOWS</h2>', unsafe_allow_html=True)
st.subheader("Total Financial Costs por Mes")
cols_sel = filtro_años("fo_graf")
fo_vals = {col: total_financial[col] for col in cols_sel}
df_fo_graf = pd.DataFrame({
    "Mes":   list(fo_vals.keys()),
    "Valor": [pd.to_numeric(v, errors="coerce") for v in fo_vals.values()],
}).dropna(subset=["Valor"])
colores_fo = ["#2ECC71" if "Total" in m else "#9B4CE8" for m in df_fo_graf["Mes"]]
fig_fo = px.bar(df_fo_graf, x="Mes", y="Valor", text="Valor", labels={"Mes": "", "Valor": ""})
fig_fo.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textangle=0,
                     textfont=dict(size=14), marker_color=colores_fo)
fig_fo.update_layout(xaxis_tickangle=-45, yaxis=dict(tickformat=",.0f"),
                     plot_bgcolor="white", height=500)
st.plotly_chart(fig_fo, use_container_width=True)

st.subheader("Financial Outflows")
cols_sel = filtro_años("fo_tab")
mostrar_tabla(filtrar_tabla(df_fo, cols_sel))

st.markdown('<hr style="border:none;border-top:2px solid #cccccc;margin:32px 0;">', unsafe_allow_html=True)
st.markdown('<div id="taxes" style="margin-top:3rem;"></div>', unsafe_allow_html=True)
st.markdown('<h2 style="font-weight:900;letter-spacing:1px;color:#0052FF;">TAXES</h2>', unsafe_allow_html=True)
st.subheader("Total Taxes por Mes")
cols_sel = filtro_años("tx_graf")
tx_vals = {col: total_taxes[col] for col in cols_sel}
df_tx_graf = pd.DataFrame({
    "Mes":   list(tx_vals.keys()),
    "Valor": [pd.to_numeric(v, errors="coerce") for v in tx_vals.values()],
}).dropna(subset=["Valor"])
colores_tx = ["#2ECC71" if "Total" in m else "#E8C34C" for m in df_tx_graf["Mes"]]
fig_tx = px.bar(df_tx_graf, x="Mes", y="Valor", text="Valor", labels={"Mes": "", "Valor": ""})
fig_tx.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textangle=0,
                     textfont=dict(size=14), marker_color=colores_tx)
fig_tx.update_layout(xaxis_tickangle=-45, yaxis=dict(tickformat=",.0f"),
                     plot_bgcolor="white", height=500)
st.plotly_chart(fig_tx, use_container_width=True)

st.subheader("Taxes")
cols_sel = filtro_años("tx_tab")
mostrar_tabla(filtrar_tabla(df_taxes, cols_sel))

st.markdown('<hr style="border:none;border-top:2px solid #cccccc;margin:32px 0;">', unsafe_allow_html=True)
# NON CASH EXPENSES — oculto temporalmente
# st.markdown('<div id="non-cash-expenses"></div>', unsafe_allow_html=True)
# ...

st.markdown('<hr style="border:none;border-top:2px solid #cccccc;margin:32px 0;">', unsafe_allow_html=True)
st.markdown('<div id="financial-outflows-others" style="margin-top:3rem;"></div>', unsafe_allow_html=True)
st.markdown('<h2 style="font-weight:900;letter-spacing:1px;color:#0052FF;">FINANCIAL OUTFLOWS - OTHERS</h2>', unsafe_allow_html=True)
st.subheader("Total Ajustes por Mes")
cols_sel = filtro_años("flo_graf")
flo_vals = {col: total_ajustes[col] for col in cols_sel}
df_flo_graf = pd.DataFrame({
    "Mes":   list(flo_vals.keys()),
    "Valor": [pd.to_numeric(v, errors="coerce") for v in flo_vals.values()],
}).dropna(subset=["Valor"])
colores_flo = ["#2ECC71" if "Total" in m else "#4CE8C3" for m in df_flo_graf["Mes"]]
fig_flo = px.bar(df_flo_graf, x="Mes", y="Valor", text="Valor", labels={"Mes": "", "Valor": ""})
fig_flo.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textangle=0,
                      textfont=dict(size=14), marker_color=colores_flo)
fig_flo.update_layout(xaxis_tickangle=-45, yaxis=dict(tickformat=",.0f"),
                      plot_bgcolor="white", height=500)
st.plotly_chart(fig_flo, use_container_width=True)

st.subheader("Financial Outflows - Others")
cols_sel = filtro_años("flo_tab")
mostrar_tabla(filtrar_tabla(df_flo, cols_sel))
