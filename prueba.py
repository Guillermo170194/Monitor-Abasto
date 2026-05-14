import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import os
import base64

from pptx import Presentation
from pptx.util import Inches, Pt

st.set_page_config(page_title="Consolidada IMB", layout="wide")
from PIL import Image

# =========================
# ESTILO MARK 1
# =========================
st.markdown("""
<style>

/* ===== APP ===== */
.stApp {
    background-color: #F4F6F9;
    font-family: 'Segoe UI', sans-serif;
    color: #1E1E1E;
}
html, body, [class*="css"] {
    font-size: 16px;
    font-weight: 600;
    color: #1E1E1E;
}

/* ===== HEADER ===== */
.logo-card {
    background: white;
    padding: 12px;
    border-radius: 18px;
    display: flex;
    justify-content: center;
    align-items: center;
    height: 110px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.06);
}
.main-title {
    font-size: 46px;
    font-weight: 900;
    color: #9F2241;
    text-align: center;
    margin-bottom: 0px;
    letter-spacing: 1px;
}

.sub-title {
    text-align: center;
    color: #235B4E;
    font-size: 18px;
    margin-top: -8px;
    margin-bottom: 30px;
    font-weight: 500;
}

h1 {
    color:#9F2241 !important;
    font-weight:900 !important;
    font-size:52px !important;
    margin-bottom:0px !important;
}

h3 {
    color:#235B4E !important;
    font-weight:600 !important;
    margin-top:-10px !important;
}
/* ===== SIDEBAR ===== */
section[data-testid="stSidebar"] {
    background: linear-gradient(
        180deg,
        #235B4E 0%,
        #163832 100%
    );
}

section[data-testid="stSidebar"] * {
    color: white !important;
}

/* ===== KPI PREMIUM ===== */
.kpi-card {
    background: white;
    padding: 22px;
    border-radius: 22px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.08);
    transition: 0.3s;
    border-left: 8px solid #235B4E;
}

.kpi-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 8px 20px rgba(0,0,0,0.12);
}

.kpi-title {
    font-size: 17px;
    font-weight: 700;
    color: #666;
    margin-bottom: 10px;
}

.kpi-value {
    font-size: 38px;
    font-weight: 900;
    color: #1E1E1E;
}

.kpi-icon {
    font-size: 32px;
}

.kpi-green {
    border-left: 8px solid #1FAE4B;
}

.kpi-red {
    border-left: 8px solid #C62828;
}

.kpi-yellow {
    border-left: 8px solid #F9A825;
}

.kpi-blue {
    border-left: 8px solid #1565C0;
}
/* ===== TABLAS ===== */
[data-testid="stDataFrame"] {
    background: white;
    border-radius: 18px;
    padding: 10px;
    border: none;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
}

/* ===== HEADERS TABLAS ===== */
thead tr th {
    background-color:#235B4E !important;
    color:white !important;
    font-weight:bold !important;
    border:none !important;
}

/* ===== BOTONES ===== */
.stButton > button {
    width: 100%;
    border-radius: 12px;
    height: 50px;
    border: none;
    background-color: #235B4E;
    color: white;
    font-weight: bold;
    transition: 0.3s;
}

.stButton > button:hover {
    background-color: #9F2241;
    transform: scale(1.02);
}

/* ===== INPUTS ===== */
.stTextInput input {
    border-radius: 12px;
}

/* ===== SELECT ===== */
div[data-baseweb="select"] {
    background-color: white;
    border-radius: 12px;
}
.stRadio > div {
    background: rgba(255,255,255,0.08);
    padding: 10px;
    border-radius: 12px;
}

</style>
""", unsafe_allow_html=True)

# =========================
# ARCHIVO
# =========================
# =========================
# CARGA DE ARCHIVO
# =========================

st.sidebar.markdown("""
# 🧠 MARK 1

### Centro de Monitoreo

---
""")
menu = st.sidebar.radio(
    "Navegación",
    [
        "🏠 Dashboard Nacional",
        "📍 Estados",
        "🔎 Buscar Clave",
        "🚨 Alertas",
        "📊 Analítica",
        "📤 Exportaciones"
    ]
)
with st.sidebar.expander("📂 Cargar archivos"):

    archivo = st.file_uploader(
        "Consolidada",
        type=["xlsb", "xlsx"]
    )

    archivo_inv = st.file_uploader(
        "Inventario",
        type=["xlsx", "xlsb"]
    )

    archivo_cpm = st.file_uploader(
        "CPM",
        type=["xlsx", "xlsb"]
    )
if (
    archivo is None or
    archivo_inv is None or
    archivo_cpm is None
):
    st.warning("⬅ Debes cargar todos los archivos.")
    st.stop()
@st.cache_data(show_spinner=False)
def cargar():
    reemplazos = {
        "MICHOACAN": "MICHOACAN DE OCAMPO",
        "CIUDAD DE MEXICO": "CIUDAD DE MEXICO",
        "CDMX": "CIUDAD DE MEXICO",
        "MEXICO D.F.": "CIUDAD DE MEXICO",
        "DISTRITO FEDERAL": "CIUDAD DE MEXICO",
        "BAJA SUR": "BAJA CALIFORNIA SUR",
        "B.C.S.": "BAJA CALIFORNIA SUR",
        "BCS": "BAJA CALIFORNIA SUR"
    }
    df = pd.read_excel(
        archivo,
        engine="pyxlsb",
        usecols=[0, 3, 17, 18, 24, 27, 55]
    )
    df.columns = df.columns.str.strip()
    # =========================
    # INVENTARIO
    # =========================

    inv = pd.read_excel(
        archivo_inv
    )

    inv.columns = inv.columns.str.strip()

    col_inv_estado = inv.columns[0]
    col_inv_clave = inv.columns[3]
    col_inv_piezas = inv.columns[5]

    inv["ENTIDAD"] = (
        inv[col_inv_estado]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
    )

    inv["ENTIDAD"] = (
        inv["ENTIDAD"]
        .replace(reemplazos)
    )

    inv["CLAVE"] = (
        inv[col_inv_clave]
        .astype(str)
        .str.strip()
    )

    inv["CLAVE_MERGE"] = (
        inv["CLAVE"]
        .str.replace(".", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    inv["PIEZAS_INV"] = (
        inv[col_inv_piezas]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    # =========================
    # FILTRAR INVENTARIO VIGENTE
    # =========================

    col_inv_estatus = inv.columns[7]

    inv["ESTATUS"] = (
        inv[col_inv_estatus]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    inv = inv[
        ~inv["ESTATUS"].str.contains(
            "CADUC",
            na=False
        )
    ]
    inv["PIEZAS_INV"] = pd.to_numeric(
        inv["PIEZAS_INV"],
        errors="coerce"
    ).fillna(0)

    # CONSOLIDAR INVENTARIO
    inv_group = (
        inv.groupby(
            ["ENTIDAD", "CLAVE_MERGE"],
            dropna=False
        )["PIEZAS_INV"]
        .sum()
        .reset_index()
    )
    # =========================
    # CPM
    # =========================

    cpm = pd.read_excel(
        archivo_cpm
    )

    cpm.columns = cpm.columns.str.strip()

    col_cpm_estado = cpm.columns[0]
    col_cpm_clave = cpm.columns[1]
    col_cpm = cpm.columns[3]

    cpm["ENTIDAD"] = (
        cpm[col_cpm_estado]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
    )

    cpm["ENTIDAD"] = (
        cpm["ENTIDAD"]
        .replace(reemplazos)
    )

    cpm["CLAVE"] = (
        cpm[col_cpm_clave]
        .astype(str)
        .str.strip()
    )

    cpm["CLAVE_MERGE"] = (
        cpm["CLAVE"]
        .str.replace(".", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    cpm["CPM"] = (
        cpm[col_cpm]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )

    cpm["CPM"] = pd.to_numeric(
        cpm["CPM"],
        errors="coerce"
    ).fillna(0)

    cpm_group = (
        cpm.groupby(
            ["ENTIDAD", "CLAVE_MERGE"],
            dropna=False
        )["CPM"]
        .sum()
        .reset_index()
    )
    # =========================
    # LIMPIAR CONSOLIDADA
    # =========================

    df["ENTIDAD"] = (
        df[df.columns[1]]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
    )

    df["ENTIDAD"] = (
        df["ENTIDAD"]
        .replace(reemplazos)
    )

    df[df.columns[1]] = df["ENTIDAD"]

    df["CLAVE"] = (
        df[df.columns[2]]
        .astype(str)
        .str.strip()
    )

    df["CLAVE_MERGE"] = (
        df["CLAVE"]
        .str.replace(".", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    # =========================
    # MERGE INVENTARIO
    # =========================

    df = df.merge(
        inv_group,
        how="left",
        on=["ENTIDAD", "CLAVE_MERGE"],
    )

    df["PIEZAS_INV"] = (
        df["PIEZAS_INV"]
        .fillna(0)
    )

    # =========================
    # MERGE CPM
    # =========================

    df = df.merge(
        cpm_group,
        how="left",
        on=["ENTIDAD", "CLAVE_MERGE"],
    )

    df["CPM"] = (
        df["CPM"]
        .fillna(0)
    )

    # =========================
    # AGREGAR CLAVES SIN EMISIÓN
    # =========================

    base_extra = pd.merge(
        inv_group,
        cpm_group,
        how="outer",
        on=["ENTIDAD", "CLAVE_MERGE"]
    )

    existentes = set(
        zip(df["ENTIDAD"], df["CLAVE_MERGE"])
    )

    base_extra = base_extra[
        ~base_extra.apply(
            lambda x: (
                x["ENTIDAD"],
                x["CLAVE_MERGE"]
            ) in existentes,
            axis=1
        )
    ]

    if len(base_extra) > 0:

        extra = pd.DataFrame({
            df.columns[0]: "SIN EMISION",
            df.columns[1]: base_extra["ENTIDAD"],
            df.columns[2]: base_extra["CLAVE_MERGE"],	
            df.columns[3]: "SIN DESCRIPCIÓN",
            df.columns[4]: 0,
            df.columns[5]: 0,
            df.columns[6]: 0,
            "PIEZAS_INV": base_extra["PIEZAS_INV"],
            "CPM": base_extra["CPM"]
        })

        df = pd.concat(
            [df, extra],
            ignore_index=True
        )
    return df, inv_group, cpm_group

df, inv_group, cpm_group = cargar()
# =========================
# HEADER LIMPIO MARK 1
# =========================
col_titulo, col_status = st.columns([6,1.5])

with col_titulo:

    st.title("MARK 1")

    st.subheader(
        "Centro Nacional de Monitoreo Estratégico de Abasto"
    )

with col_status:

    st.success(
        "🟢 SISTEMA ACTIVO"
    )

# =========================
# COLUMNAS
# =========================
col_orden = df.columns[0]
col_estado = df.columns[1]
col_clave = df.columns[2]
col_descripcion = df.columns[3]
col_precio = df.columns[4]
col_emitidas = df.columns[5]
col_entregadas = df.columns[6]
# =========================
# BASE BUSCADOR
# =========================

df["PIEZAS_INV"] = pd.to_numeric(
    df["PIEZAS_INV"],
    errors="coerce"
).fillna(0)

df["CPM"] = pd.to_numeric(
    df["CPM"],
    errors="coerce"
).fillna(0)
base_busqueda = pd.merge(
    inv_group,
    cpm_group,
    how="outer",
    on=["ENTIDAD", "CLAVE_MERGE"]
)

base_busqueda["PIEZAS_INV"] = (
    base_busqueda["PIEZAS_INV"]
    .fillna(0)
)

base_busqueda["CPM"] = (
    base_busqueda["CPM"]
    .fillna(0)
)

emitido = (
    df.groupby(
        ["ENTIDAD", "CLAVE_MERGE"],
        dropna=False
    )[col_emitidas]
    .sum()
    .reset_index()
)

entregado = (
    df.groupby(
        ["ENTIDAD", "CLAVE_MERGE"],
        dropna=False
    )[col_entregadas]
    .sum()
    .reset_index()
)

descripcion = (
    df.groupby(
        ["ENTIDAD", "CLAVE_MERGE"],
        dropna=False
    )[col_descripcion]
    .first()
    .reset_index()
)

base_busqueda = base_busqueda.merge(
    emitido,
    how="left",
    on=["ENTIDAD", "CLAVE_MERGE"]
)

base_busqueda = base_busqueda.merge(
    entregado,
    how="left",
    on=["ENTIDAD", "CLAVE_MERGE"]
)

base_busqueda = base_busqueda.merge(
    descripcion,
    how="left",
    on=["ENTIDAD", "CLAVE_MERGE"]
)
base_busqueda = base_busqueda.merge(
    df[
        ["CLAVE_MERGE", "CLAVE"]
    ].drop_duplicates(),
    how="left",
    on="CLAVE_MERGE"
)
base_busqueda[col_emitidas] = (
    base_busqueda[col_emitidas]
    .fillna(0)
)

base_busqueda[col_entregadas] = (
    base_busqueda[col_entregadas]
    .fillna(0)
)


# =========================
# FORMATOS
# =========================
def fmt(x):
    return f"{int(x):,}"

def fmt_money(x):
    return f"${int(x):,}"

def formatear_clave(clave):

    clave = str(clave)

    clave = (
        clave
        .replace(".", "")
        .replace(" ", "")
        .replace("-", "")
    )

    if not clave.isdigit():
        return clave

    # 3-3-4-2
    if len(clave) == 12:

        return (
            clave[:3] + "." +
            clave[3:6] + "." +
            clave[6:10] + "." +
            clave[10:]
        )

    # 3-3-4
    elif len(clave) == 10:

        return (
            clave[:3] + "." +
            clave[3:6] + "." +
            clave[6:]
        )

    # 3-3-3
    elif len(clave) == 9:

        return (
            clave[:3] + "." +
            clave[3:6] + "." +
            clave[6:]
        )

    return clave
# =========================
# MÉTRICAS
# =========================
def calcular_metricas(df_estado):

    df_estado = df_estado.copy()
    # =========================
    # ABASTO
    # =========================

    df_estado["PIEZAS_INV"] = pd.to_numeric(
        df_estado["PIEZAS_INV"],
        errors="coerce"
    ).fillna(0)

    df_estado["CPM"] = pd.to_numeric(
        df_estado["CPM"],
        errors="coerce"
    ).fillna(0)

    df_estado["NIVEL_ABASTO"] = 0.0

    mask = df_estado["CPM"] > 0

    df_estado.loc[mask, "NIVEL_ABASTO"] = (
        df_estado.loc[mask, "PIEZAS_INV"] /
        df_estado.loc[mask, "CPM"]
    )

    # CLAVES CON INVENTARIO Y CPM 0
    df_estado.loc[
        (df_estado["PIEZAS_INV"] > 0) &
        (df_estado["CPM"] == 0),
        "NIVEL_ABASTO"
    ] = 2
    # =========================
    # CLASIFICACIÓN
    # =========================

    def clasificar(x):

        if x == 0:
            return "Agotado"

        elif x < 1:
            return "Próx agotarse"

        elif x <= 1.5:
            return "Bajo"

        elif x <= 5:
            return "Óptimo"

        else:
            return "Sobre stock"

    df_estado["CLASIFICACION"] = (
        df_estado["NIVEL_ABASTO"]
        .apply(clasificar)
    )

    df_estado["Emitido"] = pd.to_numeric(
        df_estado[col_emitidas],
        errors="coerce"
    ).fillna(0)

    df_estado["Entregado"] = pd.to_numeric(
        df_estado[col_entregadas],
        errors="coerce"
    ).fillna(0)

    df_estado[col_precio] = pd.to_numeric(
        df_estado[col_precio],
        errors="coerce"
    ).fillna(0)
    df_estado[col_precio] = pd.to_numeric(
        df_estado[col_precio],
        errors="coerce"
    ).fillna(0)

    df_estado["Transito"] = (
        df_estado["Emitido"] - df_estado["Entregado"]
    ).clip(lower=0)

    df_estado["Monto_Emitido"] = df_estado["Emitido"] * df_estado[col_precio]
    df_estado["Monto_Entregado"] = df_estado["Entregado"] * df_estado[col_precio]
    df_estado["Monto_Transito"] = df_estado["Monto_Emitido"] - df_estado["Monto_Entregado"]

    p = df_estado["Emitido"].sum()
    e = df_estado["Entregado"].sum()
    t = df_estado["Transito"].sum()

    m_p = df_estado["Monto_Emitido"].sum()
    m_e = df_estado["Monto_Entregado"].sum()
    m_t = df_estado["Monto_Transito"].sum()

    o_total = df_estado[col_orden].nunique()
    c_total = df_estado[col_clave].nunique()

    o_ent = df_estado[df_estado["Entregado"] >= df_estado["Emitido"]][col_orden].nunique()
    o_tran = df_estado[df_estado["Entregado"] < df_estado["Emitido"]][col_orden].nunique()

    c_ent = df_estado[df_estado["Entregado"] >= df_estado["Emitido"]][col_clave].nunique()
    c_tran = df_estado[df_estado["Entregado"] < df_estado["Emitido"]][col_clave].nunique()

    inv_total = (
    df_estado[
        ["ENTIDAD", "CLAVE", "PIEZAS_INV"]
    ]
    .drop_duplicates()
    ["PIEZAS_INV"]
    .sum()
)
    cpm_total = (
    df_estado[
        ["ENTIDAD", "CLAVE", "CPM"]
    ]
    .drop_duplicates()
    ["CPM"]
    .sum()
)

    agotado = (
        df_estado["CLASIFICACION"]
        == "Agotado"
    ).sum()

    proximo = (
        df_estado["CLASIFICACION"]
        == "Próx agotarse"
    ).sum()

    bajo = (
        df_estado["CLASIFICACION"]
        == "Bajo"
    ).sum()

    optimo = (
        df_estado["CLASIFICACION"]
        == "Óptimo"
    ).sum()

    sobre = (
        df_estado["CLASIFICACION"]
        == "Sobre stock"
    ).sum()

    nivel_promedio = (
        df_estado["NIVEL_ABASTO"]
        .mean()
    )    
    tabla = pd.DataFrame([
        ["Piezas", fmt(p), fmt(t), fmt(e)],
        ["Claves", fmt(c_total), fmt(c_tran), fmt(c_ent)],
        ["Órdenes", fmt(o_total), fmt(o_tran), fmt(o_ent)],
        ["Montos", fmt_money(m_p), fmt_money(m_t), fmt_money(m_e)]
    ], columns=["Métrica", "Emitido", "En tránsito", "Entregado"])

    return {
        "inv_total": inv_total,
        "cpm_total": cpm_total,
        "agotado": agotado,
        "proximo": proximo,
        "bajo": bajo,
        "optimo": optimo,
        "sobre": sobre,
        "nivel_promedio": nivel_promedio,
        "p": p,
        "e": e,
        "t": t,
        "m_p": m_p,
        "m_e": m_e,
        "m_t": m_t,
        "o_total": o_total,
        "o_ent": o_ent,
        "o_tran": o_tran,
        "c_total": c_total,
        "c_ent": c_ent,
        "c_tran": c_tran,
        "tabla": tabla
    }

# =========================
# GRÁFICAS
# =========================
def grafica(titulo, valores, labels, colores):

    fig, ax = plt.subplots(figsize=(10.5, 3.6))

    ax.barh(labels, valores, color=colores)

    ax.set_title(
        titulo,
        fontsize=15,
        fontweight="bold"
    )

    ax.set_xticks([])
    ax.tick_params(axis="y", labelsize=11)

    max_val = max(valores) if max(valores) > 0 else 1

    for i, v in enumerate(valores):
        ax.text(
            v + (max_val * 0.01),
            i,
            f"{int(v):,}",
            va="center",
            fontsize=11,
            fontweight="bold"
        )

    ax.set_xlim(0, max_val * 1.18)
    ax.set_facecolor("none")
    fig.patch.set_alpha(0)

    plt.tight_layout()

    return fig

def crear_graficas(m):

    fig_piezas = grafica(
        "Piezas",
        [m["p"], m["e"], m["t"]],
        ["Emitido", "Entregado", "Tránsito"],
        ["#235B4E", "#9F2241", "#B38E5D"]
    )

    fig_montos = grafica(
        "Montos ($)",
        [m["m_p"], m["m_e"], m["m_t"]],
        ["Emitido", "Entregado", "Tránsito"],
        ["#235B4E", "#9F2241", "#B38E5D"]
    )

    fig_ordenes = grafica(
        "Órdenes",
        [m["o_total"], m["o_ent"], m["o_tran"]],
        ["Total", "Entregadas", "En tránsito"],
        ["#235B4E", "#9F2241", "#B38E5D"]
    )

    fig_claves = grafica(
        "Claves",
        [m["c_total"], m["c_ent"], m["c_tran"]],
        ["Total", "Entregadas", "En tránsito"],
        ["#235B4E", "#9F2241", "#B38E5D"]
    )

    return fig_piezas, fig_montos, fig_ordenes, fig_claves

if menu == "🏠 Dashboard Nacional":
	
    # =========================
    # RANKING NACIONAL
    # =========================

    st.markdown("## 🏆 Ranking nacional de estados")

    concurrentes = [
        "BAJA CALIFORNIA",
        "BAJA CALIFORNIA SUR",
        "CAMPECHE",
        "CHIAPAS",
        "CIUDAD DE MEXICO",
        "COLIMA",
        "GUERRERO",
        "HIDALGO",
        "MEXICO",
        "MICHOACAN",
        "MORELOS",
        "NAYARIT",
        "OAXACA",
        "PUEBLA",
        "QUINTANA ROO",
        "SAN LUIS POTOSI",
        "SINALOA",
        "SONORA",
        "TABASCO",
        "TAMAULIPAS",
        "TLAXCALA",
        "VERACRUZ",
        "YUCATAN",
        "ZACATECAS"
    ]

    no_concurrentes = [
        "AGUASCALIENTES",
        "CHIHUAHUA",
        "COAHUILA DE ZARAGOZA",
        "DURANGO",
        "GUANAJUATO",
        "JALISCO",
        "NUEVO LEON",
        "QUERETARO DE ARTEAGA"
    ]

    tipo_ranking = st.selectbox(
        "Tipo de entidad",
        [
            "Todos",
            "Concurrentes",
            "No concurrentes"
        ]
    )

    ranking = df.copy()

    ranking["NIVEL_ABASTO"] = 0.0

    mask = ranking["CPM"] > 0

    ranking.loc[mask, "NIVEL_ABASTO"] = (
        ranking["PIEZAS_INV"] /
        ranking["CPM"]
    )

    ranking.loc[
        (ranking["PIEZAS_INV"] > 0) &
        (ranking["CPM"] == 0),
        "NIVEL_ABASTO"
    ] = 2

    # =========================
    # COBERTURA REAL ESTATAL
    # =========================

    ranking_claves = (
        ranking.groupby(
            ["ENTIDAD", "CLAVE_MERGE"],
            dropna=False
        )
        .agg({
            "PIEZAS_INV": "max",
            "CPM": "max"
        })
        .reset_index()
    )

    ranking_claves["INV_VALIDO"] = (
        ranking_claves.apply(
            lambda x: x["PIEZAS_INV"]
            if x["CPM"] > 0 else 0,
            axis=1
        )
    )

    ranking_estados = (
        ranking_claves.groupby("ENTIDAD")
        .agg({
            "INV_VALIDO": "sum",
            "CPM": "sum"
        })
        .reset_index()
    )

    ranking_estados["NIVEL_ABASTO"] = (
        ranking_estados["INV_VALIDO"] /
        ranking_estados["CPM"]
    )

    ranking_estados = ranking_estados.rename(columns={
        "ENTIDAD": "Estado",
        "INV_VALIDO": "Inventario",
        "CPM": "CPM",
        "NIVEL_ABASTO": "Nivel abasto"
    })

    # =========================
    # FILTROS
    # =========================

    if tipo_ranking == "Concurrentes":

        ranking_estados = ranking_estados[
            ranking_estados["Estado"]
            .isin(concurrentes)
        ]

    elif tipo_ranking == "No concurrentes":

        ranking_estados = ranking_estados[
            ranking_estados["Estado"]
            .isin(no_concurrentes)
        ]

    # =========================
    # ORDENAR
    # =========================

    ranking_estados = ranking_estados.sort_values(
        "Nivel abasto"
    )

    # =========================
    # FORMATO VISUAL
    # =========================

    ranking_estados["Inventario"] = (
        ranking_estados["Inventario"]
        .fillna(0)
        .apply(lambda x: f"{int(x):,} piezas")
    )

    ranking_estados["CPM"] = (
        ranking_estados["CPM"]
        .fillna(0)
        .apply(lambda x: f"{int(x):,} piezas")
    )

    ranking_estados["Nivel abasto"] = (
        ranking_estados["Nivel abasto"]
        .fillna(0)
        .round(2)
        .apply(lambda x: f"{x} meses")
    )

    st.dataframe(
        ranking_estados,
        use_container_width=True,
        hide_index=True
    )
    # =========================
    # INVENTARIO SIN CPM
    # =========================

    st.markdown("## 📦 Inventario sin CPM")

    sin_cpm_estados = (
        ranking_claves[
            (ranking_claves["CPM"] == 0) &
            (ranking_claves["PIEZAS_INV"] > 0)
        ]
        .groupby("ENTIDAD")
        .agg({
            "PIEZAS_INV": "sum",
            "CLAVE_MERGE": "count"
        })
        .reset_index()
    )
    sin_cpm_estados = sin_cpm_estados.rename(columns={
        "ENTIDAD": "Estado",
        "PIEZAS_INV": "Inventario",
        "CLAVE_MERGE": "Claves sin CPM"
    })
    sin_cpm_estados["Inventario"] = (
        sin_cpm_estados["Inventario"]
        .fillna(0)
        .apply(lambda x: f"{int(x):,} piezas")
    )
    sin_cpm_estados["Claves sin CPM"] = (
        sin_cpm_estados["Claves sin CPM"]
        .fillna(0)
        .apply(lambda x: f"{int(x):,} claves")
    )

    st.dataframe(
        sin_cpm_estados,
        use_container_width=True,
        hide_index=True
    )
    # =========================
    # RESUMEN GENERAL NACIONAL
    # =========================

    metricas_general = calcular_metricas(df)

    st.markdown("## 🇲🇽 Resumen general nacional")

    g1, g2, g3 = st.columns(3)
    g4, g5, g6 = st.columns(3)

    with g1:

        st.markdown(f"""
<div class="kpi-card kpi-blue">

<div class="kpi-title">
📦 Emitido
</div>

<div class="kpi-value">
{fmt(metricas_general["p"])}
</div>

</div>
""", unsafe_allow_html=True)

    with g2:

        st.markdown(f"""
<div class="kpi-card kpi-green">

<div class="kpi-title">
🚚 Entregado
</div>

<div class="kpi-value">
{fmt(metricas_general["e"])}
</div>

</div>
""", unsafe_allow_html=True)

    with g3:

        st.markdown(f"""
<div class="kpi-card kpi-yellow">

<div class="kpi-title">
⏳ En tránsito
</div>

<div class="kpi-value">
{fmt(metricas_general["t"])}
</div>

</div>
""", unsafe_allow_html=True)

    with g4:

        st.markdown(f"""
<div class="kpi-card">

<div class="kpi-title">
📈 CPM
</div>

<div class="kpi-value">
{fmt(metricas_general["cpm_total"])}
</div>

</div>
""", unsafe_allow_html=True)

    with g5:

        st.markdown(f"""
<div class="kpi-card kpi-red">

<div class="kpi-title">
🏥 Inventario
</div>

<div class="kpi-value">
{fmt(metricas_general["inv_total"])}
</div>

</div>
""", unsafe_allow_html=True)

    st.markdown("### 📋 Tabla nacional")

    col_n1, col_n2, col_n3 = st.columns([1, 3, 1])

    with col_n2:

        st.dataframe(
            metricas_general["tabla"],
            use_container_width=True,
            hide_index=True
        )

    st.divider()
if menu == "🔎 Buscar Clave":

    # =========================
    # BUSCADOR DE CLAVE
    # =========================

    st.markdown("## 🔎 Buscar clave nacional")

    clave_busqueda = st.text_input(
        "Ingresar clave"
    )

    clave_busqueda = (
        clave_busqueda
        .strip()
    )

    if clave_busqueda:

        df_busqueda = base_busqueda.copy()

        df_busqueda = df_busqueda[
            df_busqueda["CLAVE"]
            .astype(str)
            .str.contains(
                clave_busqueda,
                case=False,
                na=False
            )
        ]

        if len(df_busqueda) > 0:

            df_busqueda["NIVEL_ABASTO"] = 0.0

            mask = df_busqueda["CPM"] > 0

            df_busqueda.loc[mask, "NIVEL_ABASTO"] = (
                df_busqueda["PIEZAS_INV"] /
                df_busqueda["CPM"]
            )

            df_busqueda.loc[
                (df_busqueda["PIEZAS_INV"] > 0) &
                (df_busqueda["CPM"] == 0),
                "NIVEL_ABASTO"
            ] = 2

            def clasificar_busqueda(x):

                if x == 0:
                    return "Agotado"

                elif x < 1:
                    return "Próx agotarse"

                elif x <= 1.5:
                    return "Bajo"

                elif x <= 5:
                    return "Óptimo"

                else:
                    return "Sobre stock"

            df_busqueda["Clasificación"] = (
                df_busqueda["NIVEL_ABASTO"]
                .apply(clasificar_busqueda)
            )

            tabla_busqueda = df_busqueda.copy()

            tabla_busqueda["Tránsito"] = (
                tabla_busqueda[col_emitidas] -
                tabla_busqueda[col_entregadas]
            ).clip(lower=0)

            tabla_busqueda = tabla_busqueda.rename(columns={
                "ENTIDAD": "Estado",
                "CLAVE": "Clave",
                "PIEZAS_INV": "Inventario",
                "CPM": "CPM",
                "NIVEL_ABASTO": "Nivel",
                "Clasificación": "Clasificación",
                col_entregadas: "Entregado",
                col_emitidas: "Emitido"
            })

            tabla_busqueda["Nivel"] = (
                tabla_busqueda["Nivel"]
                .round(2)
            )

            descripcion = (
                str(df_busqueda.iloc[0][df.columns[3]])
            )

            st.markdown(
                f"### 🔎 Clave encontrada: {descripcion}"
            )

            # =========================
            # REDISTRIBUCIÓN SUGERIDA
            # =========================

            st.markdown("## 🔄 Redistribución sugerida")

            receptores = (
                tabla_busqueda[
                    tabla_busqueda["Nivel"] < 1
                ]
                .sort_values("Nivel")
            )

            donadores = (
                tabla_busqueda[
                    tabla_busqueda["Nivel"] > 5
                ]
                .sort_values(
                    "Nivel",
                    ascending=False
                )
            )

            sugerencias = []

            total = min(
                len(receptores),
                len(donadores)
            )

            for i in range(total):

                recibe = receptores.iloc[i]
                dona = donadores.iloc[i]

                sugerencias.append({
                    "Recibe": recibe["Estado"],
                    "Nivel receptor": round(
                        recibe["Nivel"],
                        2
                    ),
                    "Dona": dona["Estado"],
                    "Nivel donador": round(
                        dona["Nivel"],
                        2
                    )
                })

            sugerencias_df = pd.DataFrame(
                sugerencias
            )

            if len(sugerencias_df) > 0:

                st.dataframe(
                    sugerencias_df,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.success(
                    "No se detectaron redistribuciones sugeridas."
                )

            tabla_busqueda = tabla_busqueda.sort_values(
                "Inventario",
                ascending=False
            )

            tabla_busqueda = tabla_busqueda.drop(
                columns=["CLAVE_MERGE"],
                errors="ignore"
            )

            st.dataframe(
                tabla_busqueda,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.error(
                "No se encontró la clave."
            )
if menu == "📍 Estados":

    # =========================
    # FILTRO
    # =========================

    estados_unicos = (
        df[col_estado]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
        .dropna()
        .unique()
    )

    estado_sel = st.selectbox(
        "📍 Estado",
        sorted(estados_unicos)
    )
    df_f = df[df[col_estado] == estado_sel].copy()
    # =========================
    # MÉTRICAS EJECUTIVAS
    # =========================

    base_estado = (
        df_f.groupby(
            ["ENTIDAD", "CLAVE_MERGE"],
            dropna=False
        )
        .agg({
            "PIEZAS_INV": "max",
            "CPM": "max"
        })
        .reset_index()
    )

    inventario_total = (
        base_estado["PIEZAS_INV"]
        .sum()
    )

    inventario_sin_cpm = (
        base_estado[
            (base_estado["CPM"] == 0) &
            (base_estado["PIEZAS_INV"] > 0)
        ]["PIEZAS_INV"]
        .sum()
    )

    inventario_con_cpm = (
        base_estado[
            base_estado["CPM"] > 0
        ]["PIEZAS_INV"]
        .sum()
    )

    cpm_total_estado = (
        base_estado["CPM"]
        .sum()
    )

    nivel_estado = 0

    if cpm_total_estado > 0:

        nivel_estado = (
            inventario_con_cpm /
            cpm_total_estado
        )

    e1, e2, e3, e4 = st.columns(4)

    with e1:

        st.markdown(f"""
        <div class="kpi-card kpi-blue">

        <div class="kpi-title">
        📈 Nivel de abasto
        </div>

        <div class="kpi-value">
        {nivel_estado:.2f} meses
        </div>

        </div>
        """, unsafe_allow_html=True)

    with e2:

        st.markdown(f"""
        <div class="kpi-card kpi-green">

        <div class="kpi-title">
        📦 Inventario total
        </div>

        <div class="kpi-value">
        {inventario_total:,.0f}
        </div>

        </div>
        """, unsafe_allow_html=True)
    with e3:

        st.markdown(f"""
        <div class="kpi-card kpi-yellow">

        <div class="kpi-title">
        📦 Inventario con CPM
        </div>

        <div class="kpi-value">
        {inventario_con_cpm:,.0f}
        </div>

        </div>
        """, unsafe_allow_html=True)

    with e4:

        st.markdown(f"""
        <div class="kpi-card kpi-red">

        <div class="kpi-title">
        🚨 Inventario sin CPM
        </div>

        <div class="kpi-value">
        {inventario_sin_cpm:,.0f}
        </div>

        </div>
        """, unsafe_allow_html=True)


    # =========================
    # RESUMEN ABASTO
    # =========================

    df_f = df_f.copy()

    df_f["NIVEL_ABASTO"] = 0.0

    mask = df_f["CPM"] > 0

    df_f.loc[mask, "NIVEL_ABASTO"] = (
        df_f["PIEZAS_INV"] /
        df_f["CPM"]
    )

    df_f.loc[
        (df_f["PIEZAS_INV"] > 0) &
        (df_f["CPM"] == 0),
        "NIVEL_ABASTO"
    ] = 2

    def clasificar_tabla(x):

        if x == 0:
            return "Agotado"

        elif x < 1:
            return "Próx agotarse"

        elif x <= 1.5:
            return "Bajo"

        elif x <= 5:
            return "Óptimo"

        else:
            return "Sobre stock"

    df_f["CLASIFICACION"] = (
        df_f["NIVEL_ABASTO"]
        .apply(clasificar_tabla)
    )

    abasto_claves = (
        df_f.groupby(
            ["CLAVE_MERGE"],
            dropna=False
        )
        .agg({
            col_clave: "first",
            "PIEZAS_INV": "max",
            "CPM": "max",
            col_emitidas: "sum",
            col_entregadas: "sum",
            "NIVEL_ABASTO": "first",
            "CLASIFICACION": "first"
        })
        .reset_index()
    )

    abasto_claves.columns = [
        "CLAVE_MERGE",
        "Clave",
        "Inventario",
        "CPM",
        "Emitido",
        "Entregado",
        "Nivel",
        "Clasificación"
    ]
    abasto_claves["Tránsito"] = (
        abasto_claves["Emitido"] -
        abasto_claves["Entregado"]
    ).clip(lower=0)
    abasto_claves = abasto_claves[
        [
            "Clave",
            "Inventario",
            "CPM",
            "Emitido",
            "Tránsito",
            "Entregado",
            "Nivel",
            "Clasificación"
        ]
    ]
    abasto_claves = abasto_claves.drop(
        columns=["CLAVE_MERGE"],
        errors="ignore"
    )

    abasto_claves["Nivel"] = (
        abasto_claves["Nivel"]
        .round(2)
    )

    abasto_claves = abasto_claves.sort_values(
        "Nivel"
    )

    metricas = calcular_metricas(df_f)

    tabla_base = metricas["tabla"]

    fig_piezas, fig_montos, fig_ordenes, fig_claves = crear_graficas(metricas)

    # =========================
    # TABLA STREAMLIT
    # =========================

    st.markdown(f"## 📋 Tabla operativa - {estado_sel}")

    col1, col2, col3 = st.columns([1, 3, 1])

    with col2:

        st.dataframe(
            tabla_base,
            use_container_width=True,
            hide_index=True
        )

    # =========================
    # VISUALIZACIÓN
    # =========================

    resumen_abasto = (
        abasto_claves["Clasificación"]
        .value_counts()
        .reset_index()
    )

    resumen_abasto.columns = [
        "Clasificación",
        "Total claves"
    ]
    # =========================
    # EXPORTAR EXCEL ESTATAL
    # =========================

    archivo_estado = f"Analisis_{estado_sel}.xlsx"

    abasto_export = abasto_claves.copy()

    with pd.ExcelWriter(
        archivo_estado,
        engine="openpyxl"
    ) as writer:

        abasto_export.to_excel(
            writer,
            index=False,
            startrow=1,
            sheet_name="Análisis"
        )

        wb = writer.book
        ws = writer.sheets["Análisis"]
        from openpyxl.styles import (
            Font,
            PatternFill,
            Border,
            Side,
            Alignment
        )
        # =========================
        # TITULO PRINCIPAL
        # =========================

        ws.merge_cells("A1:H1")

        ws["A1"] = f"Análisis estatal - {estado_sel}"

        ws["A1"].font = Font(
            name="Noto Sans",
            size=18,
            bold=True,
            color="FFFFFF"
        )

        ws["A1"].fill = PatternFill(
            start_color="235B4E",
            end_color="235B4E",
            fill_type="solid"
        )

        ws["A1"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws.row_dimensions[1].height = 28

        # =========================
        # BAJAR TABLA
        # =========================


        # =========================
        # ESTILOS
        # =========================

        verde = "235B4E"

        fill_header = PatternFill(
            start_color=verde,
            end_color=verde,
            fill_type="solid"
        )

        font_header = Font(
            color="FFFFFF",
            bold=True,
            name="Noto Sans",
            size=11
        )

        font_body = Font(
            name="Noto Sans",
            size=10
        )

        thin = Side(
            border_style="thin",
            color="000000"
        )

        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        # =========================
        # RESUMEN EJECUTIVO
        # =========================

        # Separación visual
        ws.column_dimensions["I"].width = 28
        ws.column_dimensions["J"].width = 30
        ws.column_dimensions["K"].width = 22

        # Título
        ws.merge_cells("J2:K2")

        ws["J2"] = "Resumen ejecutivo"

        ws["J2"].font = Font(
            name="Noto Sans",
            size=14,
            bold=True,
            color="FFFFFF"
        )

        ws["J2"].fill = PatternFill(
            start_color="235B4E",
            end_color="235B4E",
            fill_type="solid"
        )

        ws["J2"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        resumen = [
            ["Nivel abasto", f"{nivel_estado:.2f} meses"],
            ["Inventario total", f"{inventario_total:,.0f}"],
            ["Inventario con CPM", f"{inventario_con_cpm:,.0f}"],
            ["Inventario sin CPM", f"{inventario_sin_cpm:,.0f}"]
        ]

        fila = 3

        for titulo, valor in resumen:

            ws[f"J{fila}"] = titulo
            ws[f"K{fila}"] = valor

            # TITULO
            ws[f"J{fila}"].font = Font(
                bold=True,
                name="Noto Sans",
                color="FFFFFF"
            )

            ws[f"J{fila}"].fill = PatternFill(
                start_color="9F2241",
                end_color="9F2241",
                fill_type="solid"
            )

            # VALOR
            ws[f"K{fila}"].font = Font(
                bold=True,
                name="Noto Sans"
            )

            # BORDES
            ws[f"J{fila}"].border = border
            ws[f"K{fila}"].border = border

            # ALINEACIÓN
            ws[f"J{fila}"].alignment = alignment
            ws[f"K{fila}"].alignment = alignment

            fila += 1
        # =========================
        # ENCABEZADOS
        # =========================

        for cell in ws[2]:

            cell.fill = fill_header
            cell.font = font_header
            cell.border = border
            cell.alignment = alignment

        # =========================
        # CUERPO
        # =========================

        for row in ws.iter_rows(
            min_row=3,
            max_col=8
        ):

            for cell in row:

                cell.font = font_body
                cell.border = border
                cell.alignment = alignment
        # =========================
        # QUITAR BORDE DERECHO
        # =========================

        for cell in ws["H"]:

            cell.border = Border(
                left=thin,
                top=thin,
                bottom=thin
            )

        # =========================
        # FORMATOS NUMÉRICOS
        # =========================

        columnas_numericas = [
            "B",
            "C",
            "D",
            "E",
            "F"
        ]

        for col in columnas_numericas:

            for cell in ws[col][1:]:

                cell.number_format = '#,##0'

        # =========================
        # AUTOAJUSTE
        # =========================

        from openpyxl.utils import get_column_letter

        for i, col in enumerate(ws.columns, 1):

            max_length = 0

            column = get_column_letter(i)

            for cell in col:

                try:

                    if cell.value is not None:

                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )

                except:
                    pass

            adjusted_width = max_length + 5

            ws.column_dimensions[column].width = adjusted_width
        # =========================
        # FILTRO
        # =========================

        ws.auto_filter.ref = "A2:H2"

        # =========================
        # CONGELAR
        # =========================

        ws.freeze_panes = "A3"
    with open(archivo_estado, "rb") as f:

        st.download_button(
            "⬇ Descargar análisis estatal",
            f,
            file_name=archivo_estado
        )
    st.markdown(f"## 🏥 Resumen de abasto - {estado_sel}")

    st.dataframe(
        resumen_abasto,
        use_container_width=True,
        hide_index=True
    )

    st.markdown(f"## 📊 Visualización - {estado_sel}")

    c1, c2 = st.columns(2)
    c3, c4 = st.columns(2)

    with c1:
        st.pyplot(fig_piezas)

    with c2:
        st.pyplot(fig_montos)

    with c3:
        st.pyplot(fig_ordenes)

    with c4:
        st.pyplot(fig_claves)
if menu == "🚨 Alertas":

    st.markdown("## 🚨 Centro de alertas")

    alertas = base_busqueda.copy()

    alertas = (
        alertas.groupby(
            ["CLAVE_MERGE"],
            dropna=False
        )
    .agg({
        "CLAVE": lambda x: (
            x[
                x.astype(str)
                .str.contains(r"\.")
            ].iloc[0]
            if len(
                x[
                    x.astype(str)
                    .str.contains(r"\.")
                ]
            ) > 0
            else (
                x.dropna().iloc[0]
                if len(x.dropna()) > 0
                else None
            )
        ),
            col_descripcion: "first",
            "PIEZAS_INV": "sum",
            "CPM": "sum"
        })
        .reset_index()
    )

    alertas["CLAVE"] = (
        alertas["CLAVE"]
        .fillna(
            alertas["CLAVE_MERGE"]
        )
        .apply(formatear_clave)
    )

    alertas[col_descripcion] = (
        alertas[col_descripcion]
        .fillna("SIN DESCRIPCIÓN")
    )

    alertas["NIVEL_ABASTO"] = 0.0

    mask = alertas["CPM"] > 0

    alertas.loc[mask, "NIVEL_ABASTO"] = (
        alertas["PIEZAS_INV"] /
        alertas["CPM"]
    )

    alertas.loc[
        (alertas["PIEZAS_INV"] > 0) &
        (alertas["CPM"] == 0),
        "NIVEL_ABASTO"
    ] = 2

    # =========================
    # ALERTAS
    # =========================

    agotado = alertas[
        alertas["NIVEL_ABASTO"] == 0
    ]

    proximo = alertas[
        (alertas["NIVEL_ABASTO"] > 0) &
        (alertas["NIVEL_ABASTO"] < 1)
    ]

    sobrestock = alertas[
        alertas["NIVEL_ABASTO"] > 5
    ]

    sin_cpm = alertas[
        (alertas["CPM"] == 0) &
        (alertas["PIEZAS_INV"] > 0)
    ]

    # =========================
    # KPIs ALERTAS
    # =========================

    a1, a2, a3, a4 = st.columns(4)

    with a1:

        st.markdown(f"""
        <div class="kpi-card kpi-red">

        <div class="kpi-title">
        🔴 Agotado
        </div>

        <div class="kpi-value">
        {len(agotado):,}
        </div>

        </div>
        """, unsafe_allow_html=True)

    with a2:

        st.markdown(f"""
        <div class="kpi-card kpi-yellow">

        <div class="kpi-title">
        🟠 Próximo agotarse
        </div>

        <div class="kpi-value">
        {len(proximo):,}
        </div>

        </div>
        """, unsafe_allow_html=True)

    with a3:

        st.markdown(f"""
        <div class="kpi-card kpi-blue">

        <div class="kpi-title">
        🔵 Sobre stock
        </div>

        <div class="kpi-value">
        {len(sobrestock):,}
        </div>

        </div>
        """, unsafe_allow_html=True)

    with a4:

        st.markdown(f"""
        <div class="kpi-card kpi-green">

        <div class="kpi-title">
        🟢 Sin CPM
        </div>

        <div class="kpi-value">
        {len(sin_cpm):,}
        </div>

        </div>
        """, unsafe_allow_html=True)
    # =========================
    # TABLAS ALERTAS
    # =========================

    st.markdown("### 🔴 Claves agotadas")

    st.dataframe(
        agotado[
            [
                "CLAVE",
                col_descripcion,
                "PIEZAS_INV",
                "CPM"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )
    # =========================
    # EXPORTAR ALERTAS
    # =========================

    excel_alertas = agotado.copy()

    excel_alertas = excel_alertas[
        [
            "CLAVE",
            col_descripcion,
            "PIEZAS_INV",
            "CPM",
            "NIVEL_ABASTO"
        ]
    ]

    archivo_excel = "alertas_criticas.xlsx"

    excel_alertas.to_excel(
        archivo_excel,
        index=False
    )

    with open(archivo_excel, "rb") as f:

        st.download_button(
            "⬇ Descargar alertas críticas",
            f,
            file_name=archivo_excel
        )
# =========================
# ANALÍTICA NACIONAL
# =========================

if menu == "📊 Analítica":

    st.markdown(
        "## 📊 Centro analítico nacional"
    )

    # =========================
    # BASE ANALÍTICA NACIONAL
    # =========================

    analitica = (
        base_busqueda.groupby(
            ["CLAVE_MERGE"],
            dropna=False
        )
    .agg({
        "CLAVE": lambda x: (
            x[
                x.astype(str)
                .str.contains(r"\.")
            ].iloc[0]
            if len(
                x[
                    x.astype(str)
                    .str.contains(r"\.")
                ]
            ) > 0
            else (
                x.dropna().iloc[0]
                if len(x.dropna()) > 0
                else None
            )
        ),  
            col_descripcion: "first",
            "PIEZAS_INV": "sum",
            "CPM": "sum"
        })
        .reset_index()
    )

    analitica["CLAVE"] = (
        analitica["CLAVE"]
        .fillna(
            analitica["CLAVE_MERGE"]
        )
        .apply(formatear_clave)
    )
    analitica[col_descripcion] = (
        analitica[col_descripcion]
        .fillna("SIN DESCRIPCIÓN")
    )

    analitica["NIVEL_ABASTO"] = 0.0

    mask = analitica["CPM"] > 0

    analitica.loc[mask, "NIVEL_ABASTO"] = (
        analitica["PIEZAS_INV"] /
        analitica["CPM"]
    )

    analitica.loc[
        (analitica["PIEZAS_INV"] > 0) &
        (analitica["CPM"] == 0),
        "NIVEL_ABASTO"
    ] = 2

    # =========================
    # TOP AGOTADAS CRÍTICAS
    # =========================

    st.markdown(
        "## 🔴 Top claves agotadas críticas"
    )

    agotadas = analitica[
        (analitica["PIEZAS_INV"] == 0) &
        (analitica["CPM"] > 0)
    ].copy()

    agotadas = agotadas.sort_values(
        "CPM",
        ascending=False
    )

    agotadas = agotadas.rename(columns={
        "CLAVE": "Clave",
        col_descripcion: "Descripción",
        "PIEZAS_INV": "Inventario",
        "CPM": "CPM"
    })

    st.dataframe(
        agotadas[
            [
                "Clave",
                "Descripción",
                "Inventario",
                "CPM"
            ]
        ].head(20),
        use_container_width=True,
        hide_index=True
    )

    # =========================
    # TOP RIESGO
    # =========================

    st.markdown(
        "## 🟠 Top claves en riesgo"
    )

    riesgo = analitica[
        (analitica["NIVEL_ABASTO"] > 0) &
        (analitica["NIVEL_ABASTO"] < 1)
    ].copy()

    riesgo = riesgo.sort_values(
        "NIVEL_ABASTO"
    )

    riesgo = riesgo.rename(columns={
        "CLAVE": "Clave",
        col_descripcion: "Descripción",
        "PIEZAS_INV": "Inventario",
        "CPM": "CPM",
        "NIVEL_ABASTO": "Nivel"
    })

    st.dataframe(
        riesgo[
            [
                "Clave",
                "Descripción",
                "Inventario",
                "CPM",
                "Nivel"
            ]
        ].head(20),
        use_container_width=True,
        hide_index=True
    )

    # =========================
    # TOP SOBRE STOCK
    # =========================

    st.markdown(
        "## 🔵 Top claves sobre stock"
    )

    sobrestock = analitica[
        analitica["NIVEL_ABASTO"] > 5
    ].copy()

    sobrestock = sobrestock.sort_values(
        "NIVEL_ABASTO",
        ascending=False
    )

    sobrestock = sobrestock.rename(columns={
        "CLAVE": "Clave",
        col_descripcion: "Descripción",
        "PIEZAS_INV": "Inventario",
        "CPM": "CPM",
        "NIVEL_ABASTO": "Nivel"
    })

    st.dataframe(
        sobrestock[
            [
                "Clave",
                "Descripción",
                "Inventario",
                "CPM",
                "Nivel"
            ]
        ].head(20),
        use_container_width=True,
        hide_index=True
    )

    # =========================
    # MEJOR COBERTURA
    # =========================

    st.markdown(
        "## 🟢 Mejor cobertura nacional"
    )

    mejor = analitica[
        analitica["CPM"] > 0
    ].copy()

    mejor = mejor.sort_values(
        "NIVEL_ABASTO",
        ascending=False
    )

    mejor = mejor.rename(columns={
        "CLAVE": "Clave",
        col_descripcion: "Descripción",
        "PIEZAS_INV": "Inventario",
        "CPM": "CPM",
        "NIVEL_ABASTO": "Nivel"
    })

    st.dataframe(
        mejor[
            [
                "Clave",
                "Descripción",
                "Inventario",
                "CPM",
                "Nivel"
            ]
        ].head(20),
        use_container_width=True,
        hide_index=True
    )
# =========================
# POWERPOINT
# =========================
MACHOTE = "MACHOTE_PRESENTACION.pptx"

def obtener_layout(prs):
    if len(prs.slide_layouts) > 5:
        return prs.slide_layouts[5]
    elif len(prs.slide_layouts) > 0:
        return prs.slide_layouts[0]
    else:
        raise Exception("El PowerPoint no tiene layouts disponibles.")

def limpiar_slide(slide):
    for shape in list(slide.shapes):
        try:
            sp = shape._element
            sp.getparent().remove(sp)
        except Exception:
            pass

def agregar_contenido_slide(
    slide,
    tabla,
    fig_piezas,
    fig_montos,
    fig_ordenes,
    fig_claves,
    estado
):

    txBox = slide.shapes.add_textbox(
        Inches(0.6),
        Inches(0.2),
        Inches(8),
        Inches(0.5)
    )

    tf = txBox.text_frame
    tf.text = f"Abasto - {str(estado).title()}"

    p_title = tf.paragraphs[0]
    p_title.font.size = Pt(30)
    p_title.font.bold = True

    fig_tab, ax = plt.subplots(figsize=(10, 2))
    ax.axis("off")

    table = ax.table(
        cellText=tabla.values,
        colLabels=tabla.columns,
        loc="center",
        cellLoc="center"
    )

    table.scale(1.4, 1.5)

    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor("black")

        if r == 0:
            cell.set_facecolor("#235B4E")
            cell.set_text_props(
                color="white",
                weight="bold",
                size=10
            )
        else:
            cell.set_facecolor((1, 1, 1, 0))
            cell.set_text_props(size=10)

    fig_tab.patch.set_alpha(0)

    img_tab = io.BytesIO()

    plt.savefig(
        img_tab,
        format="png",
        bbox_inches="tight",
        transparent=True
    )

    img_tab.seek(0)

    plt.close(fig_tab)

    slide.shapes.add_picture(
        img_tab,
        Inches(1.75),
        Inches(0.75),
        width=Inches(9.7)
    )

    def add(fig, x, y):

        img = io.BytesIO()

        fig.savefig(
            img,
            format="png",
            bbox_inches="tight",
            transparent=True
        )

        img.seek(0)

        slide.shapes.add_picture(
            img,
            Inches(x),
            Inches(y),
            width=Inches(4.9)
        )

        plt.close(fig)

    add(fig_piezas, 0.55, 2.05)
    add(fig_montos, 6.55, 2.05)
    add(fig_ordenes, 0.55, 4.25)
    add(fig_claves, 6.55, 4.25)

def exportar_ppt_estado(
    tabla_base,
    fig_piezas,
    fig_montos,
    fig_ordenes,
    fig_claves,
    estado_sel
):

    prs = Presentation(MACHOTE) if os.path.exists(MACHOTE) else Presentation()

    if len(prs.slides) == 0:
        slide = prs.slides.add_slide(obtener_layout(prs))
    else:
        slide = prs.slides[0]

    limpiar_slide(slide)

    agregar_contenido_slide(
        slide,
        tabla_base,
        fig_piezas,
        fig_montos,
        fig_ordenes,
        fig_claves,
        estado_sel
    )

    salida = f"Dashboard_IMB_{estado_sel}.pptx"

    prs.save(salida)

    return salida

def exportar_ppt_todos_estados():

    prs = Presentation(MACHOTE) if os.path.exists(MACHOTE) else Presentation()

    while len(prs.slides._sldIdLst) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    estados = sorted(df[col_estado].dropna().unique())

    barra = st.progress(0)
    total = len(estados)

    for i, estado in enumerate(estados):

        df_estado = df[df[col_estado] == estado].copy()

        m = calcular_metricas(df_estado)

        fig_p, fig_m, fig_o, fig_c = crear_graficas(m)

        slide = prs.slides.add_slide(obtener_layout(prs))

        limpiar_slide(slide)

        agregar_contenido_slide(
            slide,
            m["tabla"],
            fig_p,
            fig_m,
            fig_o,
            fig_c,
            estado
        )

        barra.progress((i + 1) / total)

    salida = "Dashboard_Nacional_IMB_Todos_Estados.pptx"

    prs.save(salida)

    return salida

# =========================
# BOTONES
# =========================
st.markdown("## 📤 Exportación PPT")

col_exp1, col_exp2 = st.columns(2)

with col_exp1:

    if st.button("📊 Exportar estado seleccionado"):

        archivo_ppt = exportar_ppt_estado(
            tabla_base,
            fig_piezas,
            fig_montos,
            fig_ordenes,
            fig_claves,
            estado_sel
        )

        st.success("PPT del estado generado correctamente ✅")

        with open(archivo_ppt, "rb") as f:
            st.download_button(
                "⬇ Descargar PPT del estado",
                f,
                file_name=archivo_ppt
            )

with col_exp2:

    if st.button("🌎 Exportar TODOS los estados"):

        archivo_total = exportar_ppt_todos_estados()

        st.success("PowerPoint nacional generado correctamente ✅")

        with open(archivo_total, "rb") as f:
            st.download_button(
                "⬇ Descargar PowerPoint Nacional",
                f,
                file_name=archivo_total
            )