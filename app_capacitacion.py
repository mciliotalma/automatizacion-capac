# -*- coding: utf-8 -*-
"""
Created on Mon Mar  9 09:07:40 2026
import streamlit as st
import pandas as pd

# --------------------------------------------------
# CONFIGURACION DE PAGINA
# --------------------------------------------------

st.set_page_config(
    page_title="Control de Capacitaciones",
    page_icon="📊",
    layout="wide"
)

# --------------------------------------------------
# ESTILO CORPORATIVO TALMA
# --------------------------------------------------

st.markdown("""
<style>

.main {
    background-color: #f5f7fb;
}

.block-container{
    border:4px solid #003A8F;
    padding:2rem;
    border-radius:15px;
}

h1,h2,h3{
    color:#003A8F;
}

.metric-box{
    text-align:center;
    padding:20px;
    border-radius:12px;
    font-weight:bold;
}

.verde{
    background-color:#28a745;
    color:white;
}

.amarillo{
    background-color:#ffc107;
    color:black;
}

.rojo{
    background-color:#dc3545;
    color:white;
}

</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# TITULO
# --------------------------------------------------

st.title("📊 Control de Capacitaciones")

# --------------------------------------------------
# MENSAJE SOBRE NOMBRE DE HOJA
# --------------------------------------------------

st.info("⚠️ El archivo Excel debe contener una hoja llamada **'Acumulado Portal'**.")

# --------------------------------------------------
# SUBIR ARCHIVO
# --------------------------------------------------

archivo = st.file_uploader("📂 Cargar archivo Excel", type=["xlsx"])

if archivo:

    # --------------------------------------------------
    # VALIDAR HOJA
    # --------------------------------------------------

    hojas = pd.ExcelFile(archivo).sheet_names

    if "Acumulado Portal" not in hojas:
        st.error("❌ El archivo no contiene la hoja **'Acumulado Portal'**. Verifique el nombre.")
        st.stop()

    # --------------------------------------------------
    # LEER DATA
    # --------------------------------------------------

    df = pd.read_excel(archivo, sheet_name="Acumulado Portal", header=1)

    st.success("Archivo cargado correctamente.")

    # --------------------------------------------------
    # BUSCAR COLUMNAS DE ESTADO
    # --------------------------------------------------

    columnas_estado = [col for col in df.columns if "ESTADO" in str(col).upper()]

    vigentes = 0
    por_vencer = 0
    vencidos = 0

    for col in columnas_estado:
        vigentes += (df[col] == "VIGENTE").sum()
        por_vencer += (df[col] == "POR VENCER").sum()
        vencidos += (df[col] == "VENCIDO").sum()

    # --------------------------------------------------
    # RESUMEN
    # --------------------------------------------------

    st.subheader("Resumen de Capacitaciones")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            f"<div class='metric-box verde'>🟢 Vigentes<br><h2>{vigentes}</h2></div>",
            unsafe_allow_html=True
        )

    with c2:
        st.markdown(
            f"<div class='metric-box amarillo'>🟡 Por vencer<br><h2>{por_vencer}</h2></div>",
            unsafe_allow_html=True
        )

    with c3:
        st.markdown(
            f"<div class='metric-box rojo'>🔴 Vencidos<br><h2>{vencidos}</h2></div>",
            unsafe_allow_html=True
        )

    # --------------------------------------------------
    # TABLA
    # --------------------------------------------------

    st.subheader("Datos cargados")

    st.dataframe(df)

else:

    st.warning("Sube un archivo para comenzar.")
