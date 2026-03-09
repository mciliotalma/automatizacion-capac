# -*- coding: utf-8 -*-
"""
Reporte Profesional de Capacitaciones TALMA - Versión completa
"""

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill

# --------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# --------------------------------------------------

st.set_page_config(
    page_title="Reporte Capacitaciones TALMA",
    page_icon="📊",
    layout="wide"
)

# --------------------------------------------------
# ESTILO CORPORATIVO TALMA
# --------------------------------------------------

st.markdown("""
<style>
.block-container{
    border:4px solid #003A8F;
    padding:2rem;
    border-radius:15px;
    background-color:#f4f7fb;
}

h1,h2,h3{
    color:#003A8F;
}

.stButton>button, .stDownloadButton>button{
    background-color:#A7D129;
    color:#004C97;
    font-weight:bold;
    border-radius:8px;
}

.metric-box{
    text-align:center;
    padding:20px;
    border-radius:12px;
    font-weight:bold;
    font-size:18px;
}

.verde{background-color:#28a745;color:white;}
.amarillo{background-color:#ffc107;color:black;}
.rojo{background-color:#dc3545;color:white;}
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# HEADER
# --------------------------------------------------

st.title("📊 Reporte Profesional de Capacitaciones TALMA")
st.info("⚠️ El archivo Excel debe contener una hoja llamada **'Acumulado Portal'**")

# --------------------------------------------------
# SUBIR ARCHIVO
# --------------------------------------------------

uploaded_file = st.file_uploader(
    "📂 Arrastra o selecciona tu archivo Excel",
    type=["xlsx","xlsm"]
)

if uploaded_file:

    # --------------------------------------------------
    # LEER HOJA
    # --------------------------------------------------

    xls = pd.ExcelFile(uploaded_file)
    if "Acumulado Portal" not in xls.sheet_names:
        st.error("❌ La hoja del Excel debe llamarse **'Acumulado Portal'**")
        st.stop()

    df_raw = pd.read_excel(xls, sheet_name="Acumulado Portal", header=None)

    # --------------------------------------------------
    # IDENTIFICAR ENCABEZADOS DOBLE FILA
    # --------------------------------------------------

    header_curso = df_raw.iloc[0, 7:]  # fila 0, desde columna H
    header_sub = df_raw.iloc[1, 7:]    # fila 1, desde columna H

    columnas = []
    for c1, c2 in zip(header_curso, header_sub):
        if pd.isna(c1):
            columnas.append(c2)
        elif pd.isna(c2):
            columnas.append(c1)
        else:
            columnas.append(f"{c1} - {c2}")

    # Encabezados base
    base_cols = ["DNI","Nombre Completo","CARGO","F. Ingreso","OFICINA","CENTRO COSTO","CENTRO COSTO CODIGO"]
    all_columns = base_cols + columnas

    # --------------------------------------------------
    # LEER DATOS
    # --------------------------------------------------

    df_data = df_raw.iloc[2:, :len(all_columns)]
    df_data.columns = all_columns
    df_data = df_data.reset_index(drop=True)

    # --------------------------------------------------
    # TRANSFORMAR A FORMATO LARGO
    # --------------------------------------------------

    cursos = list(set([col.split(" - ")[0] for col in columnas if " - " in col]))

    registros = []
    for idx, row in df_data.iterrows():
        for curso in cursos:
            f_dictado = row.get(f"{curso} - F. DICTADO")
            nota = row.get(f"{curso} - NOTA")
            venc = row.get(f"{curso} - VENCIMIENTO")
            dias = row.get(f"{curso} - VENC. DIAS")
            estado = row.get(f"{curso} - ESTADO")

            if pd.notna(f_dictado) or pd.notna(venc):
                registros.append({
                    "DNI": row["DNI"],
                    "Nombre Completo": row["Nombre Completo"],
                    "Cargo": row["CARGO"],
                    "F. Ingreso": row["F. Ingreso"],
                    "Oficina": row["OFICINA"],
                    "Centro Costo": row["CENTRO COSTO"],
                    "Centro Costo Codigo": row["CENTRO COSTO CODIGO"],
                    "Curso": curso,
                    "F. Dictado": f_dictado,
                    "Nota": nota,
                    "Vencimiento": venc,
                    "Venc. Dias": dias,
                    "Estado": estado
                })

    df_final = pd.DataFrame(registros)

    # --------------------------------------------------
    # CALCULAR ESTADOS AUTOMÁTICOS
    # --------------------------------------------------

    hoy = pd.Timestamp.today().normalize()
    df_final["F. Dictado"] = pd.to_datetime(df_final["F. Dictado"], errors="coerce", dayfirst=True)
    df_final["Vencimiento"] = pd.to_datetime(df_final["Vencimiento"], errors="coerce", dayfirst=True)
    df_final["Venc. Dias"] = (df_final["Vencimiento"] - hoy).dt.days

    df_final.loc[df_final["Vencimiento"].isna(),"Estado"] = "VIGENTE"
    df_final.loc[df_final["Venc. Dias"] < 0,"Estado"] = "VENCIDO"
    df_final.loc[(df_final["Venc. Dias"] >=0) & (df_final["Venc. Dias"] <=30),"Estado"] = "POR VENCER"
    df_final.loc[df_final["Venc. Dias"] > 30,"Estado"] = "VIGENTE"

    # --------------------------------------------------
    # KPIs
    # --------------------------------------------------

    st.subheader("📈 Resumen de Capacitaciones")
    vigentes = (df_final["Estado"]=="VIGENTE").sum()
    por_vencer = (df_final["Estado"]=="POR VENCER").sum()
    vencidos = (df_final["Estado"]=="VENCIDO").sum()

    c1,c2,c3 = st.columns(3)
    c1.markdown(f"<div class='metric-box verde'>🟢 Vigentes<br><h2>{vigentes}</h2></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='metric-box amarillo'>🟡 Por vencer<br><h2>{por_vencer}</h2></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='metric-box rojo'>🔴 Vencidos<br><h2>{vencidos}</h2></div>", unsafe_allow_html=True)

    # --------------------------------------------------
    # TABLA FORMATEADA
    # --------------------------------------------------

    def color_estado(val):
        val = str(val).upper()
        if val=="VIGENTE":
            return "background-color:#28a745;color:white"
        elif val=="POR VENCER":
            return "background-color:#ffc107;color:black"
        elif val=="VENCIDO":
            return "background-color:#dc3545;color:white"
        return ""

    st.subheader("📊 Vista Previa de Capacitaciones")
    st.dataframe(df_final.style.applymap(color_estado), height=500, use_container_width=True)

    # --------------------------------------------------
    # EXPORTAR EXCEL FORMATEADO
    # --------------------------------------------------

    output = BytesIO()
    df_final.to_excel(output, index=False)
    output.seek(0)
    wb2 = openpyxl.load_workbook(output)
    ws2 = wb2.active

    thin = Side(style='thin')
    border = Border(left=thin,right=thin,top=thin,bottom=thin)
    fill_map = {"VENCIDO":"FF4C4C","POR VENCER":"FFF2CC","VIGENTE":"A7D129"}

    for row in ws2.iter_rows(min_row=2):
        estado = row[12].value
        color = fill_map.get(str(estado).upper(), None)
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if color:
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.border = border

    output_final = BytesIO()
    wb2.save(output_final)
    output_final.seek(0)

    st.markdown("### 📥 Descargar reporte final")
    st.download_button(
        "⬇️ Descargar Excel Profesional TALMA",
        data=output_final,
        file_name="Capacitaciones_TALMA_Profesional.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Sube un archivo Excel para comenzar")
