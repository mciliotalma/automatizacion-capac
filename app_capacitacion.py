# -*- coding: utf-8 -*-
"""
Created on Mon Mar  9 09:07:40 2026

@author: mcilio
"""
# -*- coding: utf-8 -*-
"""
Reporte Profesional de Capacitaciones TALMA
"""

import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
from io import BytesIO
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill

# --------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------

st.set_page_config(
    page_title="Reporte Capacitaciones TALMA",
    page_icon="📊",
    layout="wide"
)

# --------------------------------------------------
# ESTILO TALMA
# --------------------------------------------------

st.markdown("""
<style>

.main{
background-color:#F4F7FB;
}

h1,h2,h3{
color:#004C97;
}

.stButton>button{
background-color:#A7D129;
color:#004C97;
font-weight:bold;
border-radius:8px;
}

.stDownloadButton>button{
background-color:#A7D129;
color:#004C97;
font-weight:bold;
border-radius:8px;
}

</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# HEADER
# --------------------------------------------------

col1, col2 = st.columns([1,6])

with col1:
    st.image("logo.jpg", width=120)

with col2:
    st.markdown("""
    <div style="
    background: linear-gradient(90deg,#004C97,#005EB8);
    padding:20px;
    border-radius:12px;">
    
    <h2 style="color:white;margin:0;">
    📊 Reporte Profesional de Capacitaciones TALMA
    </h2>
    
    <p style="color:white;margin:0;">
    Sistema automatizado de control de vigencia de capacitaciones
    </p>
    
    </div>
    """, unsafe_allow_html=True)

st.write("")

# --------------------------------------------------
# CARGAR ARCHIVO
# --------------------------------------------------

st.markdown("## 📂 Cargar archivo de capacitaciones")

uploaded_file = st.file_uploader(
    "Arrastra o selecciona tu archivo Excel",
    type=["xlsx","xlsm"]
)

# --------------------------------------------------
# PROCESAR
# --------------------------------------------------

if uploaded_file is not None:

    st.success("Archivo cargado correctamente")

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb['Acumulado Portal']

    ult_col = ws.max_column
    ult_row = ws.max_row

    # --------------------------------------------------
    # LEER DOBLE ENCABEZADO
    # --------------------------------------------------

    header_curso = [ws.cell(row=1, column=i).value for i in range(1, ult_col+1)]
    header_sub = [ws.cell(row=2, column=i).value for i in range(1, ult_col+1)]

    columnas = []

    for c1, c2 in zip(header_curso, header_sub):

        if c2 is None:
            columnas.append(c1)

        elif c1 is None:
            columnas.append(c2)

        else:
            columnas.append(f"{c1} - {c2}")

    # --------------------------------------------------
    # LEER DATOS
    # --------------------------------------------------

    data = []

    for row in ws.iter_rows(min_row=3, values_only=True):
        data.append(row)

    df = pd.DataFrame(data, columns=columnas)

    # --------------------------------------------------
    # LIMPIAR COLUMNAS BASE
    # --------------------------------------------------

    df = df.rename(columns={
        "NOMBRE COMPLETO":"Nombre Completo",
        "F. DE INGRESO":"F. Ingreso",
        "CENTRO COSTO CODIGO":"Centro Costo Codigo"
    })

    # --------------------------------------------------
    # IDENTIFICAR CURSOS
    # --------------------------------------------------

    cursos = list(set([
        col.split(" - ")[0]
        for col in df.columns
        if " - " in col
    ]))

    # --------------------------------------------------
    # TRANSFORMAR A FORMATO LARGO
    # --------------------------------------------------

    registros = []

    for _, row in df.iterrows():

        for curso in cursos:

            f_dictado = row.get(f"{curso} - F. DICTADO")
            nota = row.get(f"{curso} - NOTA")
            venc = row.get(f"{curso} - VENCIMIENTO")
            dias = row.get(f"{curso} - VENC. DIAS")
            estado = row.get(f"{curso} - ESTADO")

            if pd.notna(f_dictado) or pd.notna(venc):

                registros.append({

                    "DNI":row["DNI"],
                    "Nombre Completo":row["Nombre Completo"],
                    "Cargo":row["CARGO"],
                    "F. Ingreso":row["F. Ingreso"],
                    "Oficina":row["OFICINA"],
                    "Centro Costo":row["CENTRO COSTO"],
                    "Centro Costo Codigo":row["Centro Costo Codigo"],

                    "Curso":curso,
                    "F. Dictado":f_dictado,
                    "Nota":nota,
                    "Vencimiento":venc,
                    "Venc. Dias":dias,
                    "Estado":estado
                })

    df_final = pd.DataFrame(registros)

    # --------------------------------------------------
    # CALCULAR VENCIMIENTOS
    # --------------------------------------------------

    hoy = pd.Timestamp.today().normalize()

    df_final["F. Dictado"] = pd.to_datetime(df_final["F. Dictado"], errors="coerce")
    df_final["Vencimiento"] = pd.to_datetime(df_final["Vencimiento"], errors="coerce")

    df_final["Venc. Dias"] = (df_final["Vencimiento"] - hoy).dt.days

    df_final.loc[df_final["Vencimiento"].isna(),"Estado"] = "VIGENTE"
    df_final.loc[df_final["Venc. Dias"] < 0,"Estado"] = "VENCIDO"
    df_final.loc[(df_final["Venc. Dias"] >=0) & (df_final["Venc. Dias"]<=30),"Estado"] = "POR VENCER"
    df_final.loc[df_final["Venc. Dias"] > 30,"Estado"] = "VIGENTE"

    # --------------------------------------------------
    # KPIs
    # --------------------------------------------------

    st.markdown("### 📈 Resumen de Capacitaciones")

    vigentes = (df_final["Estado"]=="VIGENTE").sum()
    por_vencer = (df_final["Estado"]=="POR VENCER").sum()
    vencidos = (df_final["Estado"]=="VENCIDO").sum()

    c1,c2,c3 = st.columns(3)

    c1.metric("🟢 Vigentes", vigentes)
    c2.metric("🟡 Por vencer", por_vencer)
    c3.metric("🔴 Vencidos", vencidos)

    # --------------------------------------------------
    # TABLA
    # --------------------------------------------------

    st.markdown("### 📊 Vista previa")

    st.dataframe(
        df_final,
        use_container_width=True,
        height=500
    )

    # --------------------------------------------------
    # EXPORTAR EXCEL
    # --------------------------------------------------

    output = BytesIO()
    df_final.to_excel(output, index=False)
    output.seek(0)

    wb2 = openpyxl.load_workbook(output)
    ws2 = wb2.active

    thin = Side(style='thin')

    border = Border(
        left=thin,right=thin,top=thin,bottom=thin
    )

    fill_map = {
        "VENCIDO": "FF4C4C",
        "POR VENCER": "FFF2CC",
        "VIGENTE": "A7D129"
    }

    for row in ws2.iter_rows(min_row=2):

        estado = row[12].value

        color = fill_map.get(str(estado).upper(),None)

        for cell in row:

            cell.border = border
            cell.alignment = Alignment(wrap_text=True)

            if color:

                cell.fill = PatternFill(
                    start_color=color,
                    end_color=color,
                    fill_type="solid"
                )

    for cell in ws2[1]:

        cell.font = Font(bold=True)
        cell.border = border

    output_final = BytesIO()

    wb2.save(output_final)
    output_final.seek(0)

    st.markdown("### 📥 Descargar reporte")

    st.download_button(
        "⬇️ Descargar Excel Profesional TALMA",
        data=output_final,
        file_name="Capacitaciones_TALMA_Profesional.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
