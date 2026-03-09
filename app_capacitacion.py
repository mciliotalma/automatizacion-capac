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
from openpyxl.utils import get_column_letter

# --------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# --------------------------------------------------

st.set_page_config(
    page_title="Reporte Capacitaciones TALMA",
    page_icon="📊",
    layout="wide"
)

# --------------------------------------------------
# CSS CORPORATIVO TALMA
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
# HEADER CON LOGO TALMA
# --------------------------------------------------

col1, col2 = st.columns([1,6])

with col1:
    st.image("logo.jpg", width=120)

with col2:
    st.markdown("""
    <div style="
    background: linear-gradient(90deg,#004C97,#005EB8);
    padding:20px;
    border-radius:12px;
    ">
    
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
# CARGA DE ARCHIVO
# --------------------------------------------------

st.markdown("### 📂 Cargar archivo de capacitaciones")

col1,col2,col3 = st.columns([1,2,1])

with col2:

    uploaded_file = st.file_uploader(
        "Arrastra o selecciona tu archivo Excel",
        type=["xlsx","xlsm"]
    )

# --------------------------------------------------
# PROCESAR ARCHIVO
# --------------------------------------------------

if uploaded_file is not None:

    st.success("Archivo cargado correctamente")

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)

    ws = wb['Acumulado Portal']

    headers = [
        "DNI","Nombre Completo","Cargo","F. Ingreso","Oficina",
        "Centro Costo","Centro Costo Codigo","Curso","F. Dictado",
        "Nota","Vencimiento","Venc. Dias","Estado"
    ]

    data = []

    ult_fila = ws.max_row
    ult_col = ws.max_column

    cursos = [ws.cell(row=1, column=j).value for j in range(8, ult_col+1, 5)]

    # --------------------------------------------------
    # TRANSFORMAR DATOS
    # --------------------------------------------------

    for i in range(2, ult_fila + 1):

        fila = [cell.value for cell in ws[i]]

        if len(fila) < 7:
            continue

        dni = fila[0]

        if str(dni).strip().upper() == "DNI":
            continue

        nombre = fila[1]
        cargo = fila[2]
        f_ingreso = fila[3]
        oficina = fila[4]
        centro_costo = fila[5]
        centro_costo_codigo = fila[6]

        for idx, j in enumerate(range(7, ult_col, 5)):

            if j + 4 >= len(fila):
                break

            curso = cursos[idx]
            f_dictado = fila[j]
            nota = fila[j+1]
            vencimiento = fila[j+2]
            venc_dias = fila[j+3]
            estado = fila[j+4]

            if any([curso, f_dictado, nota, vencimiento, venc_dias, estado]):

                data.append([
                    dni, nombre, cargo, f_ingreso, oficina,
                    centro_costo, centro_costo_codigo,
                    curso, f_dictado, nota, vencimiento,
                    venc_dias, estado
                ])

    df = pd.DataFrame(data, columns=headers)

    # --------------------------------------------------
    # CALCULAR ESTADO
    # --------------------------------------------------

    hoy = pd.Timestamp.today().normalize()

    df['F. Dictado'] = pd.to_datetime(df['F. Dictado'], errors='coerce', dayfirst=True)

    df['Vencimiento'] = pd.to_datetime(df['Vencimiento'], errors='coerce', dayfirst=True)

    df['Venc. Dias'] = (df['Vencimiento'] - hoy).dt.days

    df.loc[df['Vencimiento'].isna(),'Estado'] = 'VIGENTE'
    df.loc[df['Venc. Dias'] < 0,'Estado'] = 'VENCIDO'
    df.loc[(df['Venc. Dias'] >=0) & (df['Venc. Dias'] <=30),'Estado'] = 'POR VENCER'
    df.loc[df['Venc. Dias'] > 30,'Estado'] = 'VIGENTE'

    # --------------------------------------------------
    # KPIs
    # --------------------------------------------------

    st.markdown("### 📈 Resumen de Capacitaciones")

    vigentes = (df["Estado"]=="VIGENTE").sum()
    por_vencer = (df["Estado"]=="POR VENCER").sum()
    vencidos = (df["Estado"]=="VENCIDO").sum()

    c1,c2,c3 = st.columns(3)

    c1.metric("🟢 Vigentes", vigentes)
    c2.metric("🟡 Por vencer", por_vencer)
    c3.metric("🔴 Vencidos", vencidos)

    # --------------------------------------------------
    # TABLA
    # --------------------------------------------------

    st.markdown("### 📊 Vista previa de la tabla")

    st.dataframe(
        df,
        use_container_width=True,
        height=500
    )

    # --------------------------------------------------
    # CREAR EXCEL FORMATEADO
    # --------------------------------------------------

    output = BytesIO()

    df.to_excel(output, index=False, engine='openpyxl')

    output.seek(0)

    wb2 = openpyxl.load_workbook(output)

    ws2 = wb2.active

    thin = Side(style='thin')

    border = Border(left=thin,right=thin,top=thin,bottom=thin)

    fill_map = {
        "VENCIDO": "FF4C4C",
        "POR VENCER": "FFEB9C",
        "VIGENTE": "A7D129"
    }

    for row in ws2.iter_rows(min_row=2):

        estado_cell = row[12]

        estado_value = estado_cell.value

        fill_color = fill_map.get(str(estado_value).upper(), None)

        for cell in row:

            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')

            if fill_color:

                cell.fill = PatternFill(
                    start_color=fill_color,
                    end_color=fill_color,
                    fill_type="solid"
                )

    for cell in ws2[1]:

        cell.font = Font(bold=True)
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical='center')

    for col in ws2.columns:

        max_length = 0
        column = col[0].column

        for cell in col:

            if cell.value:

                max_length = max(max_length, len(str(cell.value)))

        ws2.column_dimensions[get_column_letter(column)].width = min(max_length+2,40)

    for row in ws2.iter_rows():

        ws2.row_dimensions[row[0].row].height = 20

    output_final = BytesIO()

    wb2.save(output_final)

    output_final.seek(0)

    # --------------------------------------------------
    # DESCARGA
    # --------------------------------------------------

    st.markdown("### 📥 Descargar Reporte")

    st.download_button(
        "⬇️ Descargar Excel Profesional TALMA",
        data=output_final,
        file_name="Capacitaciones_Profesional_TALMA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
