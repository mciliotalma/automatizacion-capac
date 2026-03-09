# -*- coding: utf-8 -*-
"""
Created on Mon Mar  9 09:07:40 2026

@author: mcilio
"""
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
# CSS CORPORATIVO
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
# PLANTILLA REAL
# --------------------------------------------------

st.markdown("## 📄 Descargar plantilla oficial")

st.info("Utiliza esta plantilla para cargar correctamente el archivo.")

wb_template = openpyxl.Workbook()
ws_t = wb_template.active
ws_t.title = "Acumulado Portal"

base_headers = [
    "DNI","NOMBRE COMPLETO","CARGO","F. DE INGRESO",
    "OFICINA","CENTRO COSTO","CENTRO COSTO CODIGO"
]

cursos = ["Curso1","Curso2"]

sub_headers = ["F. DICTADO","NOTA","VENCIMIENTO","VENC. DIAS","ESTADO"]

col = 1

for h in base_headers:
    ws_t.cell(row=2,column=col,value=h)
    ws_t.merge_cells(start_row=1,start_column=col,end_row=2,end_column=col)
    col += 1

for curso in cursos:

    start_col = col

    for sub in sub_headers:
        ws_t.cell(row=2,column=col,value=sub)
        col += 1

    ws_t.merge_cells(
        start_row=1,
        start_column=start_col,
        end_row=1,
        end_column=start_col+4
    )

    ws_t.cell(row=1,column=start_col,value=curso)

thin = Side(style='thin')
border = Border(left=thin,right=thin,top=thin,bottom=thin)

for row in ws_t.iter_rows(min_row=1,max_row=2):
    for cell in row:
        cell.border = border
        cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
        cell.font = Font(bold=True)

ws_t.append([
    "0954082780",
    "ABAD HUACON SUSANNE PAMELA",
    "AGENTE DE SERVICIO AL PASAJERO",
    "16/10/2023",
    "GUAYAQUIL",
    "PAX GYE",
    "15030102",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
])

output_template = BytesIO()
wb_template.save(output_template)
output_template.seek(0)

st.download_button(
    "⬇️ Descargar plantilla",
    data=output_template,
    file_name="Plantilla_Capacitaciones_TALMA.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")

# --------------------------------------------------
# CARGAR ARCHIVO
# --------------------------------------------------

st.markdown("## 📂 Cargar archivo de capacitaciones")

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
    ws = wb["Acumulado Portal"]

    ult_fila = ws.max_row
    ult_col = ws.max_column

    cursos = []

    for j in range(8, ult_col+1, 5):
        cursos.append(ws.cell(row=1,column=j).value)

    headers = [
        "DNI","Nombre Completo","Cargo","F. Ingreso",
        "Oficina","Centro Costo","Centro Costo Codigo",
        "Curso","F. Dictado","Nota","Vencimiento",
        "Venc. Dias","Estado"
    ]

    data = []

    for i in range(3, ult_fila+1):

        fila = [cell.value for cell in ws[i]]

        dni = fila[0]

        if dni is None:
            continue

        nombre = fila[1]
        cargo = fila[2]
        f_ingreso = fila[3]
        oficina = fila[4]
        centro_costo = fila[5]
        centro_costo_codigo = fila[6]

        for idx, j in enumerate(range(7, ult_col, 5)):

            if j+4 >= len(fila):
                break

            curso = cursos[idx]

            f_dictado = fila[j]
            nota = fila[j+1]
            vencimiento = fila[j+2]
            venc_dias = fila[j+3]
            estado = fila[j+4]

            if any([f_dictado, vencimiento]):

                data.append([
                    dni,nombre,cargo,f_ingreso,oficina,
                    centro_costo,centro_costo_codigo,
                    curso,f_dictado,nota,vencimiento,
                    venc_dias,estado
                ])

    df = pd.DataFrame(data, columns=headers)

    # --------------------------------------------------
    # CALCULAR ESTADOS
    # --------------------------------------------------

    hoy = pd.Timestamp.today().normalize()

    df["F. Dictado"] = pd.to_datetime(df["F. Dictado"],errors="coerce",dayfirst=True)
    df["Vencimiento"] = pd.to_datetime(df["Vencimiento"],errors="coerce",dayfirst=True)

    df["Venc. Dias"] = (df["Vencimiento"] - hoy).dt.days

    df.loc[df["Vencimiento"].isna(),"Estado"] = "VIGENTE"
    df.loc[df["Venc. Dias"] < 0,"Estado"] = "VENCIDO"
    df.loc[(df["Venc. Dias"]>=0) & (df["Venc. Dias"]<=30),"Estado"] = "POR VENCER"
    df.loc[df["Venc. Dias"]>30,"Estado"] = "VIGENTE"

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

    st.markdown("### 📊 Vista previa")

    st.dataframe(
        df,
        use_container_width=True,
        height=500
    )

    # --------------------------------------------------
    # EXPORTAR EXCEL PROFESIONAL
    # --------------------------------------------------

    output = BytesIO()
    df.to_excel(output,index=False)
    output.seek(0)

    wb2 = openpyxl.load_workbook(output)
    ws2 = wb2.active

    thin = Side(style="thin")
    border = Border(left=thin,right=thin,top=thin,bottom=thin)

    fill_map = {
        "VENCIDO":"FF4C4C",
        "POR VENCER":"FFF2CC",
        "VIGENTE":"A7D129"
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
