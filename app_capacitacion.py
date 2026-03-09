# -*- coding: utf-8 -*-
"""
Created on Mon Mar  9 09:07:40 2026

@author: mcilio
"""

# app_capacitacion_prof.py
import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
from io import BytesIO
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Reporte de Capacitaciones", layout="wide")
st.title("📊 Reporte Profesional de Capacitaciones")

st.sidebar.header("Opciones")
uploaded_file = st.sidebar.file_uploader("Sube tu archivo Excel", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    st.sidebar.info("Procesando archivo...")

    # -----------------------------
    # Cargar Excel
    # -----------------------------
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

    # -----------------------------
    # Transformación de datos
    # -----------------------------
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
                    curso, f_dictado, nota, vencimiento, venc_dias, estado
                ])

    df = pd.DataFrame(data, columns=headers)

    # -----------------------------
    # Calculo automático de Venc. Dias y Estado
    # -----------------------------
    hoy = pd.Timestamp.today().normalize()
    df['F. Dictado'] = pd.to_datetime(df['F. Dictado'], errors='coerce', dayfirst=True)
    df['Vencimiento'] = pd.to_datetime(df['Vencimiento'], errors='coerce', dayfirst=True)
    df['Venc. Dias'] = (df['Vencimiento'] - hoy).dt.days
    df.loc[df['Vencimiento'].isna(),'Estado'] = 'VIGENTE'
    df.loc[df['Venc. Dias'] < 0,'Estado'] = 'VENCIDO'
    df.loc[(df['Venc. Dias'] >=0) & (df['Venc. Dias'] <=30),'Estado'] = 'POR VENCER'
    df.loc[df['Venc. Dias'] > 30,'Estado'] = 'VIGENTE'

    # -----------------------------
    # Vista previa en Streamlit
    # -----------------------------
    st.subheader("Vista previa de la tabla")
    st.dataframe(df.style.apply(
        lambda x: ['background-color: #f8d7da' if v=="VENCIDO" else
                   'background-color: #fff3cd' if v=="POR VENCER" else
                   'background-color: #d4edda' if v=="VIGENTE" else ''
                   for v in x['Estado']], axis=1
    ))

    # -----------------------------
    # Generar Excel formateado
    # -----------------------------
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    wb2 = openpyxl.load_workbook(output)
    ws2 = wb2.active

    thin = Side(style='thin')
    border = Border(left=thin,right=thin,top=thin,bottom=thin)

    fill_map = {
        "VENCIDO": "FFC7CE",      # rojo claro
        "POR VENCER": "FFEB9C",   # amarillo claro
        "VIGENTE": "C6EFCE"       # verde claro
    }

    for row in ws2.iter_rows(min_row=2):
        estado_cell = row[12]  # columna Estado
        estado_value = estado_cell.value
        fill_color = fill_map.get(str(estado_value).upper(), None)
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if fill_color:
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical='center')

    # Ajustar ancho columnas
    for col in ws2.columns:
        max_length = 0
        column = col[0].column
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length,len(str(cell.value)))
            except:
                pass
        ws2.column_dimensions[get_column_letter(column)].width = min(max_length+2,40)

    # Altura filas
    for row in ws2.iter_rows():
        ws2.row_dimensions[row[0].row].height = 20

    output_final = BytesIO()
    wb2.save(output_final)
    output_final.seek(0)

    # -----------------------------
    # Botón de descarga
    # -----------------------------
    st.subheader("Descargar archivo final")
    st.download_button(
        label="⬇️ Descargar Excel profesional",
        data=output_final,
        file_name="Capacitaciones_Profesional.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
