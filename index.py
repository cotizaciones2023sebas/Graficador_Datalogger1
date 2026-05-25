
import streamlit as st
import pandas as pd
import numpy as np
import math
from io import BytesIO
from datetime import datetime
import plotly.express as px
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

st.set_page_config(page_title="Analizador Eléctrico Industrial", layout="wide")

st.title("⚡ Analizador Eléctrico Industrial")
st.markdown("Dashboard profesional para análisis de calidad de energía")

# ==========================
# Sidebar
# ==========================
st.sidebar.header("📋 Información General")

cliente = st.sidebar.text_input("Cliente")
proyecto = st.sidebar.text_input("Proyecto")
ciudad = st.sidebar.text_input("Ciudad")
asesor = st.sidebar.text_input("Asesor Comercial")
fecha_analisis = st.sidebar.date_input("Fecha de análisis")
observaciones = st.sidebar.text_area("Observaciones")

uploaded_files = st.file_uploader(
    "📂 Cargar archivos CSV",
    type=["csv"],
    accept_multiple_files=True
)

def detectar_delimitador(file):
    sample = file.read(2048).decode("utf-8", errors="ignore")
    file.seek(0)
    return ";" if sample.count(";") > sample.count(",") else ","

def encontrar_columna(cols, palabras):
    for c in cols:
        low = c.lower()
        for p in palabras:
            if p in low:
                return c
    return None

if uploaded_files:

    dataframes = []

    for file in uploaded_files:
        sep = detectar_delimitador(file)

        try:
            df = pd.read_csv(file, sep=sep)
        except:
            df = pd.read_csv(file, sep=sep, encoding="latin1")

        df.columns = [c.strip() for c in df.columns]
        dataframes.append(df)

    df_total = pd.concat(dataframes, ignore_index=True)

    st.success(f"✅ Archivos procesados: {len(uploaded_files)}")

    cols = df_total.columns.tolist()

    # ==========================
    # Detectar columnas
    # ==========================

    v1 = encontrar_columna(cols, ["tension l1", "voltaje l1", "voltage l1"])
    v2 = encontrar_columna(cols, ["tension l2", "voltaje l2", "voltage l2"])
    v3 = encontrar_columna(cols, ["tension l3", "voltaje l3", "voltage l3"])

    i1 = encontrar_columna(cols, ["corriente l1", "current l1"])
    i2 = encontrar_columna(cols, ["corriente l2", "current l2"])
    i3 = encontrar_columna(cols, ["corriente l3", "current l3"])

    fp_col = encontrar_columna(cols, ["factor de potencia", "fp"])

    # convertir numéricos
    for c in [v1,v2,v3,i1,i2,i3,fp_col]:
        if c:
            df_total[c] = pd.to_numeric(df_total[c], errors="coerce")

    # ==========================
    # Cálculos
    # ==========================

    corrientes = {}

    if i1: corrientes["L1"] = df_total[i1].max()
    if i2: corrientes["L2"] = df_total[i2].max()
    if i3: corrientes["L3"] = df_total[i3].max()

    fase_max = max(corrientes, key=corrientes.get)
    corriente_max = corrientes[fase_max]

    voltajes = []
    for v in [v1,v2,v3]:
        if v:
            voltajes.append(df_total[v].mean())

    voltaje_promedio = np.mean(voltajes)

    fp_prom = 0.95
    if fp_col:
        fp_prom = df_total[fp_col].mean()

    kva = math.ceil((math.sqrt(3) * voltaje_promedio * corriente_max) / 1000)
    kw = math.ceil(kva * fp_prom)

    # ==========================
    # KPIs
    # ==========================

    c1,c2,c3,c4 = st.columns(4)

    c1.metric("⚡ Corriente Máxima", f"{corriente_max:.2f} A")
    c2.metric("🔌 Voltaje Promedio", f"{voltaje_promedio:.2f} V")
    c3.metric("🏭 kVA Calculado", f"{kva} kVA")
    c4.metric("⚙️ kW Calculado", f"{kw} kW")

    st.markdown("---")

    # ==========================
    # Gráficas
    # ==========================

    graficas = []

    for col in [v1,v2,v3]:
        if col:
            fig = px.line(df_total, y=col, title=f"{col}")
            st.plotly_chart(fig, use_container_width=True)
            graficas.append(fig)

    for col in [i1,i2,i3]:
        if col:
            fig = px.line(df_total, y=col, title=f"{col}")
            st.plotly_chart(fig, use_container_width=True)
            graficas.append(fig)

    # ==========================
    # Resumen técnico
    # ==========================

    st.subheader("📑 Resumen Técnico")

    resumen = f"""
    El sistema eléctrico analizado presenta una corriente máxima de {corriente_max:.2f} A
    en la fase {fase_max}. El voltaje promedio registrado fue de {voltaje_promedio:.2f} V.

    La potencia aparente calculada es de {kva} kVA y la potencia activa estimada es de {kw} kW.

    El factor de potencia promedio del sistema fue de {fp_prom:.2f}.

    Se recomienda verificar el balance de cargas entre fases y validar el estado de la instalación
    para evitar posibles sobrecargas.
    """

    st.write(resumen)

    # ==========================
    # Exportar Excel
    # ==========================

    if st.button("📥 Generar Excel"):

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            info_df = pd.DataFrame({
                "Campo":[
                    "Cliente","Proyecto","Ciudad","Asesor","Fecha"
                ],
                "Valor":[
                    cliente, proyecto, ciudad, asesor, str(fecha_analisis)
                ]
            })

            resumen_df = pd.DataFrame({
                "Indicador":[
                    "Corriente Máxima",
                    "Voltaje Promedio",
                    "kVA",
                    "kW",
                    "FP Promedio"
                ],
                "Valor":[
                    corriente_max,
                    voltaje_promedio,
                    kva,
                    kw,
                    fp_prom
                ]
            })

            info_df.to_excel(writer, sheet_name="Información", index=False)
            resumen_df.to_excel(writer, sheet_name="Resumen", index=False)
            df_total.to_excel(writer, sheet_name="Datos", index=False)

        output.seek(0)

        st.download_button(
            label="⬇️ Descargar Excel",
            data=output,
            file_name="analisis_electrico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("📂 Cargue uno o varios archivos CSV para iniciar el análisis.")
