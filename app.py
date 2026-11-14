
import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="PAGY - Optimizador de rutas",
    layout="wide"
)

st.title("Optimizador de Transporte Porcino")
st.write("Planificación de movimientos y bioseguridad")

# Subir archivo Excel
archivo = st.file_uploader(
    "Sube el Excel con los datos de las granjas",
    type=["xlsx"]
)

if archivo is not None:

    excel = pd.ExcelFile(archivo)

    st.success("Excel cargado correctamente")

    # Mostrar las pestañas disponibles
    st.subheader("Pestañas del Excel")

    hoja = st.selectbox(
        "Selecciona una pestaña",
        excel.sheet_names
    )

    datos = pd.read_excel(excel, sheet_name=hoja)

    # Mostrar información
    col1, col2 = st.columns(2)

    col1.metric("Registros", len(datos))
    col2.metric("Columnas", len(datos.columns))

    # Mostrar tabla
    st.subheader("Datos de la pestaña seleccionada")
    st.dataframe(datos, use_container_width=True)

else:
    st.info("Sube el archivo Excel para comenzar.")
