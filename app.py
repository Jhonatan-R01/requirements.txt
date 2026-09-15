import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Conciliación Bancaria", layout="wide")

st.title("📊 Sistema de Conciliación Bancaria Inteligente")
st.write("Sube los dos archivos en formato Excel para realizar el cruce automático de información sin duplicados.")

# Carga de archivos
col1, col2 = st.columns(2)

with col1:
    file_banco = st.file_uploader("1. Cargar Extracto Bancario (.xlsx)", type=["xlsx"])

with col2:
    file_auxiliar = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx"])

if file_banco and file_auxiliar:
    st.info("Procesando archivos y ejecutando algoritmo de cruce...")
    
    # Leer archivos
    df_banco = pd.read_excel(file_banco)
    df_aux = pd.read_excel(file_auxiliar)
    
    # Estandarizar columnas principales
    # Se asume que en banco existen: Fecha, Referencia, Retiros ()
    # En auxiliar existen: Fecha, Documento, NOMBRE, Egresos (-)
    
    # Convertir fechas y montos
    df_banco['Fecha_Clean'] = pd.to_datetime(df_banco['Fecha'], dayfirst=True, errors='coerce').dt.strftime('%Y-%m-%d')
    df_aux['Fecha_Clean'] = pd.to_datetime(df_aux['Fecha'], dayfirst=True, errors='coerce').dt.strftime('%Y-%m-%d')
    
    df_banco['Monto_Clean'] = pd.to_numeric(df_banco['Retiros ()'], errors='coerce').round(2)
    df_aux['Monto_Clean'] = pd.to_numeric(df_aux['Egresos (-)'], errors='coerce').round(2)

    # Crear contador por secuencia (Ocurrencia) para romper duplicados en misma fecha y monto
    df_banco['Ocurrencia'] = df_banco.groupby(['Fecha_Clean', 'Monto_Clean']).cumcount() + 1
    df_aux['Ocurrencia'] = df_aux.groupby(['Fecha_Clean', 'Monto_Clean']).cumcount() + 1

    # Crear Llave Única
    df_banco['LLAVE'] = df_banco['Fecha_Clean'] + "_" + df_banco['Monto_Clean'].astype(str) + "_" + df_banco['Ocurrencia'].astype(str)
    df_aux['LLAVE'] = df_aux['Fecha_Clean'] + "_" + df_aux['Monto_Clean'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)

    # Realizar Cruce (Merge)
    conciliados = pd.merge(
        df_banco, 
        df_aux, 
        on='LLAVE', 
        how='inner', 
        suffixes=('_Banco', '_Auxiliar')
    )
    
    # Partidas solo en Banco
    solo_banco = df_banco[~df_banco['LLAVE'].isin(conciliados['LLAVE'])]
    
    # Partidas solo en Auxiliar
    solo_aux = df_aux[~df_aux['LLAVE'].isin(conciliados['LLAVE'])]

    st.success("✅ ¡Conciliación realizada con éxito!")

    # Métricas principales
    m1, m2, m3 = st.columns(3)
    m1.metric("Partidas Conciliadas", len(conciliados))
    m2.metric("Pendientes en Banco", len(solo_banco))
    m3.metric("Pendientes en Contabilidad", len(solo_aux))

    # Pestañas de resultados
    tab1, tab2, tab3 = st.tabs(["✅ Conciliados", "⚠️ Solo en Banco", "⚠️ Solo en Contabilidad"])

    with tab1:
        st.dataframe(conciliados[['Fecha_Banco', 'Referencia', 'Retiros ()', 'Documento', 'NOMBRE']])

    with tab2:
        st.dataframe(solo_banco[['Fecha', 'Referencia', 'Retiros ()']])

    with tab3:
        st.dataframe(solo_aux[['Fecha', 'Documento', 'NOMBRE', 'Egresos (-)']])

    # Generar Excel para descarga
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        conciliados.to_excel(writer, sheet_name='Conciliados', index=False)
        solo_banco.to_excel(writer, sheet_name='Pendientes_Banco', index=False)
        solo_aux.to_excel(writer, sheet_name='Pendientes_Contabilidad', index=False)
    
    st.download_button(
        label="📥 Descargar Resultado en Excel",
        data=output.getvalue(),
        file_name="Resultado_Conciliacion_Bancaria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
