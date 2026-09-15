import streamlit as st
import pandas as pd
import io

# Configuración de la página web
st.set_page_config(page_title="Sistema de Conciliación & Control Diario", layout="wide")

# ==========================================
# MENÚ LATERAL DE NAVEGACIÓN
# ==========================================
st.sidebar.title("📌 Menú Principal")
opcion = st.sidebar.radio(
    "Selecciona el módulo:",
    ["📖 1. Cruce Diario (CSV vs Auxiliar)", "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)"]
)

# ==========================================
# FUNCIONES DE PROCESAMIENTO Y LIMPIEZA
# ==========================================
def limpiar_valor(val):
    """Convierte cadenas de texto monetarias o numéricas a float firmado (+ / -)"""
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace('$', '').replace(' ', '')
    # Manejo de separadores miles/decimales formato latino
    if ',' in val_str and '.' in val_str:
        val_str = val_str.replace(',', '')
    elif ',' in val_str and '.' not in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

def procesar_auxiliar(df_raw):
    """Procesa la estructura fija del Auxiliar Contable"""
    # Los datos inician desde la fila 8 (índice 6 tiene los encabezados)
    headers = df_raw.iloc[6].values
    df = df_raw.iloc[8:].copy()
    df.columns = headers
    df = df.dropna(subset=['Código contable', 'Fecha elaboración'])
    
    # Calcular monto neto contable: Débito (+) / Crédito (-)
    df['Débito_Clean'] = df['Débito'].apply(limpiar_valor)
    df['Crédito_Clean'] = df['Crédito'].apply(limpiar_valor)
    df['Monto_Neto'] = df['Débito_Clean'] - df['Crédito_Clean']
    df['Monto_Abs'] = df['Monto_Neto'].abs().round(2)
    
    # Formatear fecha DD/MM/YYYY
    df['Fecha_Clean'] = pd.to_datetime(df['Fecha elaboración'], format='%d/%m/%Y', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_diario_csv(df_raw):
    """Procesa el CSV de movimiento diario sin encabezados fijados"""
    # Asignar nombres según estructura de columnas detectada
    df = df_raw.iloc[:, [0, 1, 3, 5, 6, 7]].copy()
    df.columns = ['Cuenta', 'Oficina', 'Fecha_Int', 'Valor_Raw', 'Codigo_Tx', 'Descripcion']
    
    df['Monto_Neto'] = df['Valor_Raw'].apply(limpiar_valor)
    df['Monto_Abs'] = df['Monto_Neto'].abs().round(2)
    df['Fecha_Clean'] = pd.to_datetime(df['Fecha_Int'].astype(str), format='%Y%m%d', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_extracto_excel(df_raw):
    """Procesa el extracto bancario en Excel (filas de información + movimientos)"""
    # Encabezados en fila 15 (índice 14)
    df = df_raw.iloc[15:].copy()
    df.columns = ['FECHA', 'DESCRIPCIÓN', 'SUCURSAL', 'DCTO', 'VALOR', 'SALDO', 'X1', 'X2']
    df = df.dropna(subset=['FECHA', 'VALOR'])
    df = df[~df['FECHA'].astype(str).str.contains('FECHA|FIN ESTADO', case=False, na=False)]
    
    df['Monto_Neto'] = df['VALOR'].apply(limpiar_valor)
    df['Monto_Abs'] = df['Monto_Neto'].abs().round(2)
    return df

# ==========================================
# MÓDULO 1: CRUCE DIARIO
# ==========================================
if opcion == "📖 1. Cruce Diario (CSV vs Auxiliar)":
    st.title("📖 Cruce Operativo Diario")
    st.write("Carga el movimiento diario registrado en **CSV** y el **Auxiliar Contable** para comparar transacciones del día.")

    col1, col2 = st.columns(2)
    with col1:
        file_diario = st.file_uploader("1. Cargar Movimiento Diario (.csv)", type=["csv"], key="diario_csv")
    with col2:
        file_auxiliar = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx"], key="aux_diario")

    if file_diario and file_auxiliar:
        df_diario_raw = pd.read_csv(file_diario, encoding="latin1", header=None)
        df_aux_raw = pd.read_excel(file_auxiliar)

        df_diario = procesar_diario_csv(df_diario_raw)
        df_aux = procesar_auxiliar(df_aux_raw)

        # Generar Llave de Cruce: Fecha + Monto Absoluto + Ocurrencia
        df_diario['Ocurrencia'] = df_diario.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1
        df_aux['Ocurrencia'] = df_aux.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1

        df_diario['LLAVE'] = df_diario['Fecha_Clean'] + "_" + df_diario['Monto_Abs'].astype(str) + "_" + df_diario['Ocurrencia'].astype(str)
        df_aux['LLAVE'] = df_aux['Fecha_Clean'] + "_" + df_aux['Monto_Abs'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)

        # Matcheo
        cruzados = pd.merge(df_diario, df_aux, on='LLAVE', how='inner', suffixes=('_Diario', '_Auxiliar'))
        solo_diario = df_diario[~df_diario['LLAVE'].isin(cruzados['LLAVE'])]
        solo_aux = df_aux[~df_aux['LLAVE'].isin(cruzados['LLAVE'])]

        st.success("✅ Cruce Diario ejecutado correctamente")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Partidas Cruzadas", len(cruzados))
        m2.metric("Pendientes en Registro Diario", len(solo_diario))
        m3.metric("Pendientes en Contabilidad", len(solo_aux))

        t1, t2, t3 = st.tabs(["✅ Cruzados", "⚠️ Solo en Registro Diario", "⚠️ Solo en Contabilidad"])
        with t1:
            st.dataframe(cruzados[['Fecha_Clean_Diario', 'Descripcion', 'Monto_Neto_Diario', 'Nombre del tercero', 'Comprobante']])
        with t2:
            st.dataframe(solo_diario[['Fecha_Clean', 'Descripcion', 'Codigo_Tx', 'Monto_Neto']])
        with t3:
            st.dataframe(solo_aux[['Fecha elaboración', 'Comprobante', 'Nombre del tercero', 'Débito', 'Crédito', 'Monto_Neto']])

# ==========================================
# MÓDULO 2: CONCILIACIÓN BANCARIA MENSUAL
# ==========================================
elif opcion == "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)":
    st.title("📊 Conciliación Bancaria Mensual")
    st.write("Carga el **Extracto Bancario (.xlsx)** oficial y el **Auxiliar Contable (.xlsx)** del mes.")

    col1, col2 = st.columns(2)
    with col1:
        file_ext = st.file_uploader("1. Cargar Extracto Bancario (.xlsx)", type=["xlsx"], key="ext_mensual")
    with col2:
        file_auxiliar = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx"], key="aux_mensual")

    if file_ext and file_auxiliar:
        df_ext_raw = pd.read_excel(file_ext, header=None)
        df_aux_raw = pd.read_excel(file_auxiliar)

        df_ext = procesar_extracto_excel(df_ext_raw)
        df_aux = procesar_auxiliar(df_aux_raw)

        # Generar Llave de Cruce por Monto Absoluto y Ocurrencia
        df_ext['Ocurrencia'] = df_ext.groupby(['Monto_Abs']).cumcount() + 1
        df_aux['Ocurrencia'] = df_aux.groupby(['Monto_Abs']).cumcount() + 1

        df_ext['LLAVE'] = df_ext['Monto_Abs'].astype(str) + "_" + df_ext['Ocurrencia'].astype(str)
        df_aux['LLAVE'] = df_aux['Monto_Abs'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)

        conciliados = pd.merge(df_ext, df_aux, on='LLAVE', how='inner', suffixes=('_Extracto', '_Auxiliar'))
        solo_ext = df_ext[~df_ext['LLAVE'].isin(conciliados['LLAVE'])]
        solo_aux = df_aux[~df_aux['LLAVE'].isin(conciliados['LLAVE'])]

        st.success("✅ Conciliación Mensual realizada con éxito")

        k1, k2, k3 = st.columns(3)
        k1.metric("Movimientos Conciliados", len(conciliados))
        k2.metric("Pendientes en Banco (Extracto)", len(solo_ext))
        k3.metric("Pendientes en Contabilidad", len(solo_aux))

        tab1, tab2, tab3 = st.tabs(["✅ Conciliados", "⚠️ Pendientes Extracto Banco", "⚠️ Pendientes Contabilidad"])
        with tab1:
            st.dataframe(conciliados[['FECHA', 'DESCRIPCIÓN', 'VALOR', 'Comprobante', 'Nombre del tercero', 'Monto_Neto_Auxiliar']])
        with tab2:
            st.dataframe(solo_ext[['FECHA', 'DESCRIPCIÓN', 'VALOR', 'SALDO']])
        with tab3:
            st.dataframe(solo_aux[['Fecha elaboración', 'Comprobante', 'Nombre del tercero', 'Débito', 'Crédito', 'Monto_Neto']])
