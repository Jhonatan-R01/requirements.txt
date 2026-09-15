import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import streamlit as st

# Configuración de la página web
st.set_page_config(page_title="Sistema de Conciliación & Control Diario", layout="wide")

# ==========================================
# FUNCIÓN PARA GENERAR PLANTILLA EXCEL
# ==========================================
def generar_excel_plantilla(df_aux_proc, df_ext_proc, buffer):
    wb = openpyxl.Workbook()

    # --- 1. HOJA: Resumen Conciliación ---
    ws_resumen = wb.active
    ws_resumen.title = "Resumen Conciliación"
    ws_resumen.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)

    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_subtotal = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    border_grid = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    ws_resumen['A2'] = "CONCILIACIÓN BANCARIA AUTOMATIZADA"
    ws_resumen['A2'].font = font_title
    ws_resumen['B2'] = "Egresos"
    ws_resumen['B2'].font = font_bold

    ws_resumen['A4'], ws_resumen['B4'] = "Mes de Conciliación:", 2026
    ws_resumen['A4'].font = font_bold

    ws_resumen['A5'], ws_resumen['B5'], ws_resumen['C5'] = "Cuenta Bancaria:", "Banco Principal Ahorros No. ", "167-000060-56"
    ws_resumen['A5'].font = font_bold

    headers_res = ["CONCEPTO", "VALOR (COP / USD)", "NOTAS / AUDITORÍA"]
    for col_idx, text in enumerate(headers_res, 1):
        cell = ws_resumen.cell(row=8, column=col_idx, value=text)
        cell.font, cell.fill, cell.alignment = font_header, fill_header, Alignment(horizontal="center", vertical="center")

    filas = [
        ("Saldo Mov según Auxiliar Contable", "=+Tabla1[[#Totals],[Egresos (-)]]", None),
        ("(+) Egresos No registrados por el Banco", '=SUMIF(Tabla1[Extracto],"NO ESTA EN BANCOS",Tabla1[Egresos (-)])', "Ingresos/Egresos en libros pendientes en extracto"),
        ("SALDO CONTABLE AJUSTADO", "=B9-B10", "Saldo conciliado contable"),
        ("Saldo Final según Extracto Bancario", "=+Tabla2[[#Totals],[Retiros ()]]", None),
        ("(-) Notas Débito no Contabilizadas", '=SUMIF(Tabla2[Contabilidad],"Pen Contabilidad",Tabla2[Retiros ()])', "Pagos directos del banco faltantes en contabilidad"),
        ("SALDO BANCARIO AJUSTADO", "=+B12-B13", "Saldo conciliado bancario"),
        ("DIFERENCIA FINAL POR CONCILIAR", "=+B12-B9", "Diferencia por conciliar")
    ]

    for r_idx, (concepto, formula, nota) in enumerate(filas, start=9):
        c1 = ws_resumen.cell(row=r_idx, column=1, value=concepto)
        c2 = ws_resumen.cell(row=r_idx, column=2, value=formula)
        c3 = ws_resumen.cell(row=r_idx, column=3, value=nota)

        c2.number_format = '$#,##0.00'
        for cell in (c1, c2, c3):
            cell.border = border_grid
            cell.font = font_regular

        if r_idx in (11, 14):
            for cell in (c1, c2, c3): cell.fill, cell.font = fill_subtotal, font_bold
        elif r_idx == 15:
            for cell in (c1, c2, c3): cell.fill, cell.font = fill_highlight, font_bold

    ws_resumen.column_dimensions['A'].width = 42
    ws_resumen.column_dimensions['B'].width = 25
    ws_resumen.column_dimensions['C'].width = 45

    # --- 2. HOJA: Auxiliar Contable ---
    ws_aux = wb.create_sheet(title="Auxiliar Contable")
    ws_aux.views.sheetView[0].showGridLines = True
    ws_aux.append(["Fecha", "Documento", "Concepto/Detalle", "NOMBRE", "Egresos (-)", "LLAVE", "Extracto"])

    for idx, row in df_aux_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        fecha_val = row.get('Fecha_Clean', row.get('Fecha elaboración', ''))
        doc_val = row.get('Comprobante', '')
        concepto_val = row.get('Concepto', row.get('Código contable', ''))
        nombre_val = row.get('Nombre del tercero', '')
        monto_val = float(row.get('Monto_Abs', 0))

        ws_aux.append([
            str(fecha_val),
            str(doc_val),
            str(concepto_val),
            str(nombre_val),
            monto_val,
            f'=CONCATENATE(A{r},E{r},"-",COUNTIFS($A$2:A{r},A{r},$E$2:E{r},E{r}))',
            f'=_xlfn.XLOOKUP(F{r}, \'Extracto Bancario\'!$D:$D, \'Extracto Bancario\'!$B:$B, "NO ESTA EN BANCOS", 0)'
        ])

    max_aux = max(ws_aux.max_row, 2)
    tab1 = Table(displayName="Tabla1", ref=f"A1:G{max_aux}")
    tab1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_aux.add_table(tab1)

    # --- 3. HOJA: Extracto Bancario ---
    ws_ext = wb.create_sheet(title="Extracto Bancario")
    ws_ext.views.sheetView[0].showGridLines = True
    ws_ext.append(["Fecha", "Referencia", "Retiros ()", "LLAVE", "Contabilidad", "Columna1", "Columna2"])

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        fecha_val = row.get('Fecha_Clean', row.get('FECHA', ''))
        ref_val = row.get('Descripcion', row.get('DESCRIPCIÓN', ''))
        monto_val = float(row.get('Monto_Abs', 0))

        ws_ext.append([
            str(fecha_val),
            str(ref_val),
            monto_val,
            f'=CONCATENATE(A{r},C{r},"-",COUNTIFS($A$2:A{r},A{r},$C$2:C{r},C{r}))',
            f'=_xlfn.XLOOKUP(D{r}, \'Auxiliar Contable\'!$F:$F, \'Auxiliar Contable\'!$B:$B, "Pen Contabilidad", 0)',
            "", ""
        ])

    max_ext = max(ws_ext.max_row, 2)
    tab2 = Table(displayName="Tabla2", ref=f"A1:G{max_ext}")
    tab2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_ext.add_table(tab2)

    wb.save(buffer)

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
    headers = df_raw.iloc[6].values
    df = df_raw.iloc[8:].copy()
    df.columns = headers
    df = df.dropna(subset=['Código contable', 'Fecha elaboración'])
    
    df['Débito_Clean'] = df['Débito'].apply(limpiar_valor)
    df['Crédito_Clean'] = df['Crédito'].apply(limpiar_valor)
    df['Monto_Neto'] = df['Débito_Clean'] - df['Crédito_Clean']
    df['Monto_Abs'] = df['Monto_Neto'].abs().round(2)
    
    df['Fecha_Clean'] = pd.to_datetime(df['Fecha elaboración'], format='%d/%m/%Y', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_diario_csv(df_raw):
    """Procesa el CSV de movimiento diario sin encabezados fijados"""
    df = df_raw.iloc[:, [0, 1, 3, 5, 6, 7]].copy()
    df.columns = ['Cuenta', 'Oficina', 'Fecha_Int', 'Valor_Raw', 'Codigo_Tx', 'Descripcion']
    
    df['Monto_Neto'] = df['Valor_Raw'].apply(limpiar_valor)
    df['Monto_Abs'] = df['Monto_Neto'].abs().round(2)
    df['Fecha_Clean'] = pd.to_datetime(df['Fecha_Int'].astype(str), format='%Y%m%d', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_extracto_excel(df_raw):
    """Procesa el extracto bancario en Excel"""
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

        df_diario['Ocurrencia'] = df_diario.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1
        df_aux['Ocurrencia'] = df_aux.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1

        df_diario['LLAVE'] = df_diario['Fecha_Clean'] + "_" + df_diario['Monto_Abs'].astype(str) + "_" + df_diario['Ocurrencia'].astype(str)
        df_aux['LLAVE'] = df_aux['Fecha_Clean'] + "_" + df_aux['Monto_Abs'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)

        cruzados = pd.merge(df_diario, df_aux, on='LLAVE', how='inner', suffixes=('_Diario', '_Auxiliar'))
        solo_diario = df_diario[~df_diario['LLAVE'].isin(cruzados['LLAVE'])]
        solo_aux = df_aux[~df_aux['LLAVE'].isin(cruzados['LLAVE'])]

        st.success("✅ Cruce Diario ejecutado correctamente")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Partidas Cruzadas", len(cruzados))
        m2.metric("Pendientes en Registro Diario", len(solo_diario))
        m3.metric("Pendientes en Contabilidad", len(solo_aux))

        # Botón para descargar el Excel con formato
        excel_buffer = io.BytesIO()
        generar_excel_plantilla(df_aux, df_diario, excel_buffer)
        
        st.markdown("---")
        st.download_button(
            label="📥 Descargar Reporte Conciliado en Excel (Formato Plantilla)",
            data=excel_buffer.getvalue(),
            file_name="Conciliacion_Diaria_Procesada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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

        # Botón para descargar el Excel con formato
        excel_buffer_m = io.BytesIO()
        generar_excel_plantilla(df_aux, df_ext, excel_buffer_m)

        st.markdown("---")
        st.download_button(
            label="📥 Descargar Conciliación Mensual en Excel (Formato Plantilla)",
            data=excel_buffer_m.getvalue(),
            file_name="Conciliacion_Mensual_Procesada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        tab1, tab2, tab3 = st.tabs(["✅ Conciliados", "⚠️ Pendientes Extracto Banco", "⚠️ Pendientes Contabilidad"])
        with tab1:
            st.dataframe(conciliados[['FECHA', 'DESCRIPCIÓN', 'VALOR', 'Comprobante', 'Nombre del tercero', 'Monto_Neto_Auxiliar']])
        with tab2:
            st.dataframe(solo_ext[['FECHA', 'DESCRIPCIÓN', 'VALOR', 'SALDO']])
        with tab3:
            st.dataframe(solo_aux[['Fecha elaboración', 'Comprobante', 'Nombre del tercero', 'Débito', 'Crédito', 'Monto_Neto']])
