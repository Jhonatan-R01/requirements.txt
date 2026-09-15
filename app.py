import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import streamlit as st

# Configuración de la página web
st.set_page_config(page_title="Sistema de Conciliación & Control Diario", layout="wide")

# ==========================================
# 1. FUNCIÓN PARA GENERAR PLANTILLA EXCEL
# ==========================================
def generar_excel_plantilla(df_aux_proc, df_ext_proc, buffer, tipo_reporte="General"):
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

    ws_resumen['A2'] = f"CONCILIACIÓN BANCARIA AUTOMATIZADA - REPORTE {tipo_reporte.upper()}"
    ws_resumen['A2'].font = font_title
    ws_resumen['B2'] = tipo_reporte
    ws_resumen['B2'].font = font_bold

    ws_resumen['A4'], ws_resumen['B4'] = "Fecha / Período:", 2026
    ws_resumen['A4'].font = font_bold

    ws_resumen['A5'], ws_resumen['B5'], ws_resumen['C5'] = "Cuenta Bancaria:", "Banco Principal", "Cta. Ahorros / Corriente"
    ws_resumen['A5'].font = font_bold

    headers_res = ["CONCEPTO", "VALOR (COP / USD)", "NOTAS / AUDITORÍA"]
    for col_idx, text in enumerate(headers_res, 1):
        cell = ws_resumen.cell(row=8, column=col_idx, value=text)
        cell.font, cell.fill, cell.alignment = font_header, fill_header, Alignment(horizontal="center", vertical="center")

    filas = [
        ("Saldo Mov según Auxiliar Contable", "=+Tabla1[[#Totals],[Monto (+/-)]]", None),
        ("(+) Partidas No registradas por el Banco", '=SUMIF(Tabla1[Extracto],"NO ESTA EN BANCOS",Tabla1[Monto (+/-)])', "Movimientos en libros pendientes en extracto"),
        ("SALDO CONTABLE AJUSTADO", "=B9-B10", "Saldo conciliado contable"),
        ("Saldo Final según Extracto Bancario", "=+Tabla2[[#Totals],[Monto (+/-)]]", None),
        ("(-) Partidas no Contabilizadas", '=SUMIF(Tabla2[Contabilidad],"Pen Contabilidad",Tabla2[Monto (+/-)])', "Movimientos del banco faltantes en contabilidad"),
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
    ws_aux.append(["Fecha", "Documento", "Concepto/Detalle", "NOMBRE", "Naturaleza", "Monto (+/-)", "LLAVE", "Extracto"])

    for idx, row in df_aux_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        fecha_val = row.get('Fecha_Clean', row.get('Fecha elaboración', ''))
        doc_val = row.get('Comprobante', '')
        concepto_val = row.get('Concepto', row.get('Código contable', ''))
        nombre_val = row.get('Nombre del tercero', '')
        tipo_val = row.get('Naturaleza', 'General')
        monto_val = float(row.get('Monto_Neto', 0))

        ws_aux.append([
            str(fecha_val),
            str(doc_val),
            str(concepto_val),
            str(nombre_val),
            str(tipo_val),
            monto_val,
            f'=CONCATENATE(A{r},F{r},"-",COUNTIFS($A$2:A{r},A{r},$F$2:F{r},F{r}))',
            f'=IFISNA(VLOOKUP(G{r}, \'Extracto Bancario\'!$D:$B, 1, FALSE), "NO ESTA EN BANCOS")'
        ])

    max_aux = max(ws_aux.max_row, 2)
    tab1 = Table(displayName="Tabla1", ref=f"A1:H{max_aux}")
    tab1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_aux.add_table(tab1)

    # --- 3. HOJA: Extracto Bancario ---
    ws_ext = wb.create_sheet(title="Extracto Bancario")
    ws_ext.views.sheetView[0].showGridLines = True
    ws_ext.append(["Fecha", "Referencia", "Naturaleza", "LLAVE", "Contabilidad"])

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        fecha_val = row.get('Fecha_Clean', row.get('FECHA', ''))
        ref_val = row.get('Descripcion', row.get('DESCRIPCIÓN', ''))
        tipo_val = row.get('Naturaleza', 'General')
        monto_val = float(row.get('Monto_Neto', 0))

        ws_ext.append([
            str(fecha_val),
            str(ref_val),
            str(tipo_val),
            monto_val,
            f'=CONCATENATE(A{r},D{r},"-",COUNTIFS($A$2:A{r},A{r},$D$2:D{r},D{r}))',
            f'=IFISNA(VLOOKUP(E{r}, \'Auxiliar Contable\'!$G:$B, 1, FALSE), "Pen Contabilidad")'
        ])

    max_ext = max(ws_ext.max_row, 2)
    tab2 = Table(displayName="Tabla2", ref=f"A1:E{max_ext}")
    tab2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_ext.add_table(tab2)

    wb.save(buffer)

# ==========================================
# 2. FUNCIONES DE PROCESAMIENTO Y LIMPIEZA
# ==========================================
def limpiar_valor(val):
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
    headers = df_raw.iloc[6].values
    df = df_raw.iloc[8:].copy()
    df.columns = headers
    df = df.dropna(subset=['Código contable', 'Fecha elaboración'])
    
    df['Débito_Clean'] = df['Débito'].apply(limpiar_valor)
    df['Crédito_Clean'] = df['Crédito'].apply(limpiar_valor)
    
    # Calcular Neto y Naturaleza
    df['Monto_Neto'] = df['Débito_Clean'] - df['Crédito_Clean']
    df['Monto_Abs'] = df['Monto_Neto'].abs().round(2)
    df['Naturaleza'] = df['Monto_Neto'].apply(lambda x: "Ingreso" if x > 0 else "Egreso")
    
    df['Fecha_Clean'] = pd.to_datetime(df['Fecha elaboration' if 'Fecha elaboration' in df.columns else 'Fecha elaboración'], format='%d/%m/%Y', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_diario_csv(df_raw):
    df = df_raw.iloc[:, [0, 1, 3, 5, 6, 7]].copy()
    df.columns = ['Cuenta', 'Oficina', 'Fecha_Int', 'Valor_Raw', 'Codigo_Tx', 'Descripcion']
    
    df['Monto_Neto'] = df['Valor_Raw'].apply(limpiar_valor)
    df['Monto_Abs'] = df['Monto_Neto'].abs().round(2)
    df['Naturaleza'] = df['Monto_Neto'].apply(lambda x: "Ingreso" if x > 0 else "Egreso")
    df['Fecha_Clean'] = pd.to_datetime(df['Fecha_Int'].astype(str), format='%Y%m%d', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_extracto_excel(df_raw):
    df = df_raw.iloc[15:].copy()
    df.columns = ['FECHA', 'DESCRIPCIÓN', 'SUCURSAL', 'DCTO', 'VALOR', 'SALDO', 'X1', 'X2']
    df = df.dropna(subset=['FECHA', 'VALOR'])
    df = df[~df['FECHA'].astype(str).str.contains('FECHA|FIN ESTADO', case=False, na=False)]
    
    df['Monto_Neto'] = df['VALOR'].apply(limpiar_valor)
    df['Monto_Abs'] = df['Monto_Neto'].abs().round(2)
    df['Naturaleza'] = df['Monto_Neto'].apply(lambda x: "Ingreso" if x > 0 else "Egreso")
    return df

# ==========================================
# 3. MENÚ LATERAL Y MÓDULOS
# ==========================================
st.sidebar.title("📌 Menú Principal")
opcion = st.sidebar.radio(
    "Selecciona el módulo:",
    ["📖 1. Cruce Diario (CSV vs Auxiliar)", "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)"]
)

if opcion == "📖 1. Cruce Diario (CSV vs Auxiliar)":
    st.title("📖 Cruce Operativo Diario")
    st.write("Carga el movimiento diario registrado en **CSV** y el **Auxiliar Contable**.")

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

        st.success("✅ Datos procesados con éxito.")

        st.markdown("### 📥 Descargas Disponibles")
        c_desc1, c_desc2, c_desc3 = st.columns(3)

        # 1. Descarga Egresos
        with c_desc1:
            buf_egr = io.BytesIO()
            df_a_egr = df_aux[df_aux['Naturaleza'] == 'Egreso']
            df_d_egr = df_diario[df_diario['Naturaleza'] == 'Egreso']
            generar_excel_plantilla(df_a_egr, df_d_egr, buf_egr, tipo_reporte="Egresos")
            st.download_button("🔻 Descargar EGRESOS (.xlsx)", data=buf_egr.getvalue(), file_name="Egresos_Diario.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # 2. Descarga Ingresos
        with c_desc2:
            buf_ing = io.BytesIO()
            df_a_ing = df_aux[df_aux['Naturaleza'] == 'Ingreso']
            df_d_ing = df_diario[df_diario['Naturaleza'] == 'Ingreso']
            generar_excel_plantilla(df_a_ing, df_d_ing, buf_ing, tipo_reporte="Ingresos")
            st.download_button("🟢 Descargar INGRESOS (.xlsx)", data=buf_ing.getvalue(), file_name="Ingresos_Diario.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # 3. Descarga General
        with c_desc3:
            buf_gen = io.BytesIO()
            generar_excel_plantilla(df_aux, df_diario, buf_gen, tipo_reporte="General_Consolidado")
            st.download_button("📦 Descargar GENERAL (.xlsx)", data=buf_gen.getvalue(), file_name="General_Diario.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("---")
        st.write("### 👁️ Vista Previa Consolidada")
        st.dataframe(df_aux[['Fecha_Clean', 'Comprobante', 'Nombre del tercero', 'Monto_Neto', 'Naturaleza']].head(10))

elif opcion == "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)":
    st.title("📊 Conciliación Bancaria Mensual")
    st.write("Carga el **Extracto Bancario (.xlsx)** y el **Auxiliar Contable (.xlsx)**.")

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

        st.success("✅ Conciliación calculada correctamente.")

        st.markdown("### 📥 Descargas Disponibles")
        c_desc1, c_desc2, c_desc3 = st.columns(3)

        with c_desc1:
            buf_egr_m = io.BytesIO()
            df_a_egr = df_aux[df_aux['Naturaleza'] == 'Egreso']
            df_e_egr = df_ext[df_ext['Naturaleza'] == 'Egreso']
            generar_excel_plantilla(df_a_egr, df_e_egr, buf_egr_m, tipo_reporte="Egresos")
            st.download_button("🔻 Descargar EGRESOS (.xlsx)", data=buf_egr_m.getvalue(), file_name="Conciliacion_Egresos_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c_desc2:
            buf_ing_m = io.BytesIO()
            df_a_ing = df_aux[df_aux['Naturaleza'] == 'Ingreso']
            df_e_ing = df_ext[df_ext['Naturaleza'] == 'Ingreso']
            generar_excel_plantilla(df_a_ing, df_e_ing, buf_ing_m, tipo_reporte="Ingresos")
            st.download_button("🟢 Descargar INGRESOS (.xlsx)", data=buf_ing_m.getvalue(), file_name="Conciliacion_Ingresos_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c_desc3:
            buf_gen_m = io.BytesIO()
            generar_excel_plantilla(df_aux, df_ext, buf_gen_m, tipo_reporte="General_Consolidado")
            st.download_button("📦 Descargar GENERAL (.xlsx)", data=buf_gen_m.getvalue(), file_name="Conciliacion_General_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("---")
        st.write("### 👁️ Vista Previa Extracto Procesado")
        st.dataframe(df_ext[['FECHA', 'DESCRIPCIÓN', 'VALOR', 'Naturaleza']].head(10))
