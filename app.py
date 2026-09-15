import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import streamlit as st

# Configuración de página
st.set_page_config(page_title="Sistema de Conciliación", layout="wide")

# ==========================================
# 1. FUNCIÓN QUE GENERA EL EXCEL CONCILIADO
# ==========================================
def generar_excel_conciliado_con_datos(df_aux, df_ext, buffer):
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

    for idx, row in df_aux.iterrows():
        r = idx + 2
        ws_aux.append([
            str(row.get('Fecha', '')),
            str(row.get('Documento', '')),
            str(row.get('Concepto/Detalle', '')),
            str(row.get('NOMBRE', '')),
            float(row.get('Egresos (-)', 0)) if pd.notnull(row.get('Egresos (-)')) else 0,
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

    for idx, row in df_ext.iterrows():
        r = idx + 2
        ws_ext.append([
            str(row.get('Fecha', '')),
            str(row.get('Referencia', '')),
            float(row.get('Retiros ()', 0)) if pd.notnull(row.get('Retiros ()')) else 0,
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
# 2. MENÚ LATERAL DE STREAMLIT
# ==========================================
st.sidebar.title("📌 Menú Principal")

opcion = st.sidebar.radio(
    "Selecciona el módulo:",
    ["📖 1. Cruce Diario (CSV vs Auxiliar)", "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)"]
)

# ==========================================
# 3. INTERFAZ SEGÚN LA OPCIÓN SELECCIONADA
# ==========================================

# --- MÓDULO 1: CRUCE DIARIO ---
if opcion == "📖 1. Cruce Diario (CSV vs Auxiliar)":
    st.title("📖 Cruce Operativo Diario")
    st.write("Carga el movimiento diario registrado en **CSV** y el **Auxiliar Contable** para comparar transacciones del día.")

    col1, col2 = st.columns(2)
    with col1:
        archivo_mov = st.file_uploader("1. Cargar Movimiento Diario (.csv)", type=["csv"], key="cruce_mov")
    with col2:
        archivo_aux = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx", "xls"], key="cruce_aux")

    if archivo_mov is not None and archivo_aux is not None:
        try:
            # Reintentos de codificación para archivos CSV con tildes/eñes
            try:
                df_mov_data = pd.read_csv(archivo_mov, encoding='utf-8')
            except UnicodeDecodeError:
                archivo_mov.seek(0)
                try:
                    df_mov_data = pd.read_csv(archivo_mov, encoding='latin-1')
                except UnicodeDecodeError:
                    archivo_mov.seek(0)
                    df_mov_data = pd.read_csv(archivo_mov, encoding='utf-8-sig')

            df_aux_data = pd.read_excel(archivo_aux)

            st.success("✅ Archivos leídos exitosamente.")

            excel_resultado = io.BytesIO()
            generar_excel_conciliado_con_datos(df_aux_data, df_mov_data, excel_resultado)

            st.markdown("---")
            st.subheader("🚀 Descargar Resultado Conciliado")
            st.download_button(
                label="📥 Descargar Archivo Excel Conciliado",
                data=excel_resultado.getvalue(),
                file_name="Conciliacion_Egresos_Procesada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error al procesar los archivos: {e}")

# --- MÓDULO 2: CONCILIACIÓN MENSUAL ---
elif opcion == "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)":
    st.title("📊 Conciliación Bancaria Mensual")
    st.write("Carga el extracto mensual en **Excel** y el **Auxiliar Contable** acumulado.")

    col1, col2 = st.columns(2)
    with col1:
        archivo_ext_m = st.file_uploader("1. Cargar Extracto Bancario (.xlsx)", type=["xlsx", "xls"], key="mensual_ext")
    with col2:
        archivo_aux_m = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx", "xls"], key="mensual_aux")

    if archivo_ext_m is not None and archivo_aux_m is not None:
        try:
            df_ext_m = pd.read_excel(archivo_ext_m)
            df_aux_m = pd.read_excel(archivo_aux_m)

            st.success("✅ Archivos mensuales recibidos.")

            excel_m_resultado = io.BytesIO()
            generar_excel_conciliado_con_datos(df_aux_m, df_ext_m, excel_m_resultado)

            st.markdown("---")
            st.subheader("🚀 Descargar Resultado Mensual")
            st.download_button(
                label="📥 Descargar Conciliación Mensual (Excel)",
                data=excel_m_resultado.getvalue(),
                file_name="Conciliacion_Mensual_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error al procesar la conciliación mensual: {e}")
