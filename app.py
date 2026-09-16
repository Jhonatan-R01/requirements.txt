import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(page_title="Sistema de Conciliación & Control Diario", layout="wide")

# ==========================================
# 1. ALGORITMO DE SUGERENCIAS INTELIGENTES
# ==========================================
def obtener_sugerencias_cruce(df_aux_pend, df_ext_pend, tol_monto=10.0):
    sugerencias = []
    
    if df_aux_pend.empty or df_ext_pend.empty:
        return pd.DataFrame(sugerencias)

    aux_tmp = df_aux_pend.copy()
    ext_tmp = df_ext_pend.copy()
    
    col_f_aux = 'Fecha_Clean' if 'Fecha_Clean' in aux_tmp.columns else aux_tmp.columns[0]
    col_f_ext = 'Fecha_Clean' if 'Fecha_Clean' in ext_tmp.columns else ext_tmp.columns[0]
    
    aux_tmp['Fecha_dt'] = pd.to_datetime(aux_tmp[col_f_aux], errors='coerce')
    ext_tmp['Fecha_dt'] = pd.to_datetime(ext_tmp[col_f_ext], errors='coerce')

    ext_usados = set()

    for idx_a, row_a in aux_tmp.iterrows():
        monto_a = float(row_a.get('Monto_Neto', 0))
        fecha_a = row_a['Fecha_dt']

        candidatos = []
        for idx_e, row_e in ext_tmp.iterrows():
            if idx_e in ext_usados:
                continue
                
            monto_e = float(row_e.get('Monto_Neto', 0))
            fecha_e = row_e['Fecha_dt']

            diff_monto = abs(monto_a - monto_e)
            
            if diff_monto <= tol_monto:
                dias_desfase = abs((fecha_e - fecha_a).days) if (pd.notna(fecha_a) and pd.notna(fecha_e)) else 0
                ref_banco = str(row_e.get('Descripcion', row_e.get('DESCRIPCIÓN', '')))
                fecha_banco_str = str(row_e.get(col_f_ext, ''))
                
                candidatos.append({
                    'idx_e': idx_e,
                    'Fecha_Banco': fecha_banco_str,
                    'Ref_Banco': ref_banco,
                    'Monto_Banco': monto_e,
                    'Diff_Monto': diff_monto,
                    'Dias_Desfase': dias_desfase
                })

        if candidatos:
            candidatos.sort(key=lambda x: (x['Diff_Monto'], x['Dias_Desfase']))
            mejor = candidatos[0]
            ext_usados.add(mejor['idx_e'])
            
            certeza = "Alta (Exacto)" if mejor['Diff_Monto'] == 0 else f"Media (Diff: ${mejor['Diff_Monto']:.2f})"
            fecha_contable_str = str(row_a.get(col_f_aux, ''))

            sugerencias.append({
                'Fecha Contable': fecha_contable_str,
                'Documento': str(row_a.get('Comprobante', '')),
                'Tercero': str(row_a.get('Nombre del tercero', '')),
                'Monto Contable': monto_a,
                'Fecha Banco': mejor['Fecha_Banco'],
                'Referencia Banco': mejor['Ref_Banco'],
                'Monto Banco': mejor['Monto_Banco'],
                'Diferencia ($)': mejor['Diff_Monto'],
                'Días Desfase': mejor['Dias_Desfase'],
                'Certeza': certeza
            })

    return pd.DataFrame(sugerencias)

# ==========================================
# 2. FUNCIÓN PARA GENERAR EXCEL MULTI-PESTAÑA
# ==========================================
def generar_excel_plantilla(df_aux_proc, df_ext_proc, df_sugerencias, buffer, tipo_reporte="General"):
    wb = openpyxl.Workbook()

    # --- HOJA 1: Resumen Conciliación ---
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

    headers_res = ["CONCEPTO", "VALOR (COP / USD)", "NOTAS / AUDITORÍA"]
    for col_idx, text in enumerate(headers_res, 1):
        cell = ws_resumen.cell(row=8, column=col_idx, value=text)
        cell.font, cell.fill, cell.alignment = font_header, fill_header, Alignment(horizontal="center", vertical="center")

    max_aux = len(df_aux_proc) + 1
    max_ext = len(df_ext_proc) + 1

    filas = [
        ("Saldo Mov según Auxiliar Contable", f"=SUM('Auxiliar Contable'!E2:E{max(max_aux, 2)})", None),
        ("(+) Partidas No registradas por el Banco", f'=SUMIF(\'Auxiliar Contable\'!G2:G{max(max_aux, 2)}, "NO ESTA EN BANCOS", \'Auxiliar Contable\'!E2:E{max(max_aux, 2)})', "Movimientos en libros pendientes en extracto"),
        ("SALDO CONTABLE AJUSTADO", "=B9-B10", "Saldo conciliado contable"),
        ("Saldo Final según Extracto Bancario", f"=SUM('Extracto Bancario'!C2:C{max(max_ext, 2)})", None),
        ("(-) Partidas no Contabilizadas", f'=SUMIF(\'Extracto Bancario\'!E2:E{max(max_ext, 2)}, "Pen Contabilidad", \'Extracto Bancario\'!C2:C{max(max_ext, 2)})', "Movimientos del banco faltantes en contabilidad"),
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

    # --- HOJA 2: Auxiliar Contable ---
    ws_aux = wb.create_sheet(title="Auxiliar Contable")
    ws_aux.views.sheetView[0].showGridLines = True
    ws_aux.append(["Fecha", "Documento", "Concepto/Detalle", "NOMBRE", "Monto (+/-)", "LLAVE", "Extracto"])

    for idx, row in df_aux_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        ws_aux.append([
            str(row.get('Fecha_Clean', '')),
            str(row.get('Comprobante', '')),
            str(row.get('Concepto', '')),
            str(row.get('Nombre del tercero', '')),
            float(row.get('Monto_Neto', 0)),
            str(row.get('LLAVE', '')),
            "CONCILIADO" if row.get('Conciliado', False) else "NO ESTA EN BANCOS"
        ])

    max_aux_tbl = max(ws_aux.max_row, 2)
    tab1 = Table(displayName="Tabla1", ref=f"A1:G{max_aux_tbl}")
    tab1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_aux.add_table(tab1)

    # --- HOJA 3: Extracto Bancario ---
    ws_ext = wb.create_sheet(title="Extracto Bancario")
    ws_ext.views.sheetView[0].showGridLines = True
    ws_ext.append(["Fecha", "Referencia", "Monto (+/-)", "LLAVE", "Contabilidad"])

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        ws_ext.append([
            str(row.get('Fecha_Clean', '')),
            str(row.get('Descripcion', '')),
            float(row.get('Monto_Neto', 0)),
            str(row.get('LLAVE', '')),
            "CONCILIADO" if row.get('Conciliado', False) else "Pen Contabilidad"
        ])

    max_ext_tbl = max(ws_ext.max_row, 2)
    tab2 = Table(displayName="Tabla2", ref=f"A1:E{max_ext_tbl}")
    tab2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_ext.add_table(tab2)

    # --- HOJA 4: Sugerencias de Cruce ---
    ws_sug = wb.create_sheet(title="Sugerencias de Cruce")
    ws_sug.views.sheetView[0].showGridLines = True
    
    headers_sug = ["Fecha Contable", "Documento", "Tercero", "Monto Contable", "Fecha Banco", "Referencia Banco", "Monto Banco", "Diferencia ($)", "Días Desfase", "Certeza"]
    ws_sug.append(headers_sug)

    if not df_sugerencias.empty:
        for _, r_sug in df_sugerencias.iterrows():
            ws_sug.append([
                str(r_sug['Fecha Contable']),
                str(r_sug['Documento']),
                str(r_sug['Tercero']),
                float(r_sug['Monto Contable']),
                str(r_sug['Fecha Banco']),
                str(r_sug['Referencia Banco']),
                float(r_sug['Monto Banco']),
                float(r_sug['Diferencia ($)']),
                int(r_sug['Días Desfase']),
                str(r_sug['Certeza'])
            ])
    
    max_sug_tbl = max(ws_sug.max_row, 2)
    tab3 = Table(displayName="Tabla3", ref=f"A1:J{max_sug_tbl}")
    tab3.tableStyleInfo = TableStyleInfo(name="TableStyleMedium3", showRowStripes=True)
    ws_sug.add_table(tab3)

    wb.save(buffer)

# ==========================================
# 3. FUNCIONES DE PROCESAMIENTO
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
    idx_header = None
    for i, row in df_raw.iterrows():
        row_str = [str(cell).lower() for cell in row.values]
        if any('debito' in cell or 'débito' in cell for cell in row_str):
            idx_header = i
            break
            
    if idx_header is not None:
        df_raw.columns = df_raw.iloc[idx_header].values
        df = df_raw.iloc[idx_header + 1:].copy()
    else:
        df = df_raw.copy()

    col_debito = next((c for c in df.columns if 'débito' in str(c).lower() or 'debito' in str(c).lower()), None)
    col_credito = next((c for c in df.columns if 'crédito' in str(c).lower() or 'credito' in str(c).lower()), None)
    col_fecha = next((c for c in df.columns if 'fecha' in str(c).lower()), None)
    
    if col_debito and col_credito:
        df['Débito_Clean'] = df[col_debito].apply(limpiar_valor)
        df['Crédito_Clean'] = df[col_credito].apply(limpiar_valor)
        df['Monto_Neto'] = df['Débito_Clean'] - df['Crédito_Clean']
    else:
        df['Monto_Neto'] = 0.0

    if col_fecha:
        df['Fecha_Clean'] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce').dt.strftime('%Y-%m-%d')
    else:
        df['Fecha_Clean'] = None
        
    return df.dropna(subset=['Monto_Neto'])

def procesar_diario_csv(df_raw):
    df = df_raw.iloc[:, [0, 1, 3, 5, 6, 7]].copy()
    df.columns = ['Cuenta', 'Oficina', 'Fecha_Int', 'Valor_Raw', 'Codigo_Tx', 'Descripcion']
    df['Monto_Neto'] = df['Valor_Raw'].apply(limpiar_valor)
    df['Fecha_Clean'] = pd.to_datetime(df['Fecha_Int'].astype(str), format='%Y%m%d', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_extracto_excel(df_raw):
    # Detección dinámica de la fila de encabezados en el Extracto Bancario
    idx_header = None
    for i, row in df_raw.iterrows():
        row_str = [str(cell).lower() for cell in row.values]
        if any('valor' in cell or 'monto' in cell or 'saldo' in cell for cell in row_str):
            idx_header = i
            break

    if idx_header is not None:
        df_raw.columns = df_raw.iloc[idx_header].values
        df = df_raw.iloc[idx_header + 1:].copy()
    else:
        df = df_raw.copy()

    col_valor = next((c for c in df.columns if 'valor' in str(c).lower() or 'monto' in str(c).lower()), None)
    col_fecha = next((c for c in df.columns if 'fecha' in str(c).lower()), None)
    col_desc = next((c for c in df.columns if 'descrip' in str(c).lower() or 'detalle' in str(c).lower() or 'concepto' in str(c).lower()), None)

    if col_valor:
        df['Monto_Neto'] = df[col_valor].apply(limpiar_valor)
    else:
        df['Monto_Neto'] = 0.0

    if col_fecha:
        df['Fecha_Clean'] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce').dt.strftime('%Y-%m-%d')
    else:
        df['Fecha_Clean'] = None

    df['Descripcion'] = df[col_desc].astype(str) if col_desc else ''
    
    return df.dropna(subset=['Monto_Neto'])

def ejecutar_conciliacion(df_aux, df_ext):
    df_aux['Monto_Abs'] = df_aux['Monto_Neto'].abs().round(2)
    df_ext['Monto_Abs'] = df_ext['Monto_Neto'].abs().round(2)

    df_aux['Ocurrencia'] = df_aux.groupby(['Monto_Abs']).cumcount() + 1
    df_ext['Ocurrencia'] = df_ext.groupby(['Monto_Abs']).cumcount() + 1

    df_aux['LLAVE'] = df_aux['Monto_Abs'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)
    df_ext['LLAVE'] = df_ext['Monto_Abs'].astype(str) + "_" + df_ext['Ocurrencia'].astype(str)

    # Identificación de coincidencias exactas
    llaves_comunes = set(df_aux['LLAVE']).intersection(set(df_ext['LLAVE']))
    
    df_aux['Conciliado'] = df_aux['LLAVE'].isin(llaves_comunes)
    df_ext['Conciliado'] = df_ext['LLAVE'].isin(llaves_comunes)

    return df_aux, df_ext

# ==========================================
# 4. INTERFAZ Y NAVEGACIÓN
# ==========================================
st.sidebar.title("📌 Menú Principal")
opcion = st.sidebar.radio(
    "Selecciona el módulo:",
    ["📖 1. Cruce Diario (CSV vs Auxiliar)", "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ Configuración de Cruce Inteligente")
tol_pesos = st.sidebar.number_input("Tolerancia máxima en Pesos ($):", min_value=0.0, max_value=500.0, value=10.0, step=0.50)

if opcion == "📖 1. Cruce Diario (CSV vs Auxiliar)":
    st.title("📖 Cruce Operativo Diario")

    col1, col2 = st.columns(2)
    with col1:
        file_diario = st.file_uploader("1. Cargar Movimiento Diario (.csv)", type=["csv"], key="diario_csv")
    with col2:
        file_auxiliar = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx"], key="aux_diario")

    if file_diario and file_auxiliar:
        try:
            df_diario_raw = pd.read_csv(file_diario, encoding="latin1", header=None)
            df_aux_raw = pd.read_excel(file_auxiliar, sheet_name=0, header=None)

            df_diario = procesar_diario_csv(df_diario_raw)
            df_aux = procesar_auxiliar(df_aux_raw)

            df_aux, df_diario = ejecutar_conciliacion(df_aux, df_diario)

            solo_aux = df_aux[~df_aux['Conciliado']]
            solo_diario = df_diario[~df_diario['Conciliado']]

            sug_gen = obtener_sugerencias_cruce(solo_aux, solo_diario, tol_monto=tol_pesos)

            st.success(f"✅ Análisis completado. Se cruzaron {df_aux['Conciliado'].sum()} registros y se hallaron {len(sug_gen)} sugerencias.")

            st.markdown("### 📥 Generar y Descargar Archivos")
            c1, c2, c3 = st.columns(3)

            with c1:
                buf_egr = io.BytesIO()
                generar_excel_plantilla(df_aux[df_aux['Monto_Neto'] < 0], df_diario[df_diario['Monto_Neto'] < 0], sug_gen, buf_egr, "Egresos")
                st.download_button("🔻 Descargar EGRESOS (-)", data=buf_egr.getvalue(), file_name="Egresos_Diarios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with c2:
                buf_ing = io.BytesIO()
                generar_excel_plantilla(df_aux[df_aux['Monto_Neto'] > 0], df_diario[df_diario['Monto_Neto'] > 0], sug_gen, buf_ing, "Ingresos")
                st.download_button("🟢 Descargar INGRESOS (+)", data=buf_ing.getvalue(), file_name="Ingresos_Diarios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with c3:
                buf_gen = io.BytesIO()
                generar_excel_plantilla(df_aux, df_diario, sug_gen, buf_gen, "General_Consolidado")
                st.download_button("📦 Descargar GENERAL", data=buf_gen.getvalue(), file_name="General_Diarios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Error procesando los archivos: {str(e)}")

elif opcion == "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)":
    st.title("📊 Conciliación Bancaria Mensual")

    col1, col2 = st.columns(2)
    with col1:
        file_ext = st.file_uploader("1. Cargar Extracto Bancario (.xlsx)", type=["xlsx"], key="ext_mensual")
    with col2:
        file_auxiliar = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx"], key="aux_mensual")

    if file_ext and file_auxiliar:
        try:
            df_ext_raw = pd.read_excel(file_ext, sheet_name=0, header=None)
            df_aux_raw = pd.read_excel(file_auxiliar, sheet_name=0, header=None)

            df_ext = procesar_extracto_excel(df_ext_raw)
            df_aux = procesar_auxiliar(df_aux_raw)

            # Ejecutar cruce de información
            df_aux, df_ext = ejecutar_conciliacion(df_aux, df_ext)

            solo_ext = df_ext[~df_ext['Conciliado']]
            solo_aux = df_aux[~df_aux['Conciliado']]

            sug_gen_m = obtener_sugerencias_cruce(solo_aux, solo_ext, tol_monto=tol_pesos)

            st.success(f"✅ Conciliación realizada. Registros cruzados: {df_aux['Conciliado'].sum()}. Sugerencias por diferencia: {len(sug_gen_m)}.")

            st.markdown("### 📥 Generar y Descargar Archivos")
            c1, c2, c3 = st.columns(3)

            with c1:
                buf_egr_m = io.BytesIO()
                generar_excel_plantilla(df_aux[df_aux['Monto_Neto'] < 0], df_ext[df_ext['Monto_Neto'] < 0], sug_gen_m, buf_egr_m, "Egresos")
                st.download_button("🔻 Descargar EGRESOS (-)", data=buf_egr_m.getvalue(), file_name="Conciliacion_Egresos_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with c2:
                buf_ing_m = io.BytesIO()
                generar_excel_plantilla(df_aux[df_aux['Monto_Neto'] > 0], df_ext[df_ext['Monto_Neto'] > 0], sug_gen_m, buf_ing_m, "Ingresos")
                st.download_button("🟢 Descargar INGRESOS (+)", data=buf_ing_m.getvalue(), file_name="Conciliacion_Ingresos_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with c3:
                buf_gen_m = io.BytesIO()
                generar_excel_plantilla(df_aux, df_ext, sug_gen_m, buf_gen_m, "General_Consolidado")
                st.download_button("📦 Descargar GENERAL", data=buf_gen_m.getvalue(), file_name="Conciliacion_General_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            if not sug_gen_m.empty:
                st.markdown("---")
                st.subheader("💡 Vista Previa: Sugerencias de Cruce Halladas")
                st.dataframe(sug_gen_m)

        except Exception as e:
            st.error(f"Error al ejecutar la conciliación: {str(e)}")
