import io
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import streamlit as st

# Configuración de la página Streamlit
st.set_page_config(page_title="Sistema de Conciliación Bancaria", layout="wide")

# ==========================================
# 1. FUNCIONES AUXILIARES DE LIMPIEZA
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

def formatear_fecha_original(val):
    """
    Conserva la fecha exacta del archivo original.
    Transforma formatos nativos o texto a DD/MM/YYYY sin alterar el valor real.
    """
    if pd.isna(val) or str(val).strip() == '':
        return ""
    
    val_str = str(val).strip().split(' ')[0].replace('-', '/')
    
    # Manejo de timestamps/objetos date de pandas
    try:
        dt = pd.to_datetime(val_str, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.strftime('%d/%m/%Y')
    except:
        pass
        
    return val_str

# ==========================================
# 2. ALGORITMO DE SUGERENCIAS INTELIGENTES
# ==========================================
def obtener_sugerencias_cruce(df_aux_pend, df_ext_pend, tol_monto=10.0):
    sugerencias = []
    if df_aux_pend.empty or df_ext_pend.empty:
        return pd.DataFrame(sugerencias)

    aux_tmp = df_aux_pend.copy()
    ext_tmp = df_ext_pend.copy()

    aux_tmp['Fecha_dt'] = pd.to_datetime(aux_tmp['Fecha_Clean'], format='%d/%m/%Y', errors='coerce')
    ext_tmp['Fecha_dt'] = pd.to_datetime(ext_tmp['Fecha_Clean'], format='%d/%m/%Y', errors='coerce')

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
                
                candidatos.append({
                    'idx_e': idx_e,
                    'Fecha_Banco': row_e['Fecha_Clean'],
                    'Ref_Banco': row_e.get('Descripcion', row_e.get('DESCRIPCIÓN', '')),
                    'Monto_Banco': monto_e,
                    'Diff_Monto': diff_monto,
                    'Dias_Desfase': dias_desfase
                })

        if candidatos:
            candidatos.sort(key=lambda x: (x['Diff_Monto'], x['Dias_Desfase']))
            mejor = candidatos[0]
            ext_usados.add(mejor['idx_e'])
            
            certeza = "Alta (Exacto)" if mejor['Diff_Monto'] == 0 else f"Media (Diff: ${mejor['Diff_Monto']:.2f})"
            
            sugerencias.append({
                'Fecha Contable': row_a['Fecha_Clean'],
                'Documento': row_a.get('Comprobante', ''),
                'Tercero': row_a.get('Nombre del tercero', ''),
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
# 3. PARSERS INDEPENDIENTES Y SEGUROS
# ==========================================
def procesar_auxiliar(df_raw):
    """Procesa el Auxiliar Contable estándar."""
    df = df_raw.copy()
    if len(df) > 7 and 'Código contable' in str(df.iloc[6].values):
        headers = df.iloc[6].values
        df = df.iloc[8:].copy()
        df.columns = headers

    df = df.dropna(how='all')
    
    col_deb = next((c for c in df.columns if 'Débito' in str(c) or 'Debito' in str(c)), None)
    col_cred = next((c for c in df.columns if 'Crédito' in str(c) or 'Credito' in str(c)), None)
    col_fec = next((c for c in df.columns if 'Fecha' in str(c)), None)

    df['Débito_Clean'] = df[col_deb].apply(limpiar_valor) if col_deb else 0.0
    df['Crédito_Clean'] = df[col_cred].apply(limpiar_valor) if col_cred else 0.0
    df['Monto_Neto'] = df['Débito_Clean'] - df['Crédito_Clean']

    df['Fecha_Clean'] = df[col_fec].apply(formatear_fecha_original) if col_fec else ""
    return df

def procesar_diario_csv(df_raw):
    """Procesa el CSV de movimiento diario garantizando la lectura de fecha original."""
    df = df_raw.copy()
    
    # Mapeo por posición fija de columnas del CSV diario
    if df.shape[1] >= 8:
        df_proc = df.iloc[:, [2, 3, 5, 7]].copy()
        df_proc.columns = ['Fecha_Int', 'Valor_Raw', 'Codigo_Tx', 'Descripcion']
    else:
        df_proc = df.copy()
        df_proc.columns = [f'Col_{i}' for i in range(df_proc.shape[1])]
        df_proc['Fecha_Int'] = df_proc.iloc[:, 0]
        df_proc['Valor_Raw'] = df_proc.iloc[:, 1] if df_proc.shape[1] > 1 else 0
        df_proc['Descripcion'] = df_proc.iloc[:, -1]

    df_proc['Monto_Neto'] = df_proc['Valor_Raw'].apply(limpiar_valor)
    df_proc['Fecha_Clean'] = df_proc['Fecha_Int'].apply(formatear_fecha_original)
    return df_proc

def procesar_extracto_mensual_excel(df_raw):
    """Procesa el Excel mensual de extracto omitiendo encabezados de manera segura."""
    df = df_raw.copy()
    
    # Filtrar solo filas donde la primera columna parece una fecha válida (DD/MM o YYYY)
    mask = df.iloc[:, 0].astype(str).str.contains(r'\d{1,2}[/-]\d{1,2}', regex=True, na=False)
    df_clean = df[mask].copy()
    
    if df_clean.empty:
        df_clean = df.dropna(how='all').iloc[8:].copy() # Fallback por índice si no coincide patrón

    df_clean = df_clean.iloc[:, :6]
    cols_base = ['FECHA', 'DESCRIPCIÓN', 'SUCURSAL', 'DCTO', 'VALOR', 'SALDO']
    df_clean.columns = cols_base[:df_clean.shape[1]]

    col_val = 'VALOR' if 'VALOR' in df_clean.columns else df_clean.columns[-2]
    col_fec = 'FECHA' if 'FECHA' in df_clean.columns else df_clean.columns[0]

    df_clean['Monto_Neto'] = df_clean[col_val].apply(limpiar_valor)
    df_clean['Fecha_Clean'] = df_clean[col_fec].apply(formatear_fecha_original)
    return df_clean

# ==========================================
# 4. GENERADOR DE LIBRO EXCEL (6 HOJAS)
# ==========================================
def generar_excel_plantilla(df_aux_proc, df_ext_proc, df_sugerencias, buffer, tipo_reporte="General"):
    wb = openpyxl.Workbook()

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

    # HOJA 1: CONCILIACION
    ws_resumen = wb.active
    ws_resumen.title = "CONCILIACION"
    ws_resumen.views.sheetView[0].showGridLines = True

    ws_resumen['A2'] = f"CONCILIACIÓN BANCARIA AUTOMATIZADA - REPORTE {tipo_reporte.upper()}"
    ws_resumen['A2'].font = font_title
    ws_resumen['B2'] = tipo_reporte
    ws_resumen['B2'].font = font_bold

    headers_res = ["CONCEPTO", "VALOR (COP / USD)", "NOTAS / AUDITORÍA"]
    for col_idx, text in enumerate(headers_res, 1):
        cell = ws_resumen.cell(row=8, column=col_idx, value=text)
        cell.font, cell.fill, cell.alignment = font_header, fill_header, Alignment(horizontal="center", vertical="center")

    max_aux = len(df_aux_proc) + 1
    max_ext = len(df_ext_proc) + 1

    filas = [
        ("Saldo Mov según Auxiliar Contable", f"=SUM('AUXILIAR'!E2:E{max(max_aux, 2)})", None),
        ("(+) Partidas No registradas por el Banco", f'=SUMIF(\'AUXILIAR\'!G2:G{max(max_aux, 2)}, "NO ESTA EN BANCOS", \'AUXILIAR\'!E2:E{max(max_aux, 2)})', "Movimientos en libros pendientes en extracto"),
        ("SALDO CONTABLE AJUSTADO", "=B9-B10", "Saldo conciliado contable"),
        ("Saldo Final según Extracto Bancario", f"=SUM('EXTRACTO'!C2:C{max(max_ext, 2)})", None),
        ("(-) Partidas no Contabilizadas", f'=SUMIF(\'EXTRACTO\'!E2:E{max(max_ext, 2)}, "Pen Contabilidad", \'EXTRACTO\'!C2:C{max(max_ext, 2)})', "Movimientos del banco faltantes en contabilidad"),
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

    # HOJA 2: AUXILIAR
    ws_aux = wb.create_sheet(title="AUXILIAR")
    ws_aux.views.sheetView[0].showGridLines = True
    ws_aux.append(["Fecha", "Documento", "Concepto/Detalle", "NOMBRE", "Monto (+/-)", "LLAVE", "Extracto"])

    for idx, row in df_aux_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        ws_aux.append([
            str(row.get('Fecha_Clean', '')),
            str(row.get('Comprobante', '')),
            str(row.get('Concepto', row.get('Código contable', ''))),
            str(row.get('Nombre del tercero', '')),
            float(row.get('Monto_Neto', 0)),
            f'=CONCATENATE(A{r},E{r},"-",COUNTIFS($A$2:A{r},A{r},$E$2:E{r},E{r}))',
            f'=IF(ISNUMBER(MATCH(F{r}, \'EXTRACTO\'!$D:$D, 0)), "CONCILIADO", "NO ESTA EN BANCOS")'
        ])

    max_aux_tbl = max(ws_aux.max_row, 2)
    tab1 = Table(displayName="TablaAuxiliar", ref=f"A1:G{max_aux_tbl}")
    tab1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_aux.add_table(tab1)

    # HOJA 3: EXTRACTO
    ws_ext = wb.create_sheet(title="EXTRACTO")
    ws_ext.views.sheetView[0].showGridLines = True
    ws_ext.append(["Fecha", "Referencia", "Monto (+/-)", "LLAVE", "Contabilidad"])

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        ws_ext.append([
            str(row.get('Fecha_Clean', '')),
            str(row.get('Descripcion', row.get('DESCRIPCIÓN', ''))),
            float(row.get('Monto_Neto', 0)),
            f'=CONCATENATE(A{r},C{r},"-",COUNTIFS($A$2:A{r},A{r},$C$2:C{r},C{r}))',
            f'=IF(ISNUMBER(MATCH(D{r}, \'AUXILIAR\'!$F:$F, 0)), "CONCILIADO", "Pen Contabilidad")'
        ])

    max_ext_tbl = max(ws_ext.max_row, 2)
    tab2 = Table(displayName="TablaExtracto", ref=f"A1:E{max_ext_tbl}")
    tab2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_ext.add_table(tab2)

    # HOJA 4: SUGERENCIAS
    ws_sug = wb.create_sheet(title="SUGERENCIAS")
    ws_sug.views.sheetView[0].showGridLines = True
    ws_sug.append(["Fecha Contable", "Documento", "Tercero", "Monto Contable", "Fecha Banco", "Referencia Banco", "Monto Banco", "Diferencia ($)", "Días Desfase", "Certeza"])

    if not df_sugerencias.empty:
        for _, r_sug in df_sugerencias.iterrows():
            ws_sug.append([
                str(r_sug.get('Fecha Contable', '')),
                str(r_sug.get('Documento', '')),
                str(r_sug.get('Tercero', '')),
                float(r_sug.get('Monto Contable', 0)),
                str(r_sug.get('Fecha Banco', '')),
                str(r_sug.get('Referencia Banco', '')),
                float(r_sug.get('Monto Banco', 0)),
                float(r_sug.get('Diferencia ($)', 0)),
                int(r_sug.get('Días Desfase', 0)),
                str(r_sug.get('Certeza', ''))
            ])
    
    max_sug_tbl = max(ws_sug.max_row, 2)
    tab3 = Table(displayName="TablaSugerencias", ref=f"A1:J{max_sug_tbl}")
    tab3.tableStyleInfo = TableStyleInfo(name="TableStyleMedium3", showRowStripes=True)
    ws_sug.add_table(tab3)

    # HOJA 5: PENDIENTES POR REGISTRAR
    ws_pend = wb.create_sheet(title="PENDIENTES POR REGISTRAR")
    ws_pend.views.sheetView[0].showGridLines = True
    ws_pend.append(["Fecha Banco", "Referencia / Descripción", "Monto (+/-)", "Estado"])

    refs_sugeridas = set(df_sugerencias['Referencia Banco'].dropna().unique()) if (not df_sugerencias.empty and 'Referencia Banco' in df_sugerencias.columns) else set()

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        ref_val = str(row.get('Descripcion', row.get('DESCRIPCIÓN', '')))
        if ref_val not in refs_sugeridas:
            ws_pend.append([str(row.get('Fecha_Clean', '')), ref_val, float(row.get('Monto_Neto', 0)), "Pen Contabilidad"])

    max_pend_tbl = max(ws_pend.max_row, 2)
    tab5 = Table(displayName="TablaPendientes", ref=f"A1:D{max_pend_tbl}")
    tab5.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
    ws_pend.add_table(tab5)

    # HOJA 6: GASTOS BANCARIOS
    ws_gastos = wb.create_sheet(title="GASTOS BANCARIOS")
    ws_gastos.views.sheetView[0].showGridLines = True
    ws_gastos.append(["Fecha", "Descripción / Concepto Gasto", "Monto Gasto (-)", "Clasificación"])

    patron_gastos = r'IMPTO GOBIERNO|4X1000|COMISION|IVA COMISION|INTERESES|CUOTA MANEJO|MANTE SUCURSAL'
    total_gastos = 0.0

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        ref_val = str(row.get('Descripcion', row.get('DESCRIPCIÓN', '')))
        if re.search(patron_gastos, ref_val, re.IGNORECASE):
            monto_val = float(row.get('Monto_Neto', 0))
            ws_gastos.append([str(row.get('Fecha_Clean', '')), ref_val, monto_val, "Gasto / Impuesto Financiero"])
            total_gastos += monto_val

    r_total = ws_gastos.max_row + 1
    ws_gastos.cell(row=r_total, column=1, value="---").font = font_bold
    c_tot_lbl = ws_gastos.cell(row=r_total, column=2, value="TOTAL GASTOS BANCARIOS")
    c_tot_lbl.font, c_tot_lbl.fill = font_bold, fill_subtotal
    
    c_tot_val = ws_gastos.cell(row=r_total, column=3, value=total_gastos)
    c_tot_val.font, c_tot_val.fill = font_bold, fill_subtotal
    c_tot_val.number_format = '$#,##0.00'
    
    ws_gastos.cell(row=r_total, column=4, value="SUBTOTAL CONSOLIDADO").font = font_bold

    max_gastos_tbl = max(r_total - 1, 2)
    tab6 = Table(displayName="TablaGastos", ref=f"A1:D{max_gastos_tbl}")
    tab6.tableStyleInfo = TableStyleInfo(name="TableStyleMedium7", showRowStripes=True)
    ws_gastos.add_table(tab6)

    wb.save(buffer)

# ==========================================
# 5. INTERFAZ STREAMLIT
# ==========================================
st.sidebar.title("📌 Menú Principal")
opcion = st.sidebar.radio(
    "Selecciona el módulo:",
    ["📖 1. Cruce Diario (CSV vs Auxiliar)", "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ Ajustes del Algoritmo")
tol_pesos = st.sidebar.number_input("Tolerancia en Pesos ($):", min_value=0.0, max_value=500.0, value=10.0, step=1.0)

# ----------------------------------------------------
# MÓDULO 1: DIARIO
# ----------------------------------------------------
if opcion == "📖 1. Cruce Diario (CSV vs Auxiliar)":
    st.title("📖 Cruce Operativo Diario")

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

        sug_gen = obtener_sugerencias_cruce(df_aux, df_diario, tol_monto=tol_pesos)

        st.success(f"✅ Análisis completado. Se hallaron {len(sug_gen)} posibles coincidencias.")

        st.markdown("### 📥 Generar y Descargar Archivos (6 Hojas)")
        c1, c2, c3 = st.columns(3)

        with c1:
            buf_egr = io.BytesIO()
            df_a_egr = df_aux[df_aux['Monto_Neto'] < 0]
            df_d_egr = df_diario[df_diario['Monto_Neto'] < 0]
            sug_egr = obtener_sugerencias_cruce(df_a_egr, df_d_egr, tol_monto=tol_pesos)
            generar_excel_plantilla(df_a_egr, df_d_egr, sug_egr, buf_egr, "Egresos")
            st.download_button("🔻 Descargar EGRESOS (-)", data=buf_egr.getvalue(), file_name="Egresos_Diarios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c2:
            buf_ing = io.BytesIO()
            df_a_ing = df_aux[df_aux['Monto_Neto'] > 0]
            df_d_ing = df_diario[df_diario['Monto_Neto'] > 0]
            sug_ing = obtener_sugerencias_cruce(df_a_ing, df_d_ing, tol_monto=tol_pesos)
            generar_excel_plantilla(df_a_ing, df_d_ing, sug_ing, buf_ing, "Ingresos")
            st.download_button("🟢 Descargar INGRESOS (+)", data=buf_ing.getvalue(), file_name="Ingresos_Diarios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c3:
            buf_gen = io.BytesIO()
            generar_excel_plantilla(df_aux, df_diario, sug_gen, buf_gen, "General_Consolidado")
            st.download_button("📦 Descargar GENERAL", data=buf_gen.getvalue(), file_name="General_Diarios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if not sug_gen.empty:
            st.markdown("---")
            st.subheader("💡 Vista Previa: Posibles Sugerencias de Cruce Halladas")
            st.dataframe(sug_gen)

# ----------------------------------------------------
# MÓDULO 2: MENSUAL
# ----------------------------------------------------
elif opcion == "📊 2. Conciliación Bancaria Mensual (Excel vs Auxiliar)":
    st.title("📊 Conciliación Bancaria Mensual")

    col1, col2 = st.columns(2)
    with col1:
        file_ext = st.file_uploader("1. Cargar Extracto Bancario (.xlsx)", type=["xlsx"], key="ext_mensual")
    with col2:
        file_auxiliar = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx"], key="aux_mensual")

    if file_ext and file_auxiliar:
        df_ext_raw = pd.read_excel(file_ext, header=None)
        df_aux_raw = pd.read_excel(file_auxiliar)

        df_ext = procesar_extracto_mensual_excel(df_ext_raw)
        df_aux = procesar_auxiliar(df_aux_raw)

        sug_gen_m = obtener_sugerencias_cruce(df_aux, df_ext, tol_monto=tol_pesos)

        st.success(f"✅ Conciliación realizada. Se detectaron {len(sug_gen_m)} sugerencias de cruce.")

        st.markdown("### 📥 Generar y Descargar Archivos (6 Hojas)")
        c1, c2, c3 = st.columns(3)

        with c1:
            buf_egr_m = io.BytesIO()
            df_a_egr = df_aux[df_aux['Monto_Neto'] < 0]
            df_e_egr = df_ext[df_ext['Monto_Neto'] < 0]
            sug_egr_m = obtener_sugerencias_cruce(df_a_egr, df_e_egr, tol_monto=tol_pesos)
            generar_excel_plantilla(df_a_egr, df_e_egr, sug_egr_m, buf_egr_m, "Egresos")
            st.download_button("🔻 Descargar EGRESOS (-)", data=buf_egr_m.getvalue(), file_name="Conciliacion_Egresos_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c2:
            buf_ing_m = io.BytesIO()
            df_a_ing = df_aux[df_aux['Monto_Neto'] > 0]
            df_e_ing = df_ext[df_ext['Monto_Neto'] > 0]
            sug_ing_m = obtener_sugerencias_cruce(df_a_ing, df_e_ing, tol_monto=tol_pesos)
            generar_excel_plantilla(df_a_ing, df_e_ing, sug_ing_m, buf_ing_m, "Ingresos")
            st.download_button("🟢 Descargar INGRESOS (+)", data=buf_ing_m.getvalue(), file_name="Conciliacion_Ingresos_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c3:
            buf_gen_m = io.BytesIO()
            generar_excel_plantilla(df_aux, df_ext, sug_gen_m, buf_gen_m, "General_Consolidado")
            st.download_button("📦 Descargar GENERAL", data=buf_gen_m.getvalue(), file_name="Conciliacion_General_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if not sug_gen_m.empty:
            st.markdown("---")
            st.subheader("💡 Vista Previa: Posibles Sugerencias de Cruce Halladas")
            st.dataframe(sug_gen_m)
