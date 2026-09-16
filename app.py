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
    
    # Copia con fechas procesadas en formato datetime para cálculos
    aux_tmp = df_aux_pend.copy()
    ext_tmp = df_ext_pend.copy()
    
    aux_tmp['Fecha_dt'] = pd.to_datetime(aux_tmp['Fecha_Clean'], errors='coerce')
    ext_tmp['Fecha_dt'] = pd.to_datetime(ext_tmp['Fecha_Clean'], errors='coerce')

    ext_usados = set()

    for idx_a, row_a in aux_tmp.iterrows():
        monto_a = float(row_a['Monto_Neto'])
        fecha_a = row_a['Fecha_dt']
        
        if pd.isna(fecha_a):
            continue

        candidatos = []
        for idx_e, row_e in ext_tmp.iterrows():
            if idx_e in ext_usados:
                continue
                
            monto_e = float(row_e['Monto_Neto'])
            fecha_e = row_e['Fecha_dt']
            
            if pd.isna(fecha_e):
                continue

            # Regla 1: Mismo Mes y Mismo Año
            if (fecha_a.year == fecha_e.year) and (fecha_a.month == fecha_e.month):
                diff_monto = abs(monto_a - monto_e)
                
                # Regla 2: Monto dentro de la tolerancia de centavos/pesos
                if diff_monto <= tol_monto:
                    dias_desfase = abs((fecha_e - fecha_a).days)
                    candidatos.append({
                        'idx_e': idx_e,
                        'Fecha_Banco': row_e['Fecha_Clean'],
                        'Ref_Banco': row_e.get('Descripcion', row_e.get('DESCRIPCIÓN', '')),
                        'Monto_Banco': monto_e,
                        'Diff_Monto': diff_monto,
                        'Dias_Desfase': dias_desfase
                    })

        if candidatos:
            # Ordenar candidatos: primero menor diferencia en monto, luego menor desfase en días
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
        fecha_val = row.get('Fecha_Clean', row.get('Fecha elaboración', ''))
        doc_val = row.get('Comprobante', '')
        concepto_val = row.get('Concepto', row.get('Código contable', ''))
        nombre_val = row.get('Nombre del tercero', '')
        monto_val = float(row.get('Monto_Neto', 0))

        ws_aux.append([
            str(fecha_val),
            str(doc_val),
            str(concepto_val),
            str(nombre_val),
            monto_val,
            f'=CONCATENATE(A{r},E{r},"-",COUNTIFS($A$2:A{r},A{r},$E$2:E{r},E{r}))',
            f'=IF(ISNUMBER(MATCH(F{r}, \'Extracto Bancario\'!$D:$D, 0)), "CONCILIADO", "NO ESTA EN BANCOS")'
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
        fecha_val = row.get('Fecha_Clean', row.get('FECHA', ''))
        ref_val = row.get('Descripcion', row.get('DESCRIPCIÓN', ''))
        monto_val = float(row.get('Monto_Neto', 0))

        ws_ext.append([
            str(fecha_val),
            str(ref_val),
            monto_val,
            f'=CONCATENATE(A{r},C{r},"-",COUNTIFS($A$2:A{r},A{r},$C$2:C{r},C{r}))',
            f'=IF(ISNUMBER(MATCH(D{r}, \'Auxiliar Contable\'!$F:$F, 0)), "CONCILIADO", "Pen Contabilidad")'
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
    headers = df_raw.iloc[6].values
    df = df_raw.iloc[8:].copy()
    df.columns = headers
    df = df.dropna(subset=['Código contable', 'Fecha elaboración'])
    
    df['Débito_Clean'] = df['Débito'].apply(limpiar_valor)
    df['Crédito_Clean'] = df['Crédito'].apply(limpiar_valor)
    df['Monto_Neto'] = df['Débito_Clean'] - df['Crédito_Clean']
    
    col_fecha = 'Fecha elaboration' if 'Fecha elaboration' in df.columns else 'Fecha elaboración'
    df['Fecha_Clean'] = pd.to_datetime(df[col_fecha], format='%d/%m/%Y', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_diario_csv(df_raw):
    df = df_raw.iloc[:, [0, 1, 3, 5, 6, 7]].copy()
    df.columns = ['Cuenta', 'Oficina', 'Fecha_Int', 'Valor_Raw', 'Codigo_Tx', 'Descripcion']
    df['Monto_Neto'] = df['Valor_Raw'].apply(limpiar_valor)
    df['Fecha_Clean'] = pd.to_datetime(df['Fecha_Int'].astype(str), format='%Y%m%d', errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def procesar_extracto_excel(df_raw):
    df = df_raw.iloc[15:].copy()
    df.columns = ['FECHA', 'DESCRIPCIÓN', 'SUCURSAL', 'DCTO', 'VALOR', 'SALDO', 'X1', 'X2']
    df = df.dropna(subset=['FECHA', 'VALOR'])
    df = df[~df['FECHA'].astype(str).str.contains('FECHA|FIN ESTADO', case=False, na=False)]
    df['Monto_Neto'] = df['VALOR'].apply(limpiar_valor)
    return df

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
        df_diario_raw = pd.read_csv(file_diario, encoding="latin1", header=None)
        df_aux_raw = pd.read_excel(file_auxiliar)

        df_diario = procesar_diario_csv(df_diario_raw)
        df_aux = procesar_auxiliar(df_aux_raw)

        # Identificar pendientes exactos
        df_diario['Monto_Abs'] = df_diario['Monto_Neto'].abs().round(2)
        df_aux['Monto_Abs'] = df_aux['Monto_Neto'].abs().round(2)
        
        df_diario['Ocurrencia'] = df_diario.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1
        df_aux['Ocurrencia'] = df_aux.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1

        df_diario['LLAVE'] = df_diario['Fecha_Clean'] + "_" + df_diario['Monto_Abs'].astype(str) + "_" + df_diario['Ocurrencia'].astype(str)
        df_aux['LLAVE'] = df_aux['Fecha_Clean'] + "_" + df_aux['Monto_Abs'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)

        solo_diario = df_diario[~df_diario['LLAVE'].isin(df_aux['LLAVE'])]
        solo_aux = df_aux[~df_aux['LLAVE'].isin(df_diario['LLAVE'])]

        # Algoritmo de Sugerencias
        sug_gen = obtener_sugerencias_cruce(solo_aux, solo_diario, tol_monto=tol_pesos)

        st.success(f"✅ Análisis completado. Se hallaron {len(sug_gen)} posibles coincidencias por desfase de fecha/redondeo.")

        st.markdown("### 📥 Generar y Descargar Archivos")
        c1, c2, c3 = st.columns(3)

        with c1:
            buf_egr = io.BytesIO()
            df_a_egr = df_aux[df_aux['Monto_Neto'] < 0]
            df_d_egr = df_diario[df_diario['Monto_Neto'] < 0]
            sug_egr = obtener_sugerencias_cruce(solo_aux[solo_aux['Monto_Neto'] < 0], solo_diario[solo_diario['Monto_Neto'] < 0], tol_monto=tol_pesos)
            generar_excel_plantilla(df_a_egr, df_d_egr, sug_egr, buf_egr, "Egresos")
            st.download_button("🔻 Descargar EGRESOS (-)", data=buf_egr.getvalue(), file_name="Egresos_Diarios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c2:
            buf_ing = io.BytesIO()
            df_a_ing = df_aux[df_aux['Monto_Neto'] > 0]
            df_d_ing = df_diario[df_diario['Monto_Neto'] > 0]
            sug_ing = obtener_sugerencias_cruce(solo_aux[solo_aux['Monto_Neto'] > 0], solo_diario[solo_diario['Monto_Neto'] > 0], tol_monto=tol_pesos)
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

        df_ext = procesar_extracto_excel(df_ext_raw)
        df_aux = procesar_auxiliar(df_aux_raw)

        # Identificar pendientes exactos
        df_ext['Monto_Abs'] = df_ext['Monto_Neto'].abs().round(2)
        df_aux['Monto_Abs'] = df_aux['Monto_Neto'].abs().round(2)

        df_ext['Ocurrencia'] = df_ext.groupby(['Monto_Abs']).cumcount() + 1
        df_aux['Ocurrencia'] = df_aux.groupby(['Monto_Abs']).cumcount() + 1

        df_ext['LLAVE'] = df_ext['Monto_Abs'].astype(str) + "_" + df_ext['Ocurrencia'].astype(str)
        df_aux['LLAVE'] = df_aux['Monto_Abs'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)

        solo_ext = df_ext[~df_ext['LLAVE'].isin(df_aux['LLAVE'])]
        solo_aux = df_aux[~df_aux['LLAVE'].isin(df_ext['LLAVE'])]

        # Algoritmo de Sugerencias
        sug_gen_m = obtener_sugerencias_cruce(solo_aux, solo_ext, tol_monto=tol_pesos)

        st.success(f"✅ Conciliación realizada. Se detectaron {len(sug_gen_m)} sugerencias de cruce en el mes.")

        st.markdown("### 📥 Generar y Descargar Archivos")
        c1, c2, c3 = st.columns(3)

        with c1:
            buf_egr_m = io.BytesIO()
            df_a_egr = df_aux[df_aux['Monto_Neto'] < 0]
            df_e_egr = df_ext[df_ext['Monto_Neto'] < 0]
            sug_egr_m = obtener_sugerencias_cruce(solo_aux[solo_aux['Monto_Neto'] < 0], solo_ext[solo_ext['Monto_Neto'] < 0], tol_monto=tol_pesos)
            generar_excel_plantilla(df_a_egr, df_e_egr, sug_egr_m, buf_egr_m, "Egresos")
            st.download_button("🔻 Descargar EGRESOS (-)", data=buf_egr_m.getvalue(), file_name="Conciliacion_Egresos_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c2:
            buf_ing_m = io.BytesIO()
            df_a_ing = df_aux[df_aux['Monto_Neto'] > 0]
            df_e_ing = df_ext[df_ext['Monto_Neto'] > 0]
            sug_ing_m = obtener_sugerencias_cruce(solo_aux[solo_aux['Monto_Neto'] > 0], solo_ext[solo_ext['Monto_Neto'] > 0], tol_monto=tol_pesos)
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
