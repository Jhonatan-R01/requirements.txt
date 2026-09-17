import io
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sistema de Conciliación Bancaria", layout="wide")

# ==========================================
# 1. FUNCIONES AUXILIARES DE LIMPIEZA
# ==========================================
def limpiar_valor(val):
    if pd.isna(val) or str(val).strip() in ['', '-']:
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

# ==========================================
# 2. PARSERS BASE Y TRANSFORMACIÓN DE SIGNOS
# ==========================================
def procesar_auxiliar(df_raw):
    headers = df_raw.iloc[6].values
    df = df_raw.iloc[8:].copy()
    df.columns = [str(c).strip() for c in headers]
    df = df.dropna(subset=['Código contable', 'Fecha elaboración'])
    
    df['Débito_Clean'] = df['Débito'].apply(limpiar_valor)
    df['Crédito_Clean'] = df['Crédito'].apply(limpiar_valor)
    
    # Conversión garantizada: Crédito (positivo en origen) pasa a Egreso (-)
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
    df = df_raw.iloc[:, 0:6].copy()
    df.columns = ['FECHA', 'DESCRIPCIÓN', 'SUCURSAL', 'DCTO', 'VALOR', 'SALDO']
    
    df['FECHA_STR'] = df['FECHA'].astype(str).str.strip()
    patron_fecha = r'^\d{1,2}/\d{1,2}'
    df_clean = df[df['FECHA_STR'].str.contains(patron_fecha, na=False)].copy()
    
    df_clean['Monto_Neto'] = df_clean['VALOR'].apply(limpiar_valor)
    df_clean['Fecha_Clean'] = pd.to_datetime(df_clean['FECHA_STR'] + '/2026', format='%d/%m/%Y', errors='coerce').dt.strftime('%Y-%m-%d')
    
    return df_clean.drop(columns=['FECHA_STR'])

# ==========================================
# 3. ALGORITMO DE SUGERENCIAS CON CONTROL DE SIGNO
# ==========================================
def obtener_sugerencias_cruce(df_aux_pend, df_ext_pend, tol_monto=10.0):
    sugerencias = []
    if df_aux_pend.empty or df_ext_pend.empty:
        return pd.DataFrame(sugerencias)

    aux_tmp = df_aux_pend.copy()
    ext_tmp = df_ext_pend.copy()
    
    aux_tmp['Fecha_dt'] = pd.to_datetime(aux_tmp['Fecha_Clean'], errors='coerce')
    ext_tmp['Fecha_dt'] = pd.to_datetime(ext_tmp['Fecha_Clean'], errors='coerce')

    ext_usados = set()

    for idx_a, row_a in aux_tmp.iterrows():
        monto_a = float(row_a['Monto_Neto'])
        fecha_a = row_a['Fecha_dt']
        txt_a = str(row_a.get('Nombre del tercero', '')).upper().strip()
        
        if pd.isna(fecha_a):
            continue

        candidatos = []
        for idx_e, row_e in ext_tmp.iterrows():
            if idx_e in ext_usados:
                continue
                
            monto_e = float(row_e['Monto_Neto'])
            fecha_e = row_e['Fecha_dt']
            txt_e = str(row_e.get('Descripcion', row_e.get('DESCRIPCIÓN', ''))).upper().strip()
            
            if pd.isna(fecha_e):
                continue

            # CONTROL DE SIGNO ESTRICTO: Compara únicamente (+ con +) o (- con -)
            mismo_sentido = (monto_a > 0 and monto_e > 0) or (monto_a < 0 and monto_e < 0)

            if mismo_sentido and (fecha_a.year == fecha_e.year) and (fecha_a.month == fecha_e.month):
                diff_monto = abs(abs(monto_a) - abs(monto_e))
                dias_desfase = abs((fecha_e - fecha_a).days)
                
                palabras_a = [p for p in txt_a.split() if len(p) > 3]
                match_texto = any(p in txt_e for p in palabras_a) if palabras_a else False

                if diff_monto <= tol_monto:
                    certeza = None
                    if match_texto:
                        certeza = "Alta (Monto + Texto / Tercero)"
                    elif dias_desfase <= 3:
                        certeza = f"Media (Monto + Desfase {dias_desfase} días)"
                    
                    if certeza:
                        candidatos.append({
                            'idx_e': idx_e,
                            'Fecha_Banco': row_e['Fecha_Clean'],
                            'Ref_Banco': row_e.get('Descripcion', row_e.get('DESCRIPCIÓN', '')),
                            'Monto_Banco': monto_e,
                            'Diff_Monto': diff_monto,
                            'Dias_Desfase': dias_desfase,
                            'Certeza': certeza
                        })

        if candidatos:
            candidatos.sort(key=lambda x: (0 if "Alta" in x['Certeza'] else 1, x['Diff_Monto'], x['Dias_Desfase']))
            mejor = candidatos[0]
            ext_usados.add(mejor['idx_e'])
            
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
                'Certeza': mejor['Certeza']
            })

    return pd.DataFrame(sugerencias)

# ==========================================
# 4. GENERADOR EXCEL MULTI-HOJA CON DEPURACIÓN TRIPLE
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

    # HOJA 1: Resumen Conciliación
    ws_resumen = wb.active
    ws_resumen.title = "Resumen Conciliación"
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

    # HOJA 2: Auxiliar Contable
    ws_aux = wb.create_sheet(title="Auxiliar Contable")
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
            f'=CONCATENATE(A{r}, "-", E{r}, "-", COUNTIFS($A$2:A{r}, A{r}, $E$2:E{r}, E{r}))',
            f'=IF(ISNUMBER(MATCH(F{r}, \'Extracto Bancario\'!$D:$D, 0)), "CONCILIADO", "NO ESTA EN BANCOS")'
        ])

    max_aux_tbl = max(ws_aux.max_row, 2)
    tab1 = Table(displayName="TablaAuxiliar", ref=f"A1:G{max_aux_tbl}")
    tab1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_aux.add_table(tab1)

    # HOJA 3: Extracto Bancario
    ws_ext = wb.create_sheet(title="Extracto Bancario")
    ws_ext.views.sheetView[0].showGridLines = True
    ws_ext.append(["Fecha", "Referencia", "Monto (+/-)", "LLAVE", "Contabilidad"])

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        r = idx + 2
        ws_ext.append([
            str(row.get('Fecha_Clean', '')),
            str(row.get('Descripcion', row.get('DESCRIPCIÓN', ''))),
            float(row.get('Monto_Neto', 0)),
            f'=CONCATENATE(A{r}, "-", C{r}, "-", COUNTIFS($A$2:A{r}, A{r}, $C$2:C{r}, C{r}))',
            f'=IF(ISNUMBER(MATCH(D{r}, \'Auxiliar Contable\'!$F:$F, 0)), "CONCILIADO", "Pen Contabilidad")'
        ])

    max_ext_tbl = max(ws_ext.max_row, 2)
    tab2 = Table(displayName="TablaExtracto", ref=f"A1:E{max_ext_tbl}")
    tab2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_ext.add_table(tab2)

    # HOJA 4: Sugerencias de Cruce
    ws_sug = wb.create_sheet(title="Sugerencias de Cruce")
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

    # IDENTIFICACIÓN PREVIA DE EXCLUSIONES PARA HOJA PENDIENTES
    patron_gastos = r'IMPTO GOBIERNO|4X1000|COMISION|IVA COMISION|INTERESES|CUOTA MANEJO|MANTE SUCURSAL'
    gastos_refs_set = set()

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        ref_val = str(row.get('Descripcion', row.get('DESCRIPCIÓN', '')))
        if re.search(patron_gastos, ref_val, re.IGNORECASE):
            gastos_refs_set.add(ref_val)

    refs_sugeridas = set(df_sugerencias['Referencia Banco'].dropna().unique()) if (not df_sugerencias.empty and 'Referencia Banco' in df_sugerencias.columns) else set()

    # HOJA 5: Pendientes por Registrar (Filtro Inteligente: Sin Sugerencias y Sin Gastos)
    ws_pend = wb.create_sheet(title="Pendientes por Registrar")
    ws_pend.views.sheetView[0].showGridLines = True
    ws_pend.append(["Fecha Banco", "Referencia / Descripción", "Monto (+/-)", "Estado"])

    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        ref_val = str(row.get('Descripcion', row.get('DESCRIPCIÓN', '')))
        if (ref_val not in refs_sugeridas) and (ref_val not in gastos_refs_set):
            ws_pend.append([str(row.get('Fecha_Clean', '')), ref_val, float(row.get('Monto_Neto', 0)), "Pen Contabilidad"])

    max_pend_tbl = max(ws_pend.max_row, 2)
    tab5 = Table(displayName="TablaPendientes", ref=f"A1:D{max_pend_tbl}")
    tab5.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
    ws_pend.add_table(tab5)

    # HOJA 6: Gastos Bancarios
    ws_gastos = wb.create_sheet(title="Gastos Bancarios")
    ws_gastos.views.sheetView[0].showGridLines = True
    ws_gastos.append(["Fecha", "Descripción / Concepto Gasto", "Monto Gasto (-)", "Clasificación"])

    total_gastos = 0.0
    for idx, row in df_ext_proc.reset_index(drop=True).iterrows():
        ref_val = str(row.get('Descripcion', row.get('DESCRIPCIÓN', '')))
        if ref_val in gastos_refs_set:
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
# 5. NAVEGACIÓN Y PANEL STREAMLIT
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
        # CORRECCIÓN DE EXTENSIÓN: type=["xlsx"] permite seleccionar directamente tu archivo Excel (.xlsx)
        file_auxiliar = st.file_uploader("2. Cargar Auxiliar Contable (.xlsx)", type=["xlsx"], key="aux_diario_xlsx")

    if file_diario and file_auxiliar:
        try:
            df_diario_raw = pd.read_csv(file_diario, encoding="latin1", header=None)
        except:
            df_diario_raw = pd.read_csv(file_diario, encoding="utf-8", header=None)

        df_aux_raw = pd.read_excel(file_auxiliar)

        df_diario = procesar_diario_csv(df_diario_raw)
        df_aux = procesar_auxiliar(df_aux_raw)

        df_diario['Monto_Abs'] = df_diario['Monto_Neto'].round(2)
        df_aux['Monto_Abs'] = df_aux['Monto_Neto'].round(2)

        df_diario['Ocurrencia'] = df_diario.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1
        df_aux['Ocurrencia'] = df_aux.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1

        df_diario['LLAVE'] = df_diario['Fecha_Clean'] + "_" + df_diario['Monto_Abs'].astype(str) + "_" + df_diario['Ocurrencia'].astype(str)
        df_aux['LLAVE'] = df_aux['Fecha_Clean'] + "_" + df_aux['Monto_Abs'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)

        solo_diario = df_diario[~df_diario['LLAVE'].isin(df_aux['LLAVE'])]
        solo_aux = df_aux[~df_aux['LLAVE'].isin(df_diario['LLAVE'])]

        sug_gen = obtener_sugerencias_cruce(solo_aux, solo_diario, tol_monto=tol_pesos)

        st.success(f"✅ Análisis completado. Se hallaron {len(sug_gen)} sugerencias de cruce.")

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

        df_ext['Monto_Abs'] = df_ext['Monto_Neto'].round(2)
        df_aux['Monto_Abs'] = df_aux['Monto_Neto'].round(2)

        df_ext['Ocurrencia'] = df_ext.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1
        df_aux['Ocurrencia'] = df_aux.groupby(['Fecha_Clean', 'Monto_Abs']).cumcount() + 1

        df_ext['LLAVE'] = df_ext['Fecha_Clean'] + "_" + df_ext['Monto_Abs'].astype(str) + "_" + df_ext['Ocurrencia'].astype(str)
        df_aux['LLAVE'] = df_aux['Fecha_Clean'] + "_" + df_aux['Monto_Abs'].astype(str) + "_" + df_aux['Ocurrencia'].astype(str)

        solo_ext = df_ext[~df_ext['LLAVE'].isin(df_aux['LLAVE'])]
        solo_aux = df_aux[~df_aux['LLAVE'].isin(df_ext['LLAVE'])]

        sug_gen_m = obtener_sugerencias_cruce(solo_aux, solo_ext, tol_monto=tol_pesos)

        st.success(f"✅ Conciliación realizada. Se detectaron {len(sug_gen_m)} sugerencias de cruce.")

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
