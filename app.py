import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

def generar_excel_conciliacion(nombre_archivo="Conciliacion_Egresos_CtaCte.xlsx"):
    wb = openpyxl.Workbook()

    # -------------------------------------------------------------
    # 1. HOJA: Resumen Conciliación
    # -------------------------------------------------------------
    ws_resumen = wb.active
    ws_resumen.title = "Resumen Conciliación"
    ws_resumen.views.sheetView[0].showGridLines = True

    # Estilos
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

    # Encabezados
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

    # Filas con Fórmulas Estructuradas
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

    # -------------------------------------------------------------
    # 2. HOJA: Auxiliar Contable
    # -------------------------------------------------------------
    ws_aux = wb.create_sheet(title="Auxiliar Contable")
    ws_aux.views.sheetView[0].showGridLines = True
    ws_aux.append(["Fecha", "Documento", "Concepto/Detalle", "NOMBRE", "Egresos (-)", "LLAVE", "Extracto"])

    # Fila de ejemplo
    ws_aux.append([
        "2026-08-03", "RP-1-2489", "900633796", "CALIDAD COLOMBIA SERVICES SAS", 1590061,
        '=CONCATENATE(A2,E2,"-",COUNTIFS($A$2:A2,A2,$E$2:E2,E2))',
        '=_xlfn.XLOOKUP(F2, \'Extracto Bancario\'!$D:$D, \'Extracto Bancario\'!$B:$B, "NO ESTA EN BANCOS", 0)'
    ])

    tab1 = Table(displayName="Tabla1", ref=f"A1:G{ws_aux.max_row}")
    tab1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_aux.add_table(tab1)

    # -------------------------------------------------------------
    # 3. HOJA: Extracto Bancario
    # -------------------------------------------------------------
    ws_ext = wb.create_sheet(title="Extracto Bancario")
    ws_ext.views.sheetView[0].showGridLines = True
    ws_ext.append(["Fecha", "Referencia", "Retiros ()", "LLAVE", "Contabilidad", "Columna1", "Columna2"])

    # Fila de ejemplo
    ws_ext.append([
        "2026-08-03", " PAGO A PROVE CALIDAD COLOMBI", 1590061,
        '=CONCATENATE(A2,C2,"-",COUNTIFS($A$2:A2,A2,$C$2:C2,C2))',
        '=_xlfn.XLOOKUP(D2, \'Auxiliar Contable\'!$F:$F, \'Auxiliar Contable\'!$B:$B, "Pen Contabilidad", 0)',
        "", ""
    ])

    tab2 = Table(displayName="Tabla2", ref=f"A1:G{ws_ext.max_row}")
    tab2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_ext.add_table(tab2)

    wb.save(nombre_archivo)
    return nombre_archivo

generar_excel_conciliacion()
