def procesar_auxiliar(df_raw):
    # Detección dinámica de la fila de encabezados asegurando que cada celda sea string
    idx_header = None
    for i, row in df_raw.iterrows():
        # Convertir explícitamente cada valor a string en minúsculas
        row_str = [str(cell).lower() for cell in row.values]
        if any('debito' in cell or 'débito' in cell for cell in row_str):
            idx_header = i
            break
            
    if idx_header is not None:
        df_raw.columns = df_raw.iloc[idx_header].values
        df = df_raw.iloc[idx_header + 1:].copy()
    else:
        df = df_raw.copy()

    # Búsqueda flexible de columnas claves
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
