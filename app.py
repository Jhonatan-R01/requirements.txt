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
            # Lógica para manejar diferentes codificaciones en el CSV (tildes, eñes, etc.)
            try:
                df_mov_data = pd.read_csv(archivo_mov, encoding='utf-8')
            except UnicodeDecodeError:
                archivo_mov.seek(0)  # Reiniciar el puntero del archivo
                try:
                    df_mov_data = pd.read_csv(archivo_mov, encoding='latin-1')
                except UnicodeDecodeError:
                    archivo_mov.seek(0)
                    df_mov_data = pd.read_csv(archivo_mov, encoding='utf-8-sig')

            # Lectura del Auxiliar Contable en Excel
            df_aux_data = pd.read_excel(archivo_aux)

            st.success("✅ Archivos leídos exitosamente.")

            # Generar Excel estructurado
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
