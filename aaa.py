import streamlit as st
import pandas as pd
import altair as alt
import io

st.set_page_config(page_title="Gestor de Sobre-Stock", layout="wide")

st.title("🍷 Monitor de Sobre-Stock e Inventario Inmovilizado")
st.write("Sube los reportes semanales completos en Excel para detectar oportunidades de movimiento de inventario.")

# --- ZONAS DE CARGA ---
col1, col2 = st.columns(2)
with col1:
    archivo_anterior = st.file_uploader("Sube el Excel de la semana ANTERIOR", type=['csv', 'xlsx', 'xls'])
with col2:
    archivo_actual = st.file_uploader("Sube el Excel de la semana ACTUAL", type=['csv', 'xlsx', 'xls'])


def leer_archivo(archivo):
    nombre = archivo.name.lower()
    if nombre.endswith('.xlsx') or nombre.endswith('.xls'):
        xls = pd.ExcelFile(archivo, engine='openpyxl')
        hojas = xls.sheet_names
        hoja_objetivo = next((h for h in hojas if 'planilla' in h.strip().lower()), None)
        if hoja_objetivo is None:
            raise ValueError(f"No se encontró la pestaña 'PLANILLA'. Disponibles: {', '.join(hojas)}")
        return pd.read_excel(archivo, sheet_name=hoja_objetivo, engine='openpyxl')
    else:
        try:
            return pd.read_csv(archivo, encoding='latin-1', sep=';')
        except:
            archivo.seek(0)
            return pd.read_csv(archivo, encoding='latin-1', sep=',')


def formato_moneda(valor):
    return f"${valor:,.0f}".replace(",", ".")


# --- PROCESAMIENTO CENTRAL ---
if archivo_anterior and archivo_actual:
    try:
        df_ant = leer_archivo(archivo_anterior)
        df_act = leer_archivo(archivo_actual)

        # Limpiar nombres de columnas
        df_ant.columns = df_ant.columns.str.strip()
        df_act.columns = df_act.columns.str.strip()

        # Validar columnas clave
        columnas_clave = ['Material', 'LOTE', 'Texto breve de material', 'Libre utilización', 'Valor libre util.',
                          'Almacén', 'Estatus']
        for col in columnas_clave:
            if col not in df_ant.columns: df_ant[col] = ''
            if col not in df_act.columns: df_act[col] = ''

        df_cruce = pd.merge(df_ant[columnas_clave], df_act[columnas_clave],
                            on=['Material', 'LOTE', 'Texto breve de material'],
                            suffixes=('_Ant', '_Act'), how='outer').fillna(0)

        # Cálculos matemáticos básicos
        df_cruce['Variacion_Unidades'] = df_cruce['Libre utilización_Act'] - df_cruce['Libre utilización_Ant']
        df_cruce['Variacion_Valor'] = df_cruce['Valor libre util._Act'] - df_cruce['Valor libre util._Ant']


        # --- NUEVO REQUERIMIENTO: ESTADO DEL MATERIAL (NUEVO O YA ESTABA) ---
        def determinar_estado(row):
            if row['Libre utilización_Ant'] == 0 and row['Libre utilización_Act'] > 0:
                return "Material Nuevo"
            return "Ya Estaba"


        df_cruce['Estado Material'] = df_cruce.apply(determinar_estado, axis=1)


        # CÁLCULO DE % DE AUMENTO
        def calcular_porcentaje(row):
            ant = row['Libre utilización_Ant']
            var = row['Variacion_Unidades']
            if var <= 0:
                return "0%"
            if ant == 0 and var > 0:
                return "100%"  # Ya indicamos que es nuevo en la otra columna
            else:
                pct = (var / ant) * 100
                return f"{pct:.1f}%"


        df_cruce['% Aumento'] = df_cruce.apply(calcular_porcentaje, axis=1)

        # Etiquetas para gráficos
        df_cruce['LOTE'] = df_cruce['LOTE'].astype(str)
        df_cruce['Nombre_Grafico'] = df_cruce['Texto breve de material'] + " (Lote: " + df_cruce['LOTE'] + ")"

        # SUMA NO VIGENTES
        mascara_estatus = df_cruce['Estatus_Act'].astype(str).str.strip().str.upper() == 'NO VIGENTE'
        mascara_almacen = df_cruce['Almacén_Act'].astype(str).str.strip().str.upper() != 'FALSO'
        total_no_vigente = int(df_cruce[mascara_estatus & mascara_almacen]['Libre utilización_Act'].sum())

        # Filtro principal de alertas
        sobre_stock = df_cruce[(df_cruce['Variacion_Unidades'] > 0) | (
                    (df_cruce['Variacion_Unidades'] == 0) & (df_cruce['Libre utilización_Act'] > 500))].copy()

        if not df_cruce.empty:
            st.divider()

            tab1, tab2 = st.tabs(["📊 Dashboard Visual", "🔍 Reportes Especiales y Descargas"])

            with tab1:
                st.header("Dashboard Ejecutivo de Inventario")
                st.info(
                    "💡 **Tip para PDF:** Presiona `Ctrl + P` y selecciona 'Guardar como PDF' para imprimir esta pantalla.")

                unidades_nuevas = int(sobre_stock[sobre_stock['Variacion_Unidades'] > 0]['Variacion_Unidades'].sum())
                valor_nuevo_ingresado = sobre_stock[sobre_stock['Variacion_Unidades'] > 0]['Variacion_Valor'].sum()

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("🔴 Materiales con Aumento", len(sobre_stock[sobre_stock['Variacion_Unidades'] > 0]))
                m2.metric("📦 Unidades Ingresadas", f"{unidades_nuevas:,}".replace(",", "."))
                m3.metric("💰 Capital Retenido", formato_moneda(valor_nuevo_ingresado))
                m4.metric("⚠️ Total Unidades 'No Vigentes'", f"{total_no_vigente:,}".replace(",", "."))

                st.write("---")

                grafico_izq, grafico_der = st.columns(2)

                with grafico_izq:
                    st.write("**📈 Top 10: Mayor Aumento en la Semana (Ingresos)**")
                    top_aumentos = sobre_stock[sobre_stock['Variacion_Unidades'] > 0].sort_values(
                        by='Variacion_Unidades', ascending=False).head(10).copy()

                    if not top_aumentos.empty:
                        top_aumentos['Texto_Etiqueta'] = top_aumentos['Variacion_Unidades'].apply(
                            lambda x: f"+{int(x):,}".replace(',', '.'))
                        bars = alt.Chart(top_aumentos).mark_bar(color='#E15A97').encode(
                            x=alt.X('Variacion_Unidades:Q', title='Unidades Aumentadas'),
                            y=alt.Y('Nombre_Grafico:N', sort='-x', title='', axis=alt.Axis(labelLimit=800)),
                            tooltip=[alt.Tooltip('Texto breve de material:N', title='Material'),
                                     alt.Tooltip('LOTE:N', title='Lote'),
                                     alt.Tooltip('Variacion_Unidades:Q', title='Unidades Aumentadas')]
                        )
                        text = bars.mark_text(align='left', baseline='middle', dx=5, fontWeight='bold').encode(
                            text=alt.Text('Texto_Etiqueta:N'))
                        st.altair_chart((bars + text).properties(height=350), use_container_width=True)

                with grafico_der:
                    st.write("**📦 Top 10: Mayor Volumen Actual en Bodega**")
                    top_volumen = df_cruce.sort_values(by='Libre utilización_Act', ascending=False).head(10).copy()

                    if not top_volumen.empty:
                        top_volumen['Texto_Etiqueta'] = top_volumen['Libre utilización_Act'].apply(
                            lambda x: f"{int(x):,}".replace(',', '.'))
                        bars_vol = alt.Chart(top_volumen).mark_bar(color='#4A90E2').encode(
                            x=alt.X('Libre utilización_Act:Q', title='Stock Total Actual'),
                            y=alt.Y('Nombre_Grafico:N', sort='-x', title='', axis=alt.Axis(labelLimit=800)),
                            tooltip=[alt.Tooltip('Texto breve de material:N', title='Material'),
                                     alt.Tooltip('LOTE:N', title='Lote'),
                                     alt.Tooltip('Libre utilización_Act:Q', title='Stock Actual')]
                        )
                        text_vol = bars_vol.mark_text(align='left', baseline='middle', dx=5, fontWeight='bold').encode(
                            text=alt.Text('Texto_Etiqueta:N'))
                        st.altair_chart((bars_vol + text_vol).properties(height=350), use_container_width=True)

            with tab2:
                # --- REPORTE DE AUMENTOS PARA EL CLIENTE ---
                st.subheader("📈 Reporte de Aumentos de Inventario")

                solo_aumentos = df_cruce[df_cruce['Variacion_Unidades'] > 0].sort_values(by='Variacion_Unidades',
                                                                                         ascending=False)

                # INCLUYENDO LA NUEVA COLUMNA EN EL REPORTE
                columnas_aumentos = ['Material', 'Estado Material', 'Almacén_Act', 'LOTE', 'Texto breve de material',
                                     'Libre utilización_Ant', 'Libre utilización_Act', 'Variacion_Unidades',
                                     '% Aumento']

                st.dataframe(
                    solo_aumentos[columnas_aumentos],
                    use_container_width=True,
                    column_config={
                        "Libre utilización_Ant": st.column_config.NumberColumn("Semana Anterior", format="%d"),
                        "Libre utilización_Act": st.column_config.NumberColumn("Semana Actual", format="%d"),
                        "Variacion_Unidades": st.column_config.NumberColumn("Diferencia (+)", format="%d"),
                        "Almacén_Act": "Almacén Actual"
                    }
                )

                st.divider()

                # --- PLAN DE ACCIÓN ---
                st.subheader("📋 Detalle General y Plan de Acción")

                if not sobre_stock.empty:
                    sobre_stock = sobre_stock.sort_values(by='Libre utilización_Act', ascending=False)


                    def generar_recomendacion(row):
                        if row['Variacion_Unidades'] > 500:
                            return "🔴 Alerta: Fuerte ingreso. Confirmar justificación."
                        elif row['Variacion_Unidades'] > 0:
                            return "🟡 Aumento de stock. Vigilar rotación."
                        elif row['Libre utilización_Act'] > 5000:
                            return "🔵 Inmovilizado Alto: Evaluar Venta Ecommerce."
                        elif row['Libre utilización_Act'] > 1000:
                            return "🟢 Inmovilizado Medio: Sugerir Solicitudes Turismo."
                        else:
                            return "⚪ Inmovilizado Bajo: Armar packs promocionales."


                    sobre_stock['Recomendación'] = sobre_stock.apply(generar_recomendacion, axis=1)

                    st.write("**Filtros de Búsqueda**")
                    f1, f2 = st.columns(2)
                    with f1:
                        busqueda = st.text_input("🔍 Buscar por Código, Nombre o Lote:")
                    with f2:
                        filtro_alerta = st.multiselect("⚙️ Filtrar por tipo de Recomendación:",
                                                       options=sobre_stock['Recomendación'].unique(), default=[])

                    df_filtrado = sobre_stock.copy()
                    if busqueda:
                        busqueda = busqueda.lower()
                        mask = df_filtrado['Texto breve de material'].str.lower().str.contains(busqueda, na=False) | \
                               df_filtrado['Material'].str.lower().str.contains(busqueda, na=False) | \
                               df_filtrado['LOTE'].str.lower().str.contains(busqueda, na=False)
                        df_filtrado = df_filtrado[mask]

                    if filtro_alerta:
                        df_filtrado = df_filtrado[df_filtrado['Recomendación'].isin(filtro_alerta)]

                    # INCLUYENDO LA NUEVA COLUMNA EN EL PLAN DE ACCIÓN
                    columnas_plan = ['Material', 'Estado Material', 'Almacén_Act', 'LOTE', 'Texto breve de material',
                                     'Libre utilización_Act', 'Variacion_Unidades', 'Valor libre util._Act',
                                     'Recomendación']

                    st.dataframe(
                        df_filtrado[columnas_plan],
                        use_container_width=True,
                        column_config={
                            "Valor libre util._Act": st.column_config.NumberColumn("Valor Actual ($)", format="$ %d"),
                            "Libre utilización_Act": st.column_config.NumberColumn("Stock Actual", format="%d"),
                            "Variacion_Unidades": st.column_config.NumberColumn("Variación (Unid.)", format="%d"),
                            "Almacén_Act": "Almacén Actual"
                        }
                    )

                    # --- DESCARGA EXCEL ORDENADO ALFABÉTICAMENTE ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # 1. ORDENAMOS TODO POR 'Material' (DE LA A a la Z) ANTES DE GUARDAR EN EXCEL
                        excel_aumentos = solo_aumentos[columnas_aumentos].sort_values(by='Material', ascending=True)
                        excel_plan = df_filtrado[columnas_plan].sort_values(by='Material', ascending=True)
                        excel_total = df_cruce.sort_values(by='Material', ascending=True)

                        # 2. Guardamos las pestañas
                        excel_aumentos.to_excel(writer, index=False, sheet_name='Aumentos vs Semana Anterior')
                        excel_plan.to_excel(writer, index=False, sheet_name='Plan de Acción Filtrado')
                        excel_total.to_excel(writer, index=False, sheet_name='Inventario Total Histórico')

                        # 3. Auto-ajuste de columnas para que se lea perfecto
                        for sheet_name in writer.sheets:
                            worksheet = writer.sheets[sheet_name]
                            for col in worksheet.columns:
                                max_length = 0
                                column_letter = col[0].column_letter
                                for cell in col:
                                    try:
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(str(cell.value))
                                    except:
                                        pass
                                worksheet.column_dimensions[column_letter].width = max_length + 2

                    st.write("")
                    st.download_button(
                        label="📥 Descargar Reporte Completo (Excel Alfabético Auto-formateado)",
                        data=output.getvalue(),
                        file_name="Reporte_Inventario_Actualizado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.success("No hay alertas de inventario esta semana.")

    except Exception as e:
        st.error(f"Error procesando los datos: {e}")