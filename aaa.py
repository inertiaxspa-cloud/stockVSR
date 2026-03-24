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

        df_ant.columns = df_ant.columns.str.strip()
        df_act.columns = df_act.columns.str.strip()

        columnas_clave = ['Material', 'LOTE', 'Texto breve de material', 'Libre utilización', 'Valor libre util.']

        df_cruce = pd.merge(df_ant[columnas_clave], df_act[columnas_clave],
                            on=['Material', 'LOTE', 'Texto breve de material'],
                            suffixes=('_Ant', '_Act'), how='outer').fillna(0)

        df_cruce['Variacion_Unidades'] = df_cruce['Libre utilización_Act'] - df_cruce['Libre utilización_Ant']
        df_cruce['Variacion_Valor'] = df_cruce['Valor libre util._Act'] - df_cruce['Valor libre util._Ant']

        df_cruce['LOTE'] = df_cruce['LOTE'].astype(str)
        df_cruce['Nombre_Grafico'] = df_cruce['Texto breve de material'] + " (Lote: " + df_cruce['LOTE'] + ")"

        sobre_stock = df_cruce[(df_cruce['Variacion_Unidades'] > 0) | (
                    (df_cruce['Variacion_Unidades'] == 0) & (df_cruce['Libre utilización_Act'] > 500))].copy()

        if not df_cruce.empty:
            st.divider()

            # --- PESTAÑAS DE NAVEGACIÓN ---
            tab1, tab2 = st.tabs(["📊 Dashboard Visual", "🔍 Tabla Dinámica y Descargas"])

            with tab1:
                st.header("Dashboard Ejecutivo de Inventario")
                st.info(
                    "💡 **Tip para PDF:** Presiona `Ctrl + P` y selecciona 'Guardar como PDF' para imprimir esta pantalla.")

                total_aumentos = len(sobre_stock[sobre_stock['Variacion_Unidades'] > 0])
                unidades_nuevas = int(sobre_stock[sobre_stock['Variacion_Unidades'] > 0]['Variacion_Unidades'].sum())
                valor_nuevo_ingresado = sobre_stock[sobre_stock['Variacion_Unidades'] > 0]['Variacion_Valor'].sum()

                m1, m2, m3 = st.columns(3)
                m1.metric("🔴 Materiales con Aumento de Stock", total_aumentos)
                m2.metric("📦 Total de Unidades Nuevas", f"{unidades_nuevas:,}".replace(",", "."))
                m3.metric("💰 Capital Retenido en Nuevos Ingresos", formato_moneda(valor_nuevo_ingresado))

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
                            text=alt.Text('Texto_Etiqueta:N')
                        )
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
                            text=alt.Text('Texto_Etiqueta:N')
                        )
                        st.altair_chart((bars_vol + text_vol).properties(height=350), use_container_width=True)

            with tab2:
                st.subheader("📋 Detalle y Plan de Acción")

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

                    # --- NUEVA SECCIÓN DE FILTROS INTERACTIVOS ---
                    st.write("**Filtros de Búsqueda**")
                    f1, f2 = st.columns(2)

                    with f1:
                        # Buscador de texto libre
                        busqueda = st.text_input(
                            "🔍 Buscar por Código, Nombre o Lote (Ej: CASA REAL, CV2162, o B41803):")

                    with f2:
                        # Selector múltiple de prioridades
                        opciones_alerta = sobre_stock['Recomendación'].unique()
                        filtro_alerta = st.multiselect("⚙️ Filtrar por tipo de Recomendación:", options=opciones_alerta,
                                                       default=[])

                    # Aplicar filtros a los datos
                    df_filtrado = sobre_stock.copy()

                    if busqueda:
                        busqueda = busqueda.lower()
                        # Busca coincidencias en Material, Texto o Lote
                        mask = df_filtrado['Texto breve de material'].str.lower().str.contains(busqueda, na=False) | \
                               df_filtrado['Material'].str.lower().str.contains(busqueda, na=False) | \
                               df_filtrado['LOTE'].str.lower().str.contains(busqueda, na=False)
                        df_filtrado = df_filtrado[mask]

                    if filtro_alerta:
                        df_filtrado = df_filtrado[df_filtrado['Recomendación'].isin(filtro_alerta)]

                    st.write(f"Mostrando {len(df_filtrado)} registros:")

                    columnas_mostrar = ['Material', 'LOTE', 'Texto breve de material', 'Libre utilización_Act',
                                        'Variacion_Unidades', 'Valor libre util._Act', 'Recomendación']

                    # Renderizar tabla con los datos filtrados
                    st.dataframe(
                        df_filtrado[columnas_mostrar],
                        use_container_width=True,
                        column_config={
                            "Valor libre util._Act": st.column_config.NumberColumn("Valor Actual ($)", format="$ %d"),
                            "Libre utilización_Act": st.column_config.NumberColumn("Stock Actual", format="%d"),
                            "Variacion_Unidades": st.column_config.NumberColumn("Variación (Unid.)", format="%d")
                        }
                    )

                    # --- DESCARGA EXCEL CON DATOS FILTRADOS ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Descarga lo que sea que esté viendo el usuario tras filtrar
                        df_filtrado[columnas_mostrar].to_excel(writer, index=False,
                                                               sheet_name='Plan de Acción Filtrado')
                        # Descarga el inventario histórico completo siempre en la hoja 2
                        df_cruce.sort_values(by='Libre utilización_Act', ascending=False).to_excel(writer, index=False,
                                                                                                   sheet_name='Inventario Total Histórico')

                    st.download_button(
                        label="📥 Descargar Tabla Actual (Excel .xlsx)",
                        data=output.getvalue(),
                        file_name="Reporte_Inventario_Filtrado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.success("No hay alertas de inventario esta semana.")

    except Exception as e:
        st.error(f"Error procesando los datos: {e}")