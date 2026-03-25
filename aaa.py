import streamlit as st
import pandas as pd
import altair as alt
import io
from datetime import date
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Gestor de Sobre-Stock", layout="wide")

st.title("🍷 Monitor de Sobre-Stock e Inventario Inmovilizado")
st.write("Sube los reportes semanales completos en Excel para detectar oportunidades de movimiento de inventario.")

# --- INICIAR CONEXIÓN A GOOGLE SHEETS ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.warning("La conexión a la base de datos no está configurada aún.")
    conn = None

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

        # --- CÁLCULOS EXACTOS PARA EL DASHBOARD (Imitando la Tabla Dinámica) ---
        df_act['Almacén'] = df_act['Almacén'].fillna('')
        df_act['Estatus'] = df_act['Estatus'].fillna('')
        df_act['Libre utilización'] = pd.to_numeric(df_act['Libre utilización'], errors='coerce').fillna(0)
        df_act['Valor libre util.'] = pd.to_numeric(df_act['Valor libre util.'], errors='coerce').fillna(0)

        # FILTROS ESTRICTOS:
        # 1. Estatus debe ser exactamente NO VIGENTE
        m_estatus_act = df_act['Estatus'].astype(str).str.strip().str.upper() == 'NO VIGENTE'

        # 2. Almacén NO debe ser Falso, False, #N/A, ni estar Vacío (en blanco)
        valores_a_excluir = ['FALSO', 'FALSE', '#N/A', 'NAN', '']
        m_almacen_act = ~df_act['Almacén'].astype(str).str.strip().str.upper().isin(valores_a_excluir)

        # Aplicamos los filtros a la Planilla base
        df_kpi_no_vigente = df_act[m_estatus_act & m_almacen_act]

        total_no_vigente = int(df_kpi_no_vigente['Libre utilización'].sum())
        valor_no_vigente = df_kpi_no_vigente['Valor libre util.'].sum()

        # --- PREPARACIÓN DEL CRUCE HISTÓRICO ---
        columnas_clave = ['Material', 'LOTE', 'Texto breve de material', 'Libre utilización', 'Valor libre util.',
                          'Almacén', 'Estatus']
        for col in columnas_clave:
            if col not in df_ant.columns: df_ant[col] = ''
            if col not in df_act.columns: df_act[col] = ''

        df_cruce = pd.merge(df_ant[columnas_clave], df_act[columnas_clave],
                            on=['Material', 'LOTE', 'Texto breve de material'],
                            suffixes=('_Ant', '_Act'), how='outer').fillna(0)

        df_cruce['Variacion_Unidades'] = df_cruce['Libre utilización_Act'] - df_cruce['Libre utilización_Ant']
        df_cruce['Variacion_Valor'] = df_cruce['Valor libre util._Act'] - df_cruce['Valor libre util._Ant']


        def determinar_estado(row):
            if row['Libre utilización_Ant'] == 0 and row['Libre utilización_Act'] > 0: return "Material Nuevo"
            return "Ya Estaba"


        df_cruce['Estado Material'] = df_cruce.apply(determinar_estado, axis=1)


        def calcular_porcentaje(row):
            ant = row['Libre utilización_Ant']
            var = row['Variacion_Unidades']
            if var <= 0: return "0%"
            if ant == 0 and var > 0:
                return "100%"
            else:
                return f"{(var / ant) * 100:.1f}%"


        df_cruce['% Aumento'] = df_cruce.apply(calcular_porcentaje, axis=1)

        df_cruce['LOTE'] = df_cruce['LOTE'].astype(str)
        df_cruce['Nombre_Grafico'] = df_cruce['Texto breve de material'] + " (Lote: " + df_cruce['LOTE'] + ")"

        sobre_stock = df_cruce[(df_cruce['Variacion_Unidades'] > 0) | (
                    (df_cruce['Variacion_Unidades'] == 0) & (df_cruce['Libre utilización_Act'] > 500))].copy()

        if not df_cruce.empty:
            st.divider()
            tab1, tab2, tab3 = st.tabs(["📊 Dashboard Visual", "🔍 Reportes y Descargas", "☁️ Trazabilidad Histórica"])

            with tab1:
                st.header("Dashboard Ejecutivo de Inventario")

                unidades_nuevas = int(sobre_stock[sobre_stock['Variacion_Unidades'] > 0]['Variacion_Unidades'].sum())
                valor_nuevo_ingresado = sobre_stock[sobre_stock['Variacion_Unidades'] > 0]['Variacion_Valor'].sum()

                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("🔴 Mat. con Aumento", len(sobre_stock[sobre_stock['Variacion_Unidades'] > 0]))
                m2.metric("📦 Unid. Ingresadas", f"{unidades_nuevas:,}".replace(",", "."))
                m3.metric("💰 Capital Retenido", formato_moneda(valor_nuevo_ingresado))
                m4.metric("⚠️ Unid. 'No Vigentes'", f"{total_no_vigente:,}".replace(",", "."))
                m5.metric("🏦 Capital 'No Vigente'", formato_moneda(valor_no_vigente))
                st.write("---")

                grafico_izq, grafico_der = st.columns(2)
                with grafico_izq:
                    top_aumentos = sobre_stock[sobre_stock['Variacion_Unidades'] > 0].sort_values(
                        by='Variacion_Unidades', ascending=False).head(10).copy()
                    if not top_aumentos.empty:
                        top_aumentos['Texto_Etiqueta'] = top_aumentos['Variacion_Unidades'].apply(
                            lambda x: f"+{int(x):,}".replace(',', '.'))
                        bars = alt.Chart(top_aumentos).mark_bar(color='#E15A97').encode(x=alt.X('Variacion_Unidades:Q'),
                                                                                        y=alt.Y('Nombre_Grafico:N',
                                                                                                sort='-x',
                                                                                                axis=alt.Axis(
                                                                                                    labelLimit=800)))
                        text = bars.mark_text(align='left', baseline='middle', dx=5, fontWeight='bold').encode(
                            text=alt.Text('Texto_Etiqueta:N'))
                        st.altair_chart((bars + text).properties(height=350), use_container_width=True)

                with grafico_der:
                    top_volumen = df_cruce.sort_values(by='Libre utilización_Act', ascending=False).head(10).copy()
                    if not top_volumen.empty:
                        top_volumen['Texto_Etiqueta'] = top_volumen['Libre utilización_Act'].apply(
                            lambda x: f"{int(x):,}".replace(',', '.'))
                        bars_vol = alt.Chart(top_volumen).mark_bar(color='#4A90E2').encode(
                            x=alt.X('Libre utilización_Act:Q'),
                            y=alt.Y('Nombre_Grafico:N', sort='-x', axis=alt.Axis(labelLimit=800)))
                        text_vol = bars_vol.mark_text(align='left', baseline='middle', dx=5, fontWeight='bold').encode(
                            text=alt.Text('Texto_Etiqueta:N'))
                        st.altair_chart((bars_vol + text_vol).properties(height=350), use_container_width=True)

            with tab2:
                # --- HERRAMIENTA DE AUDITORÍA (NUEVO) ---
                st.subheader("🕵️‍♂️ Auditoría de KPI: ¿Qué se sumó?")
                with st.expander("Ver lista de materiales sumados en el Capital 'No Vigente'"):
                    st.write(f"Total filas contabilizadas: {len(df_kpi_no_vigente)}")
                    st.write("Filtros aplicados: Estatus='NO VIGENTE' | Almacén no es 'FALSO' ni está en blanco.")
                    st.dataframe(df_kpi_no_vigente[['Material', 'Almacén', 'Estatus', 'Valor libre util.']])

                st.write("---")
                st.subheader("📈 Reporte de Aumentos de Inventario")
                solo_aumentos = df_cruce[df_cruce['Variacion_Unidades'] > 0].sort_values(by='Variacion_Unidades',
                                                                                         ascending=False)
                columnas_aumentos = ['Material', 'Estado Material', 'Almacén_Act', 'LOTE', 'Texto breve de material',
                                     'Libre utilización_Ant', 'Libre utilización_Act', 'Variacion_Unidades',
                                     '% Aumento']

                st.dataframe(solo_aumentos[columnas_aumentos], use_container_width=True, column_config={
                    "Libre utilización_Ant": st.column_config.NumberColumn("Semana Anterior", format="%d"),
                    "Libre utilización_Act": st.column_config.NumberColumn("Semana Actual", format="%d"),
                    "Variacion_Unidades": st.column_config.NumberColumn("Diferencia (+)", format="%d"),
                    "Almacén_Act": "Almacén Actual"})

                st.divider()
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

                    columnas_plan = ['Material', 'Estado Material', 'Almacén_Act', 'LOTE', 'Texto breve de material',
                                     'Libre utilización_Act', 'Variacion_Unidades', 'Valor libre util._Act',
                                     'Recomendación']
                    st.dataframe(sobre_stock[columnas_plan], use_container_width=True, column_config={
                        "Valor libre util._Act": st.column_config.NumberColumn("Valor Actual ($)", format="$ %d"),
                        "Libre utilización_Act": st.column_config.NumberColumn("Stock Actual", format="%d"),
                        "Variacion_Unidades": st.column_config.NumberColumn("Variación (Unid.)", format="%d"),
                        "Almacén_Act": "Almacén Actual"})

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        excel_aumentos = solo_aumentos[columnas_aumentos].sort_values(by='Material', ascending=True)
                        excel_plan = sobre_stock[columnas_plan].sort_values(by='Material', ascending=True)
                        excel_total = df_cruce.sort_values(by='Material', ascending=True)

                        excel_aumentos.to_excel(writer, index=False, sheet_name='Aumentos vs Semana Anterior')
                        excel_plan.to_excel(writer, index=False, sheet_name='Plan de Acción Filtrado')
                        excel_total.to_excel(writer, index=False, sheet_name='Inventario Total Histórico')

                        for sheet_name in writer.sheets:
                            worksheet = writer.sheets[sheet_name]
                            for col in worksheet.columns:
                                max_length = 0
                                column_letter = col[0].column_letter
                                for cell in col:
                                    try:
                                        if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                                    except:
                                        pass
                                worksheet.column_dimensions[column_letter].width = max_length + 2

                    st.write("")
                    st.download_button(label="📥 Descargar Reporte Completo (Excel Alfabético)", data=output.getvalue(),
                                       file_name="Reporte_Inventario_Actualizado.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.success("No hay alertas de inventario esta semana.")

            with tab3:
                st.header("Base de Datos Histórica (Google Sheets)")
                if conn is not None:
                    with st.form("form_guardar_bd"):
                        col_fecha, col_btn = st.columns([1, 2])
                        with col_fecha:
                            fecha_registro = st.date_input("Fecha de esta foto de inventario:", date.today())
                        with col_btn:
                            st.write("");
                            st.write("")
                            guardar = st.form_submit_button("💾 Enviar 'Semana Actual' a Google Sheets")

                        if guardar:
                            with st.spinner("Conectando con Google Sheets..."):
                                try:
                                    df_hist = conn.read(worksheet="Historial", usecols=list(range(6)), ttl=0).dropna(
                                        how="all")
                                except Exception:
                                    df_hist = pd.DataFrame(columns=['Fecha_Registro', 'Material', 'LOTE', 'Texto_breve',
                                                                    'Libre_utilizacion', 'Valor'])

                                df_para_bd = df_act[['Material', 'LOTE', 'Texto breve de material', 'Libre utilización',
                                                     'Valor libre util.']].copy()
                                df_para_bd.rename(columns={'Texto breve de material': 'Texto_breve',
                                                           'Libre utilización': 'Libre_utilizacion',
                                                           'Valor libre util.': 'Valor'}, inplace=True)
                                df_para_bd.insert(0, 'Fecha_Registro', str(fecha_registro))

                                if not df_hist.empty:
                                    df_hist['Fecha_Registro'] = df_hist['Fecha_Registro'].astype(str)
                                    df_hist = df_hist[df_hist['Fecha_Registro'] != str(fecha_registro)]

                                df_updated = pd.concat([df_hist, df_para_bd], ignore_index=True)
                                conn.update(worksheet="Historial", data=df_updated)
                                st.success(f"¡Inventario del {fecha_registro} guardado exitosamente!")

                    st.divider()
                    st.subheader("📈 Análisis de Tendencias")
                    if st.button("🔄 Cargar Gráficos Históricos"):
                        with st.spinner("Descargando historial desde Google..."):
                            df_hist_cloud = conn.read(worksheet="Historial", usecols=list(range(6)), ttl=0).dropna(
                                how="all")
                            if not df_hist_cloud.empty:
                                df_hist_cloud['Fecha_Registro'] = pd.to_datetime(df_hist_cloud['Fecha_Registro'])
                                material_seleccionado = st.selectbox("Selecciona un material:",
                                                                     df_hist_cloud['Texto_breve'].unique())
                                datos_grafico = df_hist_cloud[
                                    df_hist_cloud['Texto_breve'] == material_seleccionado].copy()
                                if not datos_grafico.empty:
                                    linea = alt.Chart(datos_grafico).mark_line(point=True, color='#FF5722',
                                                                               strokeWidth=3).encode(
                                        x=alt.X('Fecha_Registro:T', title='Fecha'),
                                        y=alt.Y('Libre_utilizacion:Q', title='Stock Total (Unidades)'),
                                        color=alt.Color('LOTE:N', legend=alt.Legend(title="Lotes")))
                                    st.altair_chart(linea.properties(height=400), use_container_width=True)
                            else:
                                st.info("Aún no has guardado ningún dato.")
    except Exception as e:
        st.error(f"Error procesando los datos: {e}")