import streamlit as st
import pandas as pd
import altair as alt
import io
from datetime import date
from streamlit_gsheets import GSheetsConnection
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gestor de Sobre-Stock", layout="wide")
st.title("🍷 Monitor de Sobre-Stock e Inventario Inmovilizado")
st.write("Sube entre 2 y 5 reportes semanales en orden cronológico para comparar la evolución del inventario.")

# ── Google Sheets ─────────────────────────────────────────────────────────────
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    conn = None

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────
COL_ALMACEN   = 'Almacen origen'   # columna de almacén físico real
COLS_NUMERICAS = {'Libre utilización', 'Valor libre util.'}
VALS_INVALIDOS = ['FALSO', 'FALSE', '#N/A', 'NAN', '']

COLS_CLAVE = ['Material', 'LOTE', 'Texto breve de material',
              'Libre utilización', 'Valor libre util.', COL_ALMACEN, 'Estatus']

COLS_REQUERIDAS = ['Material', 'LOTE', 'Texto breve de material',
                   'Libre utilización', 'Valor libre util.', COL_ALMACEN, 'Estatus']

# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN: Leer y limpiar archivo
# ─────────────────────────────────────────────────────────────────────────────
def leer_archivo(archivo):
    nombre = archivo.name.lower()
    if nombre.endswith('.xlsx') or nombre.endswith('.xls'):
        try:
            xls = pd.ExcelFile(archivo, engine='openpyxl')
        except Exception as e:
            raise ValueError(f"No se pudo abrir '{archivo.name}': {e}")
        hojas = xls.sheet_names
        hoja  = next((h for h in hojas if 'planilla' in h.strip().lower()), None)
        if hoja is None:
            raise ValueError(f"No se encontró la pestaña 'PLANILLA' en '{archivo.name}'.\nHojas disponibles: {', '.join(hojas)}")
        df = pd.read_excel(archivo, sheet_name=hoja, engine='openpyxl')
    else:
        try:
            df = pd.read_csv(archivo, encoding='latin-1', sep=';')
        except Exception:
            archivo.seek(0)
            try:
                df = pd.read_csv(archivo, encoding='latin-1', sep=',')
            except Exception as e:
                raise ValueError(f"No se pudo leer el CSV '{archivo.name}': {e}")

    df.columns  = df.columns.str.strip()
    filas_orig  = len(df)
    df          = df.dropna(subset=['Material']).reset_index(drop=True)
    descartes   = filas_orig - len(df)
    return df, descartes


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN: Normalizar y consolidar duplicados
# ─────────────────────────────────────────────────────────────────────────────
def normalizar(df):
    df = df.copy()
    df['Material']                = df['Material'].astype(str).str.strip()
    df['LOTE']                    = df['LOTE'].astype(str).str.strip().str.upper().replace('NAN', '')
    df['Texto breve de material'] = df['Texto breve de material'].fillna('').astype(str).str.strip()
    df['Libre utilización']       = pd.to_numeric(df.get('Libre utilización', 0), errors='coerce').fillna(0)
    df['Valor libre util.']       = pd.to_numeric(df.get('Valor libre util.', 0), errors='coerce').fillna(0)
    df[COL_ALMACEN]               = df.get(COL_ALMACEN, pd.Series([''] * len(df))).fillna('').astype(str).str.strip()
    df['Estatus']                 = df.get('Estatus', pd.Series([''] * len(df))).fillna('').astype(str).str.strip()
    return df


def consolidar_duplicados(df):
    keys = ['Material', 'LOTE', 'Texto breve de material']
    if not all(c in df.columns for c in keys):
        return df
    nums  = df.select_dtypes(include='number').columns.tolist()
    texto = [c for c in df.columns if c not in nums and c not in keys]
    agg   = {c: 'sum'   for c in nums}
    agg.update({c: 'first' for c in texto})
    return df.groupby(keys, as_index=False).agg(agg)


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN: Comparar un par de DataFrames
# ─────────────────────────────────────────────────────────────────────────────
def comparar_par(df_ant, df_act):
    # Asegurar columnas faltantes
    for col in COLS_CLAVE:
        for df_tmp in [df_ant, df_act]:
            if col not in df_tmp.columns:
                df_tmp[col] = 0 if col in COLS_NUMERICAS else ''

    # Merge outer
    df_cruce = pd.merge(
        df_ant[COLS_CLAVE], df_act[COLS_CLAVE],
        on=['Material', 'LOTE', 'Texto breve de material'],
        suffixes=('_Ant', '_Act'),
        how='outer'
    )

    # fillna inteligente
    for c in df_cruce.columns:
        if df_cruce[c].dtype == 'object':
            df_cruce[c] = df_cruce[c].fillna('')
        else:
            df_cruce[c] = df_cruce[c].fillna(0)

    # Forzar numérico en columnas de cálculo
    for c in ['Libre utilización_Ant', 'Libre utilización_Act',
              'Valor libre util._Ant',  'Valor libre util._Act']:
        df_cruce[c] = pd.to_numeric(df_cruce[c], errors='coerce').fillna(0)

    df_cruce['Variacion_Unidades'] = df_cruce['Libre utilización_Act'] - df_cruce['Libre utilización_Ant']
    df_cruce['Variacion_Valor']    = df_cruce['Valor libre util._Act'] - df_cruce['Valor libre util._Ant']

    # Clasificación
    def estado(row):
        if row['Libre utilización_Ant'] == 0 and row['Libre utilización_Act'] > 0: return '🆕 Material Nuevo'
        if row['Libre utilización_Act'] == 0 and row['Libre utilización_Ant'] > 0: return '🚫 Desaparecido'
        return 'Ya Estaba'

    def pct(row):
        ant = float(row['Libre utilización_Ant'])
        var = float(row['Variacion_Unidades'])
        if var <= 0:   return '0%'
        if ant == 0:   return '100% (Nuevo)'
        return f'+{(var / ant) * 100:.1f}%'

    df_cruce['Estado Material'] = df_cruce.apply(estado, axis=1)
    df_cruce['% Aumento']       = df_cruce.apply(pct,    axis=1)
    lote_disp                   = df_cruce['LOTE'].replace({'': '—', 'NAN': '—'})
    df_cruce['Nombre_Grafico']  = df_cruce['Texto breve de material'].astype(str) + ' (Lote: ' + lote_disp + ')'

    solo_aumentos     = df_cruce[df_cruce['Variacion_Unidades'] > 0].copy()
    solo_bajadas      = df_cruce[df_cruce['Variacion_Unidades'] < 0].copy()
    materiales_nuevos = df_cruce[df_cruce['Estado Material'] == '🆕 Material Nuevo'].copy()
    desaparecidos     = df_cruce[df_cruce['Estado Material'] == '🚫 Desaparecido'].copy()
    sobre_stock       = df_cruce[
        (df_cruce['Variacion_Unidades'] > 0) |
        ((df_cruce['Variacion_Unidades'] == 0) & (df_cruce['Libre utilización_Act'] > 500))
    ].copy()

    return dict(
        cruce=df_cruce, sobre_stock=sobre_stock,
        subidas=solo_aumentos, bajadas=solo_bajadas,
        nuevos=materiales_nuevos, desaparecidos=desaparecidos,
    )


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN: KPIs de No Vigente (sobre df_act limpio)
# ─────────────────────────────────────────────────────────────────────────────
def calcular_no_vigente(df_act):
    df = df_act.copy()
    df[COL_ALMACEN] = df[COL_ALMACEN].fillna('').astype(str).str.strip()
    df['Estatus']   = df['Estatus'].fillna('').astype(str).str.strip()
    df['Libre utilización'] = pd.to_numeric(df['Libre utilización'], errors='coerce').fillna(0)
    df['Valor libre util.'] = pd.to_numeric(df['Valor libre util.'], errors='coerce').fillna(0)

    m_est = df['Estatus'].str.upper() == 'NO VIGENTE'
    m_alm = ~df[COL_ALMACEN].str.upper().isin(VALS_INVALIDOS)
    df_nv = df[m_est & m_alm]
    return df_nv, int(df_nv['Libre utilización'].sum()), df_nv['Valor libre util.'].sum(), df['Valor libre util.'].sum()


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN: Formato condicional en Excel
# ─────────────────────────────────────────────────────────────────────────────
def dar_formato_excel(writer, df, sheet_name, col_variacion=None, col_totales=None):
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    ws = writer.sheets[sheet_name]

    fill_hdr   = PatternFill('solid', fgColor='4A235A')
    fill_pos   = PatternFill('solid', fgColor='C6EFCE')
    fill_neg   = PatternFill('solid', fgColor='FFC7CE')
    fill_tot   = PatternFill('solid', fgColor='D9D9D9')
    font_hdr   = Font(bold=True, color='FFFFFF')
    font_tot   = Font(bold=True)
    borde      = Border(bottom=Side(border_style='thin'), top=Side(border_style='thin'))
    centro     = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for cell in ws[1]:
        cell.fill = fill_hdr; cell.font = font_hdr; cell.alignment = centro

    for idx, col in enumerate(ws.columns, 1):
        max_len = max((len(str(c.value or '')) for c in col), default=0)
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 4, 45)

    if col_variacion and col_variacion in df.columns:
        ci = df.columns.get_loc(col_variacion) + 1
        for r in range(2, ws.max_row + 1):
            try:
                v = float(ws.cell(r, ci).value or 0)
                ws.cell(r, ci).fill = fill_pos if v > 0 else (fill_neg if v < 0 else ws.cell(r, ci).fill)
            except (ValueError, TypeError):
                pass

    if col_totales:
        tr = ws.max_row + 1
        for cn in col_totales:
            if cn in df.columns:
                ci = df.columns.get_loc(cn) + 1
                c  = ws.cell(tr, ci, value=df[cn].sum())
                c.fill = fill_tot; c.font = font_tot; c.border = borde
        ws.cell(tr, 1, value='TOTAL').font = font_tot


def formato_moneda(v):
    return f'${v:,.0f}'.replace(',', '.')


# ─────────────────────────────────────────────────────────────────────────────
# UI: CARGA DE ARCHIVOS (2 a 5 semanas)
# ─────────────────────────────────────────────────────────────────────────────
st.subheader('📂 Selecciona las semanas a comparar')
st.caption('Sube los archivos en orden **cronológico** (más antiguo → más reciente). Mínimo 2, máximo 5.')

num_semanas = st.slider('¿Cuántas semanas quieres comparar?', min_value=2, max_value=5, value=2)

cols_upload = st.columns(num_semanas)
archivos_cargados = []
for i, col in enumerate(cols_upload):
    with col:
        etiqueta = '🗓️ Semana base (más antigua)' if i == 0 else (
                   '🗓️ Semana más reciente'        if i == num_semanas - 1 else
                   f'🗓️ Semana {i + 1}')
        f = st.file_uploader(etiqueta, type=['csv', 'xlsx', 'xls'], key=f'semana_{i}')
        archivos_cargados.append(f)

archivos_validos = [f for f in archivos_cargados if f is not None]

# ─────────────────────────────────────────────────────────────────────────────
# PROCESAMIENTO
# ─────────────────────────────────────────────────────────────────────────────
if len(archivos_validos) < 2:
    st.info('⬆️ Sube al menos **2 archivos** para comenzar el análisis.')
    st.stop()

# ── PASO 1: Leer todos los archivos ──────────────────────────────────────────
dfs_semanas = []
for archivo in archivos_validos:
    try:
        df, descartes = leer_archivo(archivo)
    except ValueError as e:
        st.error(f'❌ {e}'); st.stop()
    except Exception as e:
        st.error(f'❌ Error inesperado al leer "{archivo.name}": {e}'); st.stop()

    # Validar columnas requeridas
    faltantes = [c for c in COLS_REQUERIDAS if c not in df.columns]
    if faltantes:
        st.error(
            f'❌ Al archivo **{archivo.name}** le faltan las columnas: `{"`, `".join(faltantes)}`\n\n'
            f'Columnas encontradas: `{"`, `".join(df.columns.tolist())}`'
        )
        st.stop()

    df = normalizar(df)

    # Detectar y consolidar duplicados
    keys = ['Material', 'LOTE', 'Texto breve de material']
    dupes = df.duplicated(subset=keys, keep=False).sum()
    if dupes > 0:
        st.warning(f'⚠️ Se encontraron **{dupes} filas duplicadas** (mismo Material + LOTE) en **{archivo.name}**. Se consolidarán sumando su stock.')
        df = consolidar_duplicados(df)

    dfs_semanas.append({'nombre': archivo.name, 'df': df, 'descartes': descartes})

# ── Indicador de calidad ──────────────────────────────────────────────────────
with st.expander('🔍 Calidad de los datos cargados', expanded=False):
    for s in dfs_semanas:
        df_s = s['df']
        alm_inv = df_s[COL_ALMACEN].str.upper().isin(VALS_INVALIDOS).sum()
        st.markdown(f"**{s['nombre']}** — {len(df_s):,} filas válidas"
                    + (f" · ⚠️ {s['descartes']} vacías descartadas" if s['descartes'] else '')
                    + (f" · ⚠️ {alm_inv} almacenes inválidos (False/FALSO)" if alm_inv else ''))

# ── PASO 2: Comparaciones consecutivas ───────────────────────────────────────
comparaciones = []
for i in range(len(dfs_semanas) - 1):
    sa = dfs_semanas[i]
    sb = dfs_semanas[i + 1]
    try:
        resultado = comparar_par(sa['df'], sb['df'])
    except Exception as e:
        st.error(f'❌ Error al comparar "{sa["nombre"]}" vs "{sb["nombre"]}": {e}'); st.stop()

    label_a = sa['nombre'].replace('.xlsx', '').replace('.xls', '').replace('.csv', '')
    label_b = sb['nombre'].replace('.xlsx', '').replace('.xls', '').replace('.csv', '')
    comparaciones.append({'label': f'{label_a} → {label_b}', 'a': sa, 'b': sb, **resultado})

# Comparación principal = la más reciente
comp = comparaciones[-1]
df_cruce      = comp['cruce']
solo_aumentos = comp['subidas']
solo_bajadas  = comp['bajadas']
mat_nuevos    = comp['nuevos']
desaparecidos = comp['desaparecidos']
sobre_stock   = comp['sobre_stock']
df_act_actual = comp['b']['df']
df_ant_actual = comp['a']['df']

# KPIs No Vigente sobre df_act de la comparación principal
df_kpi_nv, total_nv, valor_nv, valor_total = calcular_no_vigente(df_act_actual)

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tabs_labels = ['📊 Dashboard Visual', '🔍 Reportes y Descargas', '☁️ Trazabilidad Histórica']
if len(comparaciones) > 1:
    tabs_labels.append('📅 Evolución Multi-Semana')

tabs = st.tabs(tabs_labels)
tab1, tab2, tab3 = tabs[0], tabs[1], tabs[2]
tab4 = tabs[3] if len(comparaciones) > 1 else None

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    comp_label = comp['label']
    st.header(f'Dashboard Ejecutivo — {comp_label}')

    uni_sub  = int(solo_aumentos['Variacion_Unidades'].sum())
    uni_baj  = int(solo_bajadas['Variacion_Unidades'].sum())
    val_ing  = solo_aumentos['Variacion_Valor'].sum()
    pct_inm  = (valor_nv / valor_total * 100) if valor_total > 0 else 0

    # Fila 1
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric('📈 Subieron',           len(solo_aumentos))
    m2.metric('📉 Bajaron',            len(solo_bajadas))
    m3.metric('🆕 Materiales Nuevos',  len(mat_nuevos))
    m4.metric('🚫 Desaparecidos',      len(desaparecidos))
    m5.metric('📦 Unidades Ingresadas', f'{uni_sub:,}'.replace(',', '.'))

    st.write('---')

    # Fila 2
    k1, k2, k3, k4 = st.columns(4)
    k1.metric('📉 Unidades Bajaron',    f'{abs(uni_baj):,}'.replace(',', '.'),
              delta=f'{uni_baj:,}'.replace(',', '.'), delta_color='inverse')
    k2.metric('💰 Capital Ingresado',   formato_moneda(val_ing))
    k3.metric("⚠️ Unid. 'No Vigentes'", f'{total_nv:,}'.replace(',', '.'))
    k4.metric("🏦 Capital 'No Vigente'", formato_moneda(valor_nv),
              help=f'{pct_inm:.1f}% del capital total del inventario actual')

    st.caption(f'📊 Capital No Vigente: **{pct_inm:.1f}%** del inventario total actual')
    st.progress(min(pct_inm / 100, 1.0))
    st.write('---')

    # Gráficos: Subidas vs Bajadas
    g1, g2 = st.columns(2)
    with g1:
        st.subheader('📈 Top 10 Mayores Subidas')
        top_s = solo_aumentos.sort_values('Variacion_Unidades', ascending=False).head(10).copy()
        if not top_s.empty:
            top_s['Etiqueta'] = top_s['Variacion_Unidades'].apply(lambda x: f'+{int(x):,}'.replace(',', '.'))
            b = alt.Chart(top_s).mark_bar(color='#E15A97').encode(
                x=alt.X('Variacion_Unidades:Q', title='Unidades'),
                y=alt.Y('Nombre_Grafico:N', sort='-x', title='', axis=alt.Axis(labelLimit=300)))
            st.altair_chart((b + b.mark_text(align='left', dx=4, fontWeight='bold').encode(
                text='Etiqueta:N')).properties(height=350), use_container_width=True)
        else:
            st.info('No hubo subidas en este período.')

    with g2:
        st.subheader('📉 Top 10 Mayores Bajadas')
        top_b = solo_bajadas.sort_values('Variacion_Unidades', ascending=True).head(10).copy()
        if not top_b.empty:
            top_b['Etiqueta'] = top_b['Variacion_Unidades'].apply(lambda x: f'{int(x):,}'.replace(',', '.'))
            b2 = alt.Chart(top_b).mark_bar(color='#4A90E2').encode(
                x=alt.X('Variacion_Unidades:Q', title='Unidades'),
                y=alt.Y('Nombre_Grafico:N', sort='x', title='', axis=alt.Axis(labelLimit=300)))
            st.altair_chart((b2 + b2.mark_text(align='right', dx=-4, fontWeight='bold').encode(
                text='Etiqueta:N')).properties(height=350), use_container_width=True)
        else:
            st.success('✅ No hubo bajadas en este período.')

    # Nuevos y Desaparecidos
    st.write('---')
    cn, cd = st.columns(2)
    ALM_ANT = f'{COL_ALMACEN}_Ant'
    ALM_ACT = f'{COL_ALMACEN}_Act'

    with cn:
        st.subheader('🆕 Materiales Nuevos')
        if not mat_nuevos.empty:
            st.dataframe(mat_nuevos[['Material', 'Texto breve de material', 'LOTE',
                                     ALM_ACT, 'Libre utilización_Act', 'Valor libre util._Act']]
                         .rename(columns={ALM_ACT: 'Almacén',
                                          'Libre utilización_Act': 'Stock Actual',
                                          'Valor libre util._Act': 'Valor ($)'}),
                         use_container_width=True, hide_index=True)
        else:
            st.info('No hay materiales nuevos en este período.')

    with cd:
        st.subheader('🚫 Materiales Desaparecidos')
        if not desaparecidos.empty:
            st.dataframe(desaparecidos[['Material', 'Texto breve de material', 'LOTE',
                                        ALM_ANT, 'Libre utilización_Ant', 'Valor libre util._Ant']]
                         .rename(columns={ALM_ANT: 'Almacén',
                                          'Libre utilización_Ant': 'Stock Anterior',
                                          'Valor libre util._Ant': 'Valor Anterior ($)'}),
                         use_container_width=True, hide_index=True)
        else:
            st.success('✅ Ningún material desapareció en este período.')

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — REPORTES Y DESCARGAS
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    # Auditoría KPI
    st.subheader("🕵️‍♂️ Auditoría KPI: ¿Qué se sumó en 'No Vigente'?")
    with st.expander('Ver lista de materiales contabilizados'):
        st.write(f'**Total filas:** {len(df_kpi_nv):,}')
        st.write(f'**Filtros:** Estatus = NO VIGENTE | {COL_ALMACEN} válido (no FALSO, no vacío)')
        st.dataframe(df_kpi_nv[['Material', COL_ALMACEN, 'Estatus', 'Libre utilización', 'Valor libre util.']]
                     .rename(columns={'Libre utilización': 'Unidades', 'Valor libre util.': 'Valor ($)',
                                      COL_ALMACEN: 'Almacén'}),
                     use_container_width=True, hide_index=True)

    st.write('---')
    st.subheader('📈 Reporte de Aumentos')
    cols_aum = ['Material', 'Estado Material', ALM_ACT, 'LOTE',
                'Texto breve de material', 'Libre utilización_Ant',
                'Libre utilización_Act', 'Variacion_Unidades', '% Aumento']
    cols_aum = [c for c in cols_aum if c in solo_aumentos.columns]
    st.dataframe(
        solo_aumentos.sort_values('Variacion_Unidades', ascending=False)[cols_aum]
        .rename(columns={'Libre utilización_Ant': 'Stock Anterior',
                         'Libre utilización_Act': 'Stock Actual',
                         'Variacion_Unidades':    'Diferencia (+)',
                         ALM_ACT:                 'Almacén'}),
        use_container_width=True, hide_index=True,
        column_config={'Stock Anterior': st.column_config.NumberColumn(format='%d'),
                       'Stock Actual':   st.column_config.NumberColumn(format='%d'),
                       'Diferencia (+)': st.column_config.NumberColumn(format='%d')})

    st.write('---')
    st.subheader('📋 Plan de Acción')
    if not sobre_stock.empty:
        sobre_stock_s = sobre_stock.sort_values('Libre utilización_Act', ascending=False).copy()

        def recomendacion(row):
            if row['Variacion_Unidades'] > 500:   return '🔴 Alerta: Fuerte ingreso. Confirmar justificación.'
            elif row['Variacion_Unidades'] > 0:   return '🟡 Aumento de stock. Vigilar rotación.'
            elif row['Libre utilización_Act'] > 5000: return '🔵 Inmovilizado Alto: Evaluar Venta Ecommerce.'
            elif row['Libre utilización_Act'] > 1000: return '🟢 Inmovilizado Medio: Sugerir Solicitudes Turismo.'
            else:                                  return '⚪ Inmovilizado Bajo: Armar packs promocionales.'

        sobre_stock_s['Recomendación'] = sobre_stock_s.apply(recomendacion, axis=1)
        cols_plan = ['Material', 'Estado Material', ALM_ACT, 'LOTE',
                     'Texto breve de material', 'Libre utilización_Act',
                     'Variacion_Unidades', 'Valor libre util._Act', 'Recomendación']
        cols_plan = [c for c in cols_plan if c in sobre_stock_s.columns]
        st.dataframe(
            sobre_stock_s[cols_plan]
            .rename(columns={'Libre utilización_Act': 'Stock Actual',
                             'Variacion_Unidades':    'Variación (Unid.)',
                             'Valor libre util._Act': 'Valor Actual ($)',
                             ALM_ACT:                 'Almacén Actual'}),
            use_container_width=True, hide_index=True,
            column_config={'Valor Actual ($)':  st.column_config.NumberColumn(format='$ %d'),
                           'Stock Actual':      st.column_config.NumberColumn(format='%d'),
                           'Variación (Unid.)': st.column_config.NumberColumn(format='%d')})

    st.divider()
    st.subheader('📥 Descargar Reporte Completo')

    COLS_EXPORT = {
        'Material': 'Material', 'LOTE': 'Lote',
        'Texto breve de material': 'Descripción',
        ALM_ANT: 'Almacén Anterior', ALM_ACT: 'Almacén Actual',
        'Estatus_Ant': 'Estatus Anterior', 'Estatus_Act': 'Estatus Actual',
        'Libre utilización_Ant': 'Stock Semana Anterior',
        'Libre utilización_Act': 'Stock Semana Actual',
        'Variacion_Unidades': 'Diferencia Unidades',
        'Variacion_Valor':    'Diferencia Valor ($)',
        'Valor libre util._Ant': 'Valor Anterior ($)',
        'Valor libre util._Act': 'Valor Actual ($)',
        'Estado Material': 'Estado', '% Aumento': '% Cambio',
    }

    def prep(df_in, subset=None):
        cols = subset or list(COLS_EXPORT.keys())
        ok   = [c for c in cols if c in df_in.columns]
        return df_in[ok].rename(columns=COLS_EXPORT).copy()

    CTOT = ['Stock Semana Anterior', 'Stock Semana Actual', 'Diferencia Unidades']

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dar_formato_excel(writer,
            prep(solo_aumentos.sort_values('Variacion_Unidades', ascending=False),
                 ['Material', 'Texto breve de material', 'LOTE', ALM_ACT,
                  'Libre utilización_Ant', 'Libre utilización_Act', 'Variacion_Unidades',
                  '% Aumento', 'Valor libre util._Act', 'Estado Material']),
            '📈 Subidas de Stock', 'Diferencia Unidades', CTOT)

        dar_formato_excel(writer,
            prep(solo_bajadas.sort_values('Variacion_Unidades'),
                 ['Material', 'Texto breve de material', 'LOTE', ALM_ACT,
                  'Libre utilización_Ant', 'Libre utilización_Act', 'Variacion_Unidades',
                  'Valor libre util._Ant', 'Estado Material']),
            '📉 Bajadas de Stock', 'Diferencia Unidades', CTOT)

        dar_formato_excel(writer,
            prep(mat_nuevos, ['Material', 'Texto breve de material', 'LOTE', ALM_ACT,
                              'Libre utilización_Act', 'Valor libre util._Act']),
            '🆕 Materiales Nuevos', None, ['Stock Semana Actual', 'Valor Actual ($)'])

        dar_formato_excel(writer,
            prep(desaparecidos, ['Material', 'Texto breve de material', 'LOTE', ALM_ANT,
                                 'Libre utilización_Ant', 'Valor libre util._Ant']),
            '🚫 Desaparecidos', None, ['Stock Semana Anterior', 'Valor Anterior ($)'])

        dar_formato_excel(writer, prep(df_cruce.sort_values('Material')),
            '📋 Cruce Completo', 'Diferencia Unidades', CTOT)

        nv_exp = df_kpi_nv[[c for c in ['Material', 'Texto breve de material', 'LOTE',
                                         COL_ALMACEN, 'Estatus', 'Libre utilización', 'Valor libre util.']
                             if c in df_kpi_nv.columns]].rename(
            columns={'Libre utilización': 'Unidades', 'Valor libre util.': 'Valor ($)',
                     COL_ALMACEN: 'Almacén'}).copy()
        dar_formato_excel(writer, nv_exp, '⚠️ Alertas No Vigente', None, ['Unidades', 'Valor ($)'])

        # Hoja Evolución si hay 3+ semanas
        if len(comparaciones) > 1:
            rows_ev = []
            for c2 in comparaciones:
                cr = c2['cruce']
                rows_ev.append({
                    'Período':        c2['label'],
                    'Stock Anterior': int(cr['Libre utilización_Ant'].sum()),
                    'Stock Actual':   int(cr['Libre utilización_Act'].sum()),
                    'Var. Neta':      int(cr['Variacion_Unidades'].sum()),
                    'Subidas (mat.)': len(c2['subidas']),
                    'Bajadas (mat.)': len(c2['bajadas']),
                    'Nuevos':         len(c2['nuevos']),
                    'Desaparecidos':  len(c2['desaparecidos']),
                })
            dar_formato_excel(writer, pd.DataFrame(rows_ev), '📅 Evolución',
                              'Var. Neta', ['Stock Anterior', 'Stock Actual', 'Var. Neta'])

    st.download_button(
        label='📥 Descargar Reporte Completo',
        data=output.getvalue(),
        file_name='Reporte_Inventario.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type='primary')

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — TRAZABILIDAD HISTÓRICA (Google Sheets)
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header('Base de Datos Histórica (Google Sheets)')
    if conn is not None:
        with st.form('form_guardar_bd'):
            c_f, c_b = st.columns([1, 2])
            with c_f:
                fecha_reg = st.date_input('Fecha de esta foto de inventario:', date.today())
            with c_b:
                st.write(''); st.write('')
                guardar = st.form_submit_button("💾 Enviar 'Semana Actual' a Google Sheets")

            if guardar:
                with st.spinner('Conectando con Google Sheets...'):
                    try:
                        df_hist = conn.read(worksheet='Historial', usecols=list(range(6)), ttl=0).dropna(how='all')
                    except Exception:
                        df_hist = pd.DataFrame(columns=['Fecha_Registro', 'Material', 'LOTE',
                                                         'Texto_breve', 'Libre_utilizacion', 'Valor'])

                    cols_bd = ['Material', 'LOTE', 'Texto breve de material', 'Libre utilización', 'Valor libre util.']
                    falt_bd = [c for c in cols_bd if c not in df_act_actual.columns]
                    if falt_bd:
                        st.error(f'Faltan columnas para guardar: {falt_bd}')
                    else:
                        df_bd = df_act_actual[cols_bd].copy()
                        df_bd.rename(columns={'Texto breve de material': 'Texto_breve',
                                              'Libre utilización': 'Libre_utilizacion',
                                              'Valor libre util.': 'Valor'}, inplace=True)
                        df_bd.insert(0, 'Fecha_Registro', str(fecha_reg))
                        if not df_hist.empty:
                            df_hist['Fecha_Registro'] = df_hist['Fecha_Registro'].astype(str)
                            df_hist = df_hist[df_hist['Fecha_Registro'] != str(fecha_reg)]
                        conn.update(worksheet='Historial', data=pd.concat([df_hist, df_bd], ignore_index=True))
                        st.success(f'✅ Inventario del {fecha_reg} guardado correctamente.')

        st.divider()
        st.subheader('📈 Análisis de Tendencias Históricas')
        if st.button('🔄 Cargar Gráficos Históricos'):
            with st.spinner('Descargando historial desde Google...'):
                try:
                    df_hc = conn.read(worksheet='Historial', usecols=list(range(6)), ttl=0).dropna(how='all')
                    if not df_hc.empty:
                        df_hc['Fecha_Registro']    = pd.to_datetime(df_hc['Fecha_Registro'], errors='coerce')
                        df_hc['Libre_utilizacion'] = pd.to_numeric(df_hc['Libre_utilizacion'], errors='coerce')
                        mat_sel = st.selectbox('Selecciona un material:', sorted(df_hc['Texto_breve'].dropna().unique()))
                        datos_g = df_hc[df_hc['Texto_breve'] == mat_sel].copy()
                        if not datos_g.empty:
                            linea = alt.Chart(datos_g).mark_line(point=True, color='#FF5722', strokeWidth=3).encode(
                                x=alt.X('Fecha_Registro:T', title='Fecha'),
                                y=alt.Y('Libre_utilizacion:Q', title='Stock Total'),
                                color=alt.Color('LOTE:N', legend=alt.Legend(title='Lotes')))
                            st.altair_chart(linea.properties(height=400), use_container_width=True)
                    else:
                        st.info('Aún no has guardado ningún dato histórico.')
                except Exception as e:
                    st.error(f'No se pudo cargar el historial: {e}')
    else:
        st.warning('⚠️ La conexión a Google Sheets no está configurada. Los datos solo estarán disponibles en el reporte Excel.')

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — EVOLUCIÓN MULTI-SEMANA (solo si hay 3+ archivos)
# ══════════════════════════════════════════════════════════════════════════════
if tab4 is not None:
    with tab4:
        st.header('📅 Evolución del Inventario — Semana a Semana')

        # Tabla resumen
        rows = []
        for c2 in comparaciones:
            cr = c2['cruce']
            rows.append({
                'Período':           c2['label'],
                'Stock Anterior':    int(cr['Libre utilización_Ant'].sum()),
                'Stock Actual':      int(cr['Libre utilización_Act'].sum()),
                'Var. Neta (unid.)': int(cr['Variacion_Unidades'].sum()),
                '% Variación':       f"{cr['Variacion_Unidades'].sum() / max(cr['Libre utilización_Ant'].sum(), 1) * 100:.2f}%",
                'Subidas (mat.)':    len(c2['subidas']),
                'Bajadas (mat.)':    len(c2['bajadas']),
                'Nuevos':            len(c2['nuevos']),
                'Desaparecidos':     len(c2['desaparecidos']),
            })
        df_ev = pd.DataFrame(rows)
        st.dataframe(df_ev, use_container_width=True, hide_index=True,
                     column_config={
                         'Stock Anterior':    st.column_config.NumberColumn(format='%d'),
                         'Stock Actual':      st.column_config.NumberColumn(format='%d'),
                         'Var. Neta (unid.)': st.column_config.NumberColumn(format='%d'),
                     })

        st.write('---')

        # Gráfico: evolución del stock total
        stocks = [{'Semana': dfs_semanas[0]['nombre'].split('.')[0],
                   'Stock':  int(dfs_semanas[0]['df']['Libre utilización'].sum())}]
        for i, c2 in enumerate(comparaciones):
            stocks.append({'Semana': dfs_semanas[i + 1]['nombre'].split('.')[0],
                           'Stock':  int(c2['cruce']['Libre utilización_Act'].sum())})

        df_stocks = pd.DataFrame(stocks)
        g1e, g2e  = st.columns(2)
        with g1e:
            st.subheader('📦 Evolución del Stock Total')
            chart_stock = alt.Chart(df_stocks).mark_line(point=True, strokeWidth=3, color='#4A235A').encode(
                x=alt.X('Semana:N', sort=None, title='Semana'),
                y=alt.Y('Stock:Q', title='Unidades totales'))
            st.altair_chart(chart_stock.properties(height=300), use_container_width=True)

        # Gráfico: materiales con mayor variación acumulada (primera vs última semana)
        with g2e:
            st.subheader('🔝 Top 10 Variación Acumulada (primera vs última semana)')
            try:
                res_total = comparar_par(dfs_semanas[0]['df'], dfs_semanas[-1]['df'])
                top_acum  = res_total['cruce'].copy()
                top_acum  = top_acum.reindex(
                    top_acum['Variacion_Unidades'].abs().nlargest(10).index)
                top_acum['Color']    = top_acum['Variacion_Unidades'].apply(
                    lambda x: '📈 Subió' if x > 0 else '📉 Bajó')
                top_acum['Etiqueta'] = top_acum['Variacion_Unidades'].apply(
                    lambda x: f'+{int(x):,}'.replace(',', '.') if x > 0 else f'{int(x):,}'.replace(',', '.'))
                chart_acum = alt.Chart(top_acum).mark_bar().encode(
                    x=alt.X('Variacion_Unidades:Q', title='Variación acumulada'),
                    y=alt.Y('Texto breve de material:N', sort='-x', title='',
                            axis=alt.Axis(labelLimit=250)),
                    color=alt.Color('Color:N',
                                    scale=alt.Scale(domain=['📈 Subió', '📉 Bajó'],
                                                    range=['#E15A97', '#4A90E2'])))
                st.altair_chart(chart_acum.properties(height=300), use_container_width=True)
            except Exception:
                st.info('No se pudo calcular variación acumulada.')
