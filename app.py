import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re
import io
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────
REPO_NAME = "alejandraho013/registro-salida-residuos"

GESTORES_DATA = {
    "CORPOGESTAR": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
    "Recicla Oriente": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
    "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
    "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"]),
    "Otro": []
}

COLORES_EMPRESA = {
    "CORPOGESTAR": "#2196F3",
    "Recicla Oriente": "#4CAF50",
    "Quimetales NO Peligrosos": "#FF9800",
    "Quimetales Peligrosos": "#F44336",
}
COLOR_DEFAULT = "#9C27B0"

try:
    TOKEN = st.secrets["TOKEN"]
    REPO_NAME = st.secrets.get("REPO_NAME", REPO_NAME)
except Exception:
    st.error("⚠️ Configura el TOKEN en 'Advanced Settings' > Secrets.")
    st.stop()

# ─────────────────────────────────────────────────────────────
# PÁGINA
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TINTATEX · Gestión de Residuos",
    page_icon="♻️",
    layout="wide",
)

# CSS personalizado
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a237e 0%, #283593 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; }
    .main-header p  { margin: 0.3rem 0 0; opacity: 0.8; font-size: 0.95rem; }

    .kpi-card {
        background: white;
        border-radius: 10px;
        padding: 1.2rem;
        border-left: 4px solid #1a237e;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    .kpi-value { font-size: 2rem; font-weight: 700; color: #1a237e; }
    .kpi-label { font-size: 0.85rem; color: #666; margin-top: 0.2rem; }

    .section-divider {
        border: none;
        border-top: 2px solid #e8eaf6;
        margin: 1.5rem 0;
    }
    div[data-testid="stExpander"] { border-radius: 10px; }
    .stButton > button { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>♻️ TINTATEX · Gestión de Residuos</h1>
    <p>Sistema de registro y trazabilidad de salida de residuos</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
if "lista_temporal" not in st.session_state:
    st.session_state.lista_temporal = []
if "envio_exitoso" not in st.session_state:
    st.session_state.envio_exitoso = False

# ─────────────────────────────────────────────────────────────
# FUNCIONES DE DATOS
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar_datos():
    try:
        g = Github(TOKEN)
        repo = g.get_repo(REPO_NAME)
        xlsx_file = repo.get_contents("database.xlsx")
        df = pd.read_excel(
            io.BytesIO(xlsx_file.decoded_content),
            sheet_name="MASTER",
            engine="openpyxl"
        )
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        df["mes_num"] = df["fecha"].dt.to_period("M")
        return df
    except Exception:
        return pd.DataFrame()


def generar_excel_completo(df_master: pd.DataFrame) -> bytes:
    """
    Genera un Excel con 4 hojas bien formateadas:
      1. MASTER         – todos los datos originales
      2. Detalle        – tabla filtrada actual
      3. Resumen Empresas – pivot empresa × residuo
      4. Tendencia Mensual – kg por mes y empresa
    """
    output = io.BytesIO()

    # ── Preparar datos ──────────────────────────────────────
    df_m = df_master.copy()
    df_m["fecha"] = df_m["fecha"].astype(str)

    resumen_empresa = (
        df_m.groupby("empresa")
        .agg(
            Registros=("peso_kg", "count"),
            Peso_Total_kg=("peso_kg", "sum"),
            Peso_Promedio_kg=("peso_kg", "mean"),
            Peso_Maximo_kg=("peso_kg", "max"),
        )
        .round(1)
        .sort_values("Peso_Total_kg", ascending=False)
        .reset_index()
    )

    pivot_residuo = (
        df_m.pivot_table(
            index="empresa",
            columns="tipo_residuo",
            values="peso_kg",
            aggfunc="sum",
            fill_value=0,
        )
        .round(1)
        .reset_index()
    )

    tendencia = (
        df_m.groupby(["mes", "empresa"])["peso_kg"]
        .sum()
        .round(1)
        .reset_index()
        .sort_values("mes")
    )

    # ── Escribir hojas ──────────────────────────────────────
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_m.to_excel(writer, sheet_name="MASTER", index=False)
        resumen_empresa.to_excel(writer, sheet_name="Resumen Empresas", index=False)
        pivot_residuo.to_excel(writer, sheet_name="Pivot Residuos", index=False)
        tendencia.to_excel(writer, sheet_name="Tendencia Mensual", index=False)

        wb = writer.book

        # ── Estilos comunes ──────────────────────────────────
        header_fill   = PatternFill("solid", fgColor="1A237E")
        header_font   = Font(color="FFFFFF", bold=True, size=11)
        alt_fill      = PatternFill("solid", fgColor="E8EAF6")
        total_fill    = PatternFill("solid", fgColor="C5CAE9")
        total_font    = Font(bold=True, size=11)
        center_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border   = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        def _fmt_sheet(ws, col_widths=None):
            """Aplica formato básico a una hoja."""
            for cell in ws[1]:
                cell.fill      = header_fill
                cell.font      = header_font
                cell.alignment = center_align
                cell.border    = thin_border

            for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
                fill = alt_fill if row_idx % 2 == 0 else None
                for cell in row:
                    if fill:
                        cell.fill = fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            if col_widths:
                for col_letter, width in col_widths.items():
                    ws.column_dimensions[col_letter].width = width
            else:
                for col in ws.columns:
                    max_len = max((len(str(c.value or "")) for c in col), default=10)
                    ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 35)

        # ── MASTER ──────────────────────────────────────────
        ws_master = wb["MASTER"]
        _fmt_sheet(ws_master)
        ws_master.sheet_tab_color = "1A237E"

        # ── Resumen Empresas ────────────────────────────────
        ws_res = wb["Resumen Empresas"]
        _fmt_sheet(ws_res)
        ws_res.sheet_tab_color = "4CAF50"

        # Fila de totales
        last_row = ws_res.max_row + 1
        ws_res.cell(last_row, 1, "TOTAL").font = total_font
        ws_res.cell(last_row, 1).fill = total_fill
        for col in range(2, ws_res.max_column + 1):
            vals = [
                ws_res.cell(r, col).value
                for r in range(2, ws_res.max_row)
                if isinstance(ws_res.cell(r, col).value, (int, float))
            ]
            cell = ws_res.cell(last_row, col)
            cell.value = round(sum(vals), 1) if vals else ""
            cell.font  = total_font
            cell.fill  = total_fill
            cell.border = thin_border
            cell.alignment = center_align

        # Gráfico de barras incrustado
        chart = BarChart()
        chart.type  = "col"
        chart.title = "Peso Total por Empresa (kg)"
        chart.style = 10
        chart.y_axis.title = "kg"
        chart.x_axis.title = "Empresa"
        n_empresas = len(resumen_empresa)
        data_ref  = Reference(ws_res, min_col=3, min_row=1, max_row=n_empresas + 1)
        cats_ref  = Reference(ws_res, min_col=1, min_row=2, max_row=n_empresas + 1)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.shape = 4
        chart.width  = 20
        chart.height = 14
        ws_res.add_chart(chart, f"G2")

        # ── Pivot Residuos ──────────────────────────────────
        ws_piv = wb["Pivot Residuos"]
        _fmt_sheet(ws_piv)
        ws_piv.sheet_tab_color = "FF9800"

        # ── Tendencia Mensual ───────────────────────────────
        ws_tend = wb["Tendencia Mensual"]
        _fmt_sheet(ws_tend)
        ws_tend.sheet_tab_color = "F44336"

    return output.getvalue()


# ─────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📝 Registro de Salida", "📊 Reportes y Estadísticas"])


# ═════════════════════════════════════════════════════════════
# TAB 1 · REGISTRO
# ═════════════════════════════════════════════════════════════
with tab1:

    if st.session_state.envio_exitoso:
        st.success("✅ ¡Registro guardado correctamente en la base de datos!")
        st.balloons()
        if st.button("🔄 Iniciar Nuevo Registro", type="primary", use_container_width=True):
            st.session_state.envio_exitoso = False
            st.session_state.lista_temporal = []
            cargar_datos.clear()
            st.rerun()

    else:
        # ── 1. DATOS DEL VEHÍCULO ────────────────────────────
        with st.expander("🚛 **1. Datos del Vehículo y Gestor**", expanded=True):
            c1, c2 = st.columns(2)

            fecha = c1.date_input("📅 Fecha de salida", datetime.now())
            empresa_sel = c1.selectbox("🏭 Empresa Gestora", options=list(GESTORES_DATA.keys()))
            empresa_final = empresa_sel
            if empresa_sel == "Otro":
                empresa_final = c1.text_input("✏️ Nombre del Gestor").upper().strip()

            conductor = c2.text_input("👤 Conductor")
            placa_raw = c2.text_input("🚗 Placa (Ej: ABC123)").upper().strip()
            placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa_raw)) if placa_raw else False

            if placa_raw and not placa_valida:
                c2.warning("⚠️ Formato incorrecto. Ejemplo: **ABC123**")
            elif placa_raw and placa_valida:
                c2.success(f"✅ Placa válida: **{placa_raw}**")

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

        # ── 2. PESAJES ───────────────────────────────────────
        st.subheader("⚖️ 2. Detalle de Pesajes")

        opciones_res = GESTORES_DATA.get(empresa_sel, []).copy()
        opciones_res.append("Otro")

        col_res, col_pes, col_btn = st.columns([3, 2, 1])
        tipo_sel     = col_res.selectbox("Tipo de Residuo", options=opciones_res)
        residuo_final = (col_res.text_input("Especifique residuo") if tipo_sel == "Otro" else tipo_sel)
        peso = col_pes.number_input("Peso (kg)", min_value=0.0, max_value=50000.0, step=0.1, format="%.1f")

        col_btn.markdown("<br>", unsafe_allow_html=True)
        if col_btn.button("➕ Agregar", use_container_width=True):
            if not placa_valida:
                st.error("❌ Corrija el formato de la placa antes de agregar pesajes.")
            elif peso <= 0:
                st.error("❌ El peso debe ser mayor a 0 kg.")
            elif not residuo_final:
                st.error("❌ Especifique el tipo de residuo.")
            else:
                st.session_state.lista_temporal.append(
                    {"tipo_residuo": residuo_final, "peso_kg": peso}
                )
                st.toast(f"✅ {residuo_final} — {peso:,.1f} kg")

        # Tabla de pesajes acumulados
        if st.session_state.lista_temporal:
            df_temp = pd.DataFrame(st.session_state.lista_temporal)
            total_kg = df_temp["peso_kg"].sum()

            m1, m2, m3 = st.columns(3)
            m1.metric("⚖️ Peso Total", f"{total_kg:,.1f} kg")
            m2.metric("📦 Pesajes", len(st.session_state.lista_temporal))
            m3.metric("📊 Promedio", f"{df_temp['peso_kg'].mean():,.1f} kg")

            with st.expander("🔍 Pesajes ingresados", expanded=True):
                # Tabla editable con índices
                df_display = df_temp.copy()
                df_display.index = range(1, len(df_display) + 1)
                df_display.columns = ["Tipo de Residuo", "Peso (kg)"]
                st.dataframe(df_display, use_container_width=True)

                col_del, col_limpiar = st.columns([2, 1])
                idx_del = col_del.number_input(
                    "Eliminar fila #", min_value=1,
                    max_value=len(st.session_state.lista_temporal), step=1
                )
                if col_del.button("🗑️ Eliminar fila seleccionada", use_container_width=True):
                    st.session_state.lista_temporal.pop(int(idx_del) - 1)
                    st.rerun()
                if col_limpiar.button("🧹 Limpiar todo", type="secondary", use_container_width=True):
                    st.session_state.lista_temporal = []
                    st.rerun()

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

        # ── 3. EVIDENCIAS Y ENVÍO ────────────────────────────
        st.subheader("📎 3. Evidencias y Observaciones")
        f1, f2 = st.columns(2)
        foto_memo   = f1.file_uploader("🧾 Foto del Memo (Opcional)", type=["jpg", "png", "jpeg"])
        foto_camion = f2.file_uploader("🚛 Foto del Camión (Opcional)", type=["jpg", "png", "jpeg"])
        novedades   = st.text_area("📝 Observaciones / Novedades", placeholder="Ingrese cualquier novedad relevante...")

        st.markdown("")
        if st.button("📤 ENVIAR REGISTRO COMPLETO", type="primary", use_container_width=True):
            # Validaciones
            errores = []
            if not st.session_state.lista_temporal:
                errores.append("Agregue al menos un pesaje.")
            if not placa_valida:
                errores.append("Corrija el formato de la placa (Ej: ABC123).")
            if empresa_sel == "Otro" and not empresa_final:
                errores.append("Especifique el nombre del gestor.")
            if not conductor.strip():
                errores.append("Ingrese el nombre del conductor.")

            if errores:
                for e in errores:
                    st.error(f"❌ {e}")
            else:
                with st.spinner("💾 Guardando registro en la nube..."):
                    try:
                        g    = Github(TOKEN)
                        repo = g.get_repo(REPO_NAME)
                        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
                        mes_actual = fecha.strftime("%B_%Y")

                        u_memo, u_camion = "Sin foto", "Sin foto"
                        if foto_memo:
                            p_memo = f"fotos/MEMO_{ts}.jpg"
                            repo.create_file(p_memo, f"Memo {ts}", foto_memo.getvalue())
                            u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"
                        if foto_camion:
                            p_camion = f"fotos/CAMION_{ts}.jpg"
                            repo.create_file(p_camion, f"Camion {ts}", foto_camion.getvalue())
                            u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                        xlsx_file       = repo.get_contents("database.xlsx")
                        diccionario_hojas = pd.read_excel(
                            io.BytesIO(xlsx_file.decoded_content),
                            sheet_name=None, engine="openpyxl"
                        )

                        nov_final   = novedades.strip() or "Sin observaciones"
                        nuevas_filas = [
                            {
                                "fecha":        str(fecha),
                                "mes":          mes_actual,
                                "empresa":      empresa_final,
                                "conductor":    conductor.strip(),
                                "placa":        placa_raw,
                                "tipo_residuo": x["tipo_residuo"],
                                "peso_kg":      x["peso_kg"],
                                "novedades":    nov_final,
                                "url_memo":     u_memo,
                                "url_camion":   u_camion,
                            }
                            for x in st.session_state.lista_temporal
                        ]
                        df_nuevos = pd.DataFrame(nuevas_filas)

                        # Actualizar MASTER
                        if "MASTER" in diccionario_hojas:
                            diccionario_hojas["MASTER"] = pd.concat(
                                [diccionario_hojas["MASTER"], df_nuevos], ignore_index=True
                            )
                        else:
                            diccionario_hojas["MASTER"] = df_nuevos

                        # Hoja por empresa
                        nombre_hoja = re.sub(r'[\\/*?:\[\]]', "", empresa_final)[:30].strip().upper() or "OTROS"
                        diccionario_hojas[nombre_hoja] = pd.concat(
                            [diccionario_hojas.get(nombre_hoja, pd.DataFrame()), df_nuevos],
                            ignore_index=True
                        )

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            for hoja, df_h in diccionario_hojas.items():
                                df_h.to_excel(writer, sheet_name=hoja, index=False)

                        repo.update_file(
                            "database.xlsx",
                            f"Registro {empresa_final} {placa_raw} {ts}",
                            output.getvalue(),
                            xlsx_file.sha,
                        )

                        st.session_state.envio_exitoso = True
                        st.rerun()

                    except Exception as e:
                        st.error("❌ Error al guardar el registro.")
                        st.exception(e)


# ═════════════════════════════════════════════════════════════
# TAB 2 · REPORTES
# ═════════════════════════════════════════════════════════════
with tab2:

    col_ref, _ = st.columns([1, 5])
    if col_ref.button("🔄 Actualizar datos"):
        cargar_datos.clear()
        st.rerun()

    df_master = cargar_datos()

    if df_master.empty:
        st.info("ℹ️ No hay datos disponibles aún. Registra la primera salida.")
        st.stop()

    # ── FILTROS ──────────────────────────────────────────────
    with st.expander("🔍 **Filtros**", expanded=True):
        fc1, fc2, fc3, fc4 = st.columns(4)
        empresa_filtro  = fc1.selectbox("Empresa",         ["Todas"] + sorted(df_master["empresa"].dropna().unique().tolist()))
        residuo_filtro  = fc2.selectbox("Tipo de Residuo", ["Todos"] + sorted(df_master["tipo_residuo"].dropna().unique().tolist()))
        mes_filtro      = fc3.selectbox("Mes",             ["Todos"] + sorted(df_master["mes"].dropna().unique().tolist()))
        conductor_filtro = fc4.selectbox("Conductor",       ["Todos"] + sorted(df_master["conductor"].dropna().unique().tolist()))

    df_f = df_master.copy()
    if empresa_filtro  != "Todas": df_f = df_f[df_f["empresa"]      == empresa_filtro]
    if residuo_filtro  != "Todos": df_f = df_f[df_f["tipo_residuo"] == residuo_filtro]
    if mes_filtro      != "Todos": df_f = df_f[df_f["mes"]          == mes_filtro]
    if conductor_filtro != "Todos": df_f = df_f[df_f["conductor"]   == conductor_filtro]

    # ── KPIs ─────────────────────────────────────────────────
    k1, k2, k3, k4, k5 = st.columns(5)
    total_kg   = df_f["peso_kg"].sum()
    n_registros = len(df_f)
    n_empresas = df_f["empresa"].nunique()
    promedio   = df_f["peso_kg"].mean() if n_registros else 0
    n_residuos = df_f["tipo_residuo"].nunique()

    for col, val, label in zip(
        [k1, k2, k3, k4, k5],
        [f"{total_kg:,.1f} kg", f"{n_registros:,}", f"{n_empresas}", f"{promedio:,.1f} kg", f"{n_residuos}"],
        ["Peso Total", "Registros", "Empresas Activas", "Promedio / Registro", "Tipos de Residuo"],
    ):
        col.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-value">{val}</div>
            <div class="kpi-label">{label}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ── GRÁFICAS ─────────────────────────────────────────────
    st.subheader("📈 Visualizaciones")

    g1, g2 = st.columns(2)

    # Barras por empresa
    with g1:
        res_emp = (
            df_f.groupby("empresa")["peso_kg"]
            .sum().round(1)
            .sort_values(ascending=False)
            .reset_index()
        )
        colores = [COLORES_EMPRESA.get(e, COLOR_DEFAULT) for e in res_emp["empresa"]]
        fig_emp = px.bar(
            res_emp, x="empresa", y="peso_kg",
            title="⚖️ Peso Total por Empresa (kg)",
            labels={"empresa": "Empresa", "peso_kg": "Peso (kg)"},
            color="empresa",
            color_discrete_sequence=colores,
            text_auto=".1f",
        )
        fig_emp.update_layout(showlegend=False, plot_bgcolor="white", title_font_size=15)
        fig_emp.update_traces(textposition="outside")
        st.plotly_chart(fig_emp, use_container_width=True)

    # Donut por tipo de residuo
    with g2:
        res_res = (
            df_f.groupby("tipo_residuo")["peso_kg"]
            .sum().round(1)
            .sort_values(ascending=False)
            .reset_index()
        )
        fig_donut = px.pie(
            res_res, names="tipo_residuo", values="peso_kg",
            title="🗂️ Distribución por Tipo de Residuo",
            hole=0.45,
            color_discrete_sequence=px.colors.qualitative.Set3,
        )
        fig_donut.update_traces(textposition="inside", textinfo="percent+label")
        fig_donut.update_layout(title_font_size=15, showlegend=True)
        st.plotly_chart(fig_donut, use_container_width=True)

    # Tendencia temporal
    df_tend = (
        df_f.groupby(["mes", "empresa"])["peso_kg"]
        .sum().round(1)
        .reset_index()
        .sort_values("mes")
    )
    if len(df_tend["mes"].unique()) > 1:
        fig_tend = px.line(
            df_tend, x="mes", y="peso_kg", color="empresa",
            title="📅 Tendencia Mensual por Empresa (kg)",
            labels={"mes": "Mes", "peso_kg": "Peso (kg)", "empresa": "Empresa"},
            markers=True,
            color_discrete_map=COLORES_EMPRESA,
        )
        fig_tend.update_layout(plot_bgcolor="white", title_font_size=15)
        st.plotly_chart(fig_tend, use_container_width=True)

    # Heatmap empresa × residuo
    pivot_heat = df_f.pivot_table(
        index="empresa", columns="tipo_residuo",
        values="peso_kg", aggfunc="sum", fill_value=0
    ).round(1)
    if not pivot_heat.empty:
        fig_heat = px.imshow(
            pivot_heat,
            title="🗺️ Mapa de Calor: Empresa × Tipo de Residuo (kg)",
            color_continuous_scale="Blues",
            aspect="auto",
            text_auto=".0f",
        )
        fig_heat.update_layout(title_font_size=15)
        st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ── TABLAS ───────────────────────────────────────────────
    st.subheader("📋 Tablas Detalladas")

    t1, t2, t3 = st.tabs(["Datos Completos", "Resumen por Empresa", "Pivot Residuos"])

    with t1:
        cols_display = ["fecha", "empresa", "conductor", "placa", "tipo_residuo", "peso_kg", "mes", "novedades"]
        df_tabla = df_f[[c for c in cols_display if c in df_f.columns]].sort_values("fecha", ascending=False)
        st.dataframe(df_tabla, use_container_width=True, hide_index=True,
                     column_config={
                         "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                         "peso_kg": st.column_config.NumberColumn("Peso (kg)", format="%.1f"),
                     })

    with t2:
        resumen = (
            df_f.groupby("empresa")
            .agg(
                Registros=("peso_kg", "count"),
                Peso_Total_kg=("peso_kg", "sum"),
                Peso_Promedio_kg=("peso_kg", "mean"),
                Peso_Maximo_kg=("peso_kg", "max"),
            )
            .round(1)
            .sort_values("Peso_Total_kg", ascending=False)
            .reset_index()
        )
        st.dataframe(resumen, use_container_width=True, hide_index=True,
                     column_config={
                         "Peso_Total_kg":    st.column_config.NumberColumn("Peso Total (kg)",   format="%.1f"),
                         "Peso_Promedio_kg": st.column_config.NumberColumn("Promedio (kg)",     format="%.1f"),
                         "Peso_Maximo_kg":   st.column_config.NumberColumn("Máximo (kg)",       format="%.1f"),
                     })

    with t3:
        pivot_show = df_f.pivot_table(
            index="empresa", columns="tipo_residuo",
            values="peso_kg", aggfunc="sum", fill_value=0
        ).round(1).reset_index()
        st.dataframe(pivot_show, use_container_width=True, hide_index=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ── DESCARGA EXCEL ───────────────────────────────────────
    st.subheader("⬇️ Exportar Datos")
    excel_bytes = generar_excel_completo(df_f)
    st.download_button(
        label="📥 Descargar Excel Completo (4 hojas + gráfico incrustado)",
        data=excel_bytes,
        file_name=f"Reporte_TINTATEX_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
