import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re
import io
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────
REPO_NAME = "alejandraho013/registro-salida-residuos"

GESTORES_DATA = {
    "CORPOGESTAR":             sorted(["Cartón limpio","Cartón sucio","Papel de archivo","Pasta","PET limpio","PET sucio","Plástico","Retal de tela","Tubo plega"]),
    "Recicla Oriente":         sorted(["Cartón limpio","Cartón sucio","Papel de archivo","Pasta","PET limpio","PET sucio","Plástico","Retal de tela","Tubo plega"]),
    "Quimetales NO Peligrosos":sorted(["Algodón","Retal de tela","Tubo plega"]),
    "Quimetales Peligrosos":   sorted(["RAEE","Residuos laboratorio","Tela sucia"]),
    "Otro": [],
}

COLORES_EMPRESA = {
    "CORPOGESTAR":              "#2196F3",
    "Recicla Oriente":          "#4CAF50",
    "Quimetales NO Peligrosos": "#FF9800",
    "Quimetales Peligrosos":    "#F44336",
}
COLOR_DEFAULT = "#9C27B0"

# ── Secrets ──────────────────────────────────────────────────
try:
    TOKEN     = st.secrets["TOKEN"]
    REPO_NAME = st.secrets.get("REPO_NAME", REPO_NAME)
except Exception:
    st.error("⚠️ Configura el TOKEN en 'Advanced Settings' > Secrets.")
    st.stop()

# ── Decidir fuente de datos ───────────────────────────────────
# Si existen credenciales de OneDrive en secrets se usa OneDrive;
# si no, se cae a GitHub (modo legacy).
USE_ONEDRIVE = all(
    k in st.secrets for k in ("AZURE_TENANT_ID","AZURE_CLIENT_ID",
                               "AZURE_CLIENT_SECRET","ONEDRIVE_FILE_ID")
)
if USE_ONEDRIVE:
    from onedrive import (
        cargar_datos_onedrive,
        append_filas_onedrive,
        subir_foto_onedrive,
        descargar_excel_onedrive,
    )

# ─────────────────────────────────────────────────────────────
# PÁGINA
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="TINTATEX · Gestión de Residuos", layout="wide")

st.markdown("""
<style>
.main-header {
    background: linear-gradient(135deg,#1a237e 0%,#283593 100%);
    padding:1.5rem 2rem; border-radius:12px; margin-bottom:1.5rem; color:white;
}
.main-header h1 { margin:0; font-size:1.8rem; }
.main-header p  { margin:.3rem 0 0; opacity:.8; font-size:.95rem; }
.kpi-card {
    background:white; border-radius:10px; padding:1.2rem;
    border-left:4px solid #1a237e; box-shadow:0 2px 8px rgba(0,0,0,.08);
}
.kpi-value { font-size:2rem; font-weight:700; color:#1a237e; }
.kpi-label { font-size:.85rem; color:#666; margin-top:.2rem; }
.section-divider { border:none; border-top:2px solid #e8eaf6; margin:1.5rem 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>🏭 TINTATEX · Gestión de Residuos</h1>
    <p>Sistema de registro y trazabilidad de salida de residuos</p>
</div>
""", unsafe_allow_html=True)

if USE_ONEDRIVE:
    st.caption("📂 Fuente de datos: **OneDrive**")
else:
    st.caption("📂 Fuente de datos: **GitHub** (legacy)")

# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
if "lista_temporal"  not in st.session_state: st.session_state.lista_temporal  = []
if "envio_exitoso"   not in st.session_state: st.session_state.envio_exitoso   = False

# ─────────────────────────────────────────────────────────────
# RECURSOS CACHEADOS (GitHub legacy)
# ─────────────────────────────────────────────────────────────

@st.cache_resource          # conexión reutilizada entre reruns
def _github_repo():
    return Github(TOKEN).get_repo(REPO_NAME)


@st.cache_data(ttl=600)     # FIX: 600 s en vez de 300; invalidar al guardar
def _cargar_datos_github() -> pd.DataFrame:
    try:
        repo  = _github_repo()
        f     = repo.get_contents("database.xlsx")
        df    = pd.read_excel(
            io.BytesIO(f.decoded_content),
            sheet_name="MASTER",   # FIX: solo MASTER, no sheet_name=None
            engine="openpyxl"
        )
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        return df
    except Exception:
        return pd.DataFrame()


def cargar_datos() -> pd.DataFrame:
    """Capa de abstracción: devuelve datos de OneDrive o GitHub según config."""
    if USE_ONEDRIVE:
        return cargar_datos_onedrive()
    return _cargar_datos_github()


def invalidar_cache():
    if USE_ONEDRIVE:
        cargar_datos_onedrive.clear()
    else:
        _cargar_datos_github.clear()

# ─────────────────────────────────────────────────────────────
# GUARDAR REGISTRO
# ─────────────────────────────────────────────────────────────

def _guardar_onedrive(nuevas_filas, foto_memo, foto_camion, ts):
    """Guarda en OneDrive: fotos + append de filas. Sin reescribir el Excel."""
    u_memo, u_camion = "Sin foto", "Sin foto"
    if foto_memo:
        u_memo   = subir_foto_onedrive(f"MEMO_{ts}.jpg",   foto_memo.getvalue())
    if foto_camion:
        u_camion = subir_foto_onedrive(f"CAMION_{ts}.jpg", foto_camion.getvalue())

    for row in nuevas_filas:
        row["url_memo"]   = u_memo
        row["url_camion"] = u_camion

    append_filas_onedrive(nuevas_filas)     # solo escribe filas nuevas


def _guardar_github(nuevas_filas, foto_memo, foto_camion, ts, fecha, empresa_final, placa_raw):
    """
    Guarda en GitHub.
    FIX principal: lee el Excel UNA sola vez y solo la hoja MASTER,
    luego hace update_file con el mínimo de hojas necesarias.
    """
    repo = _github_repo()
    u_memo, u_camion = "Sin foto", "Sin foto"

    if foto_memo:
        p = f"fotos/MEMO_{ts}.jpg"
        repo.create_file(p, f"Memo {ts}", foto_memo.getvalue())
        u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p}"
    if foto_camion:
        p = f"fotos/CAMION_{ts}.jpg"
        repo.create_file(p, f"Camion {ts}", foto_camion.getvalue())
        u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p}"

    for row in nuevas_filas:
        row["url_memo"]   = u_memo
        row["url_camion"] = u_camion

    # FIX: leer solo MASTER (no todas las hojas)
    xlsx_file = repo.get_contents("database.xlsx")
    df_master = pd.read_excel(
        io.BytesIO(xlsx_file.decoded_content),
        sheet_name="MASTER",
        engine="openpyxl"
    )

    df_nuevos = pd.DataFrame(nuevas_filas)
    df_master = pd.concat([df_master, df_nuevos], ignore_index=True)

    nombre_hoja = re.sub(r'[\\/*?:\[\]]', "", empresa_final)[:30].strip().upper() or "OTROS"

    # Leer hoja de empresa (lazy: solo si existe)
    try:
        df_emp = pd.read_excel(
            io.BytesIO(xlsx_file.decoded_content),
            sheet_name=nombre_hoja,
            engine="openpyxl"
        )
        df_emp = pd.concat([df_emp, df_nuevos], ignore_index=True)
    except Exception:
        df_emp = df_nuevos.copy()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="MASTER",      index=False)
        df_emp.to_excel(   writer, sheet_name=nombre_hoja,   index=False)

    repo.update_file(
        "database.xlsx",
        f"Registro {empresa_final} {placa_raw} {ts}",
        output.getvalue(),
        xlsx_file.sha,
    )


# ─────────────────────────────────────────────────────────────
# EXCEL DE DESCARGA
# ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _generar_excel_completo(df_hash: int, df_master: pd.DataFrame) -> bytes:
    """
    FIX: cacheado por hash del DataFrame.
    No regenera si los datos no cambiaron entre clics.
    """
    output = io.BytesIO()
    df_m = df_master.copy()
    df_m["fecha"] = df_m["fecha"].astype(str)

    resumen_empresa = (
        df_m.groupby("empresa")
        .agg(Registros=("peso_kg","count"), Peso_Total_kg=("peso_kg","sum"),
             Peso_Promedio_kg=("peso_kg","mean"), Peso_Maximo_kg=("peso_kg","max"))
        .round(1).sort_values("Peso_Total_kg", ascending=False).reset_index()
    )
    pivot_residuo = (
        df_m.pivot_table(index="empresa", columns="tipo_residuo",
                         values="peso_kg", aggfunc="sum", fill_value=0)
        .round(1).reset_index()
    )
    tendencia = (
        df_m.groupby(["mes","empresa"])["peso_kg"].sum().round(1)
        .reset_index().sort_values("mes")
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_m.to_excel(         writer, sheet_name="MASTER",           index=False)
        resumen_empresa.to_excel(writer, sheet_name="Resumen Empresas", index=False)
        pivot_residuo.to_excel( writer, sheet_name="Pivot Residuos",    index=False)
        tendencia.to_excel(     writer, sheet_name="Tendencia Mensual", index=False)

        wb = writer.book
        hf = PatternFill("solid", fgColor="1A237E")
        ff = Font(color="FFFFFF", bold=True, size=11)
        ca = Alignment(horizontal="center", vertical="center", wrap_text=True)
        tb = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"),  bottom=Side(style="thin"))
        af = PatternFill("solid", fgColor="E8EAF6")

        def _fmt(ws):
            for c in ws[1]:
                c.fill=hf; c.font=ff; c.alignment=ca; c.border=tb
            for i, row in enumerate(ws.iter_rows(min_row=2), 2):
                for c in row:
                    if i%2==0: c.fill=af
                    c.alignment=Alignment(horizontal="center",vertical="center")
                    c.border=tb
            ws.freeze_panes="A2"
            ws.auto_filter.ref=ws.dimensions
            for col in ws.columns:
                mx=max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width=min(mx+4,35)

        for color, sheet in [("1A237E","MASTER"),("4CAF50","Resumen Empresas"),
                              ("FF9800","Pivot Residuos"),("F44336","Tendencia Mensual")]:
            ws = wb[sheet]
            _fmt(ws)
            ws.sheet_tab_color = color

        # Gráfico incrustado en Resumen Empresas
        ws_res = wb["Resumen Empresas"]
        n = len(resumen_empresa)
        chart = BarChart()
        chart.type="col"; chart.title="Peso Total por Empresa (kg)"
        chart.style=10; chart.y_axis.title="kg"; chart.width=20; chart.height=14
        chart.add_data(Reference(ws_res,min_col=3,min_row=1,max_row=n+1),titles_from_data=True)
        chart.set_categories(Reference(ws_res,min_col=1,min_row=2,max_row=n+1))
        ws_res.add_chart(chart,"G2")

    return output.getvalue()


def generar_excel_para_descarga(df: pd.DataFrame) -> bytes:
    if USE_ONEDRIVE:
        # Si es OneDrive devuelve el archivo real tal cual está en la nube
        return descargar_excel_onedrive()
    # GitHub: genera localmente con caché
    return _generar_excel_completo(hash(df.to_json()), df)


# ═════════════════════════════════════════════════════════════
# TABS
# ═════════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📝 Registro de Salida", "📊 Reportes y Estadísticas"])


# ─────────────────────────────────────────────────────────────
# TAB 1 · REGISTRO
# ─────────────────────────────────────────────────────────────
with tab1:

    if st.session_state.envio_exitoso:
        st.success("✅ ¡Registro guardado correctamente!")
        st.balloons()
        if st.button("🔄 Iniciar Nuevo Registro", type="primary", use_container_width=True):
            st.session_state.envio_exitoso = False
            st.session_state.lista_temporal = []
            invalidar_cache()
            st.rerun()

    else:
        with st.expander("🚛 **1. Datos del Vehículo y Gestor**", expanded=True):
            c1, c2 = st.columns(2)
            fecha        = c1.date_input("Fecha de salida", datetime.now())
            empresa_sel  = c1.selectbox("Empresa Gestora", options=list(GESTORES_DATA.keys()))
            empresa_final = empresa_sel
            if empresa_sel == "Otro":
                empresa_final = c1.text_input("Nombre del Gestor").upper().strip()

            conductor  = c2.text_input("Conductor")
            placa_raw  = c2.text_input("Placa (Ej: ABC123)").upper().strip()
            placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa_raw)) if placa_raw else False
            if placa_raw and not placa_valida:
                c2.warning("⚠️ Formato incorrecto. Ejemplo: **ABC123**")
            elif placa_raw and placa_valida:
                c2.success(f"✅ Placa válida: **{placa_raw}**")

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        st.subheader("⚖️ 2. Detalle de Pesajes")

        opciones_res = GESTORES_DATA.get(empresa_sel, []).copy() + ["Otro"]
        col_res, col_pes, col_btn = st.columns([3,2,1])
        tipo_sel      = col_res.selectbox("Tipo de Residuo", options=opciones_res)
        residuo_final = col_res.text_input("Especifique residuo") if tipo_sel=="Otro" else tipo_sel
        peso          = col_pes.number_input("Peso (kg)", min_value=0.0, max_value=50000.0, step=0.1, format="%.1f")

        col_btn.markdown("<br>", unsafe_allow_html=True)
        if col_btn.button("➕ Agregar", use_container_width=True):
            if not placa_valida:
                st.error("❌ Corrija el formato de la placa antes de agregar pesajes.")
            elif peso <= 0:
                st.error("❌ El peso debe ser mayor a 0 kg.")
            elif not residuo_final:
                st.error("❌ Especifique el tipo de residuo.")
            else:
                st.session_state.lista_temporal.append({"tipo_residuo": residuo_final, "peso_kg": peso})
                st.toast(f"✅ {residuo_final} — {peso:,.1f} kg")

        if st.session_state.lista_temporal:
            df_temp  = pd.DataFrame(st.session_state.lista_temporal)
            total_kg = df_temp["peso_kg"].sum()
            m1,m2,m3 = st.columns(3)
            m1.metric("Peso Total", f"{total_kg:,.1f} kg")
            m2.metric("Pesajes", len(st.session_state.lista_temporal))
            m3.metric("Promedio", f"{df_temp['peso_kg'].mean():,.1f} kg")

            with st.expander("🔍 Pesajes ingresados", expanded=True):
                df_disp = df_temp.copy()
                df_disp.index = range(1, len(df_disp)+1)
                df_disp.columns = ["Tipo de Residuo","Peso (kg)"]
                st.dataframe(df_disp, use_container_width=True)

                col_del, col_limpiar = st.columns([2,1])
                idx_del = col_del.number_input("Eliminar fila #", min_value=1,
                                               max_value=len(st.session_state.lista_temporal), step=1)
                if col_del.button("🗑️ Eliminar fila seleccionada", use_container_width=True):
                    st.session_state.lista_temporal.pop(int(idx_del)-1)
                    st.rerun()
                if col_limpiar.button("🧹 Limpiar todo", type="secondary", use_container_width=True):
                    st.session_state.lista_temporal = []
                    st.rerun()

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        st.subheader("📎 3. Evidencias y Observaciones")
        f1, f2    = st.columns(2)
        foto_memo   = f1.file_uploader("Foto del Memo (Opcional)",   type=["jpg","png","jpeg"])
        foto_camion = f2.file_uploader("Foto del Camión (Opcional)", type=["jpg","png","jpeg"])
        novedades   = st.text_area("📝 Observaciones / Novedades",
                                   placeholder="Ingrese cualquier novedad relevante...")

        st.markdown("")
        if st.button("📤 ENVIAR REGISTRO COMPLETO", type="primary", use_container_width=True):
            errores = []
            if not st.session_state.lista_temporal:   errores.append("Agregue al menos un pesaje.")
            if not placa_valida:                       errores.append("Corrija el formato de la placa.")
            if empresa_sel=="Otro" and not empresa_final: errores.append("Especifique el nombre del gestor.")
            if not conductor.strip():                  errores.append("Ingrese el nombre del conductor.")

            if errores:
                for e in errores: st.error(f"❌ {e}")
            else:
                with st.spinner("💾 Guardando registro..."):
                    try:
                        ts         = datetime.now().strftime("%Y%m%d_%H%M%S")
                        mes_actual = fecha.strftime("%B_%Y")
                        nov_final  = novedades.strip() or "Sin observaciones"

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
                                "url_memo":     "",   # se rellena en _guardar_*
                                "url_camion":   "",
                            }
                            for x in st.session_state.lista_temporal
                        ]

                        if USE_ONEDRIVE:
                            _guardar_onedrive(nuevas_filas, foto_memo, foto_camion, ts)
                        else:
                            _guardar_github(nuevas_filas, foto_memo, foto_camion,
                                            ts, fecha, empresa_final, placa_raw)

                        st.session_state.envio_exitoso = True
                        st.rerun()

                    except Exception as e:
                        # FIX: no exponer traceback completo en producción
                        st.error("❌ Error al guardar el registro. Contacte al administrador.")
                        if st.secrets.get("DEBUG", False):
                            st.exception(e)


# ─────────────────────────────────────────────────────────────
# TAB 2 · REPORTES
# ─────────────────────────────────────────────────────────────
with tab2:

    col_ref, _ = st.columns([1,5])
    if col_ref.button("🔄 Actualizar datos"):
        invalidar_cache()
        st.rerun()

    df_master = cargar_datos()

    # FIX: usar if/else en lugar de st.stop() dentro de un tab
    if df_master.empty:
        st.info("ℹ️ No hay datos disponibles aún. Registra la primera salida.")
    else:
        # ── FILTROS ──────────────────────────────────────────
        with st.expander("🔍 **Filtros**", expanded=True):
            fc1,fc2,fc3,fc4 = st.columns(4)
            # FIX: usar sentinel único "_TODOS_" para evitar choque con valores reales
            TODOS = "_TODOS_"
            empresa_f   = fc1.selectbox("Empresa",         [TODOS]+sorted(df_master["empresa"].dropna().unique().tolist()),      format_func=lambda x:"Todas" if x==TODOS else x)
            residuo_f   = fc2.selectbox("Tipo de Residuo", [TODOS]+sorted(df_master["tipo_residuo"].dropna().unique().tolist()), format_func=lambda x:"Todos" if x==TODOS else x)
            mes_f       = fc3.selectbox("Mes",             [TODOS]+sorted(df_master["mes"].dropna().unique().tolist()),          format_func=lambda x:"Todos" if x==TODOS else x)
            conductor_f = fc4.selectbox("Conductor",       [TODOS]+sorted(df_master["conductor"].dropna().unique().tolist()),    format_func=lambda x:"Todos" if x==TODOS else x)

        df_f = df_master.copy()
        if empresa_f   != TODOS: df_f = df_f[df_f["empresa"]      == empresa_f]
        if residuo_f   != TODOS: df_f = df_f[df_f["tipo_residuo"] == residuo_f]
        if mes_f       != TODOS: df_f = df_f[df_f["mes"]          == mes_f]
        if conductor_f != TODOS: df_f = df_f[df_f["conductor"]    == conductor_f]

        # ── KPIs ─────────────────────────────────────────────
        # FIX: calcular groupbys una sola vez y reutilizar abajo
        total_kg    = df_f["peso_kg"].sum()
        n_registros = len(df_f)
        n_empresas  = df_f["empresa"].nunique()
        promedio    = df_f["peso_kg"].mean() if n_registros else 0
        n_residuos  = df_f["tipo_residuo"].nunique()

        # Agrupaciones pre-calculadas (reutilizadas en gráficos y tablas)
        res_emp = (df_f.groupby("empresa")["peso_kg"].sum().round(1)
                   .sort_values(ascending=False).reset_index())
        res_res = (df_f.groupby("tipo_residuo")["peso_kg"].sum().round(1)
                   .sort_values(ascending=False).reset_index())
        df_tend = (df_f.groupby(["mes","empresa"])["peso_kg"].sum().round(1)
                   .reset_index().sort_values("mes"))
        resumen_tabla = (
            df_f.groupby("empresa")
            .agg(Registros=("peso_kg","count"), Peso_Total_kg=("peso_kg","sum"),
                 Peso_Promedio_kg=("peso_kg","mean"), Peso_Maximo_kg=("peso_kg","max"))
            .round(1).sort_values("Peso_Total_kg", ascending=False).reset_index()
        )
        pivot_show = (df_f.pivot_table(index="empresa", columns="tipo_residuo",
                                        values="peso_kg", aggfunc="sum", fill_value=0)
                      .round(1).reset_index())

        k1,k2,k3,k4,k5 = st.columns(5)
        for col, val, label in zip(
            [k1,k2,k3,k4,k5],
            [f"{total_kg:,.1f} kg", f"{n_registros:,}", str(n_empresas),
             f"{promedio:,.1f} kg", str(n_residuos)],
            ["Peso Total","Registros","Empresas Activas","Promedio / Registro","Tipos de Residuo"],
        ):
            col.markdown(f'<div class="kpi-card"><div class="kpi-value">{val}</div>'
                         f'<div class="kpi-label">{label}</div></div>', unsafe_allow_html=True)

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

        # ── GRÁFICAS ─────────────────────────────────────────
        st.subheader("📈 Visualizaciones")
        g1, g2 = st.columns(2)

        with g1:
            colores = [COLORES_EMPRESA.get(e, COLOR_DEFAULT) for e in res_emp["empresa"]]
            fig_emp = px.bar(res_emp, x="empresa", y="peso_kg",
                             title="⚖️ Peso Total por Empresa (kg)",
                             labels={"empresa":"Empresa","peso_kg":"Peso (kg)"},
                             color="empresa", color_discrete_sequence=colores, text_auto=".1f")
            fig_emp.update_layout(showlegend=False, plot_bgcolor="white", title_font_size=15)
            fig_emp.update_traces(textposition="outside")
            st.plotly_chart(fig_emp, use_container_width=True)

        with g2:
            fig_donut = px.pie(res_res, names="tipo_residuo", values="peso_kg",
                               title="🗂️ Distribución por Tipo de Residuo", hole=0.45,
                               color_discrete_sequence=px.colors.qualitative.Set3)
            fig_donut.update_traces(textposition="inside", textinfo="percent+label")
            fig_donut.update_layout(title_font_size=15)
            st.plotly_chart(fig_donut, use_container_width=True)

        if len(df_tend["mes"].unique()) > 1:
            fig_tend = px.line(df_tend, x="mes", y="peso_kg", color="empresa",
                               title="📅 Tendencia Mensual por Empresa (kg)",
                               labels={"mes":"Mes","peso_kg":"Peso (kg)","empresa":"Empresa"},
                               markers=True, color_discrete_map=COLORES_EMPRESA)
            fig_tend.update_layout(plot_bgcolor="white", title_font_size=15)
            st.plotly_chart(fig_tend, use_container_width=True)

        pivot_heat = df_f.pivot_table(index="empresa", columns="tipo_residuo",
                                       values="peso_kg", aggfunc="sum", fill_value=0).round(1)
        if not pivot_heat.empty:
            fig_heat = px.imshow(pivot_heat, title="🗺️ Mapa de Calor: Empresa × Tipo de Residuo (kg)",
                                 color_continuous_scale="Blues", aspect="auto", text_auto=".0f")
            fig_heat.update_layout(title_font_size=15)
            st.plotly_chart(fig_heat, use_container_width=True)

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

        # ── TABLAS ───────────────────────────────────────────
        st.subheader("📋 Tablas Detalladas")
        t1,t2,t3 = st.tabs(["Datos Completos","Resumen por Empresa","Pivot Residuos"])

        with t1:
            cols = ["fecha","empresa","conductor","placa","tipo_residuo","peso_kg","mes","novedades"]
            df_tabla = df_f[[c for c in cols if c in df_f.columns]].sort_values("fecha", ascending=False)
            st.dataframe(df_tabla, use_container_width=True, hide_index=True,
                         column_config={
                             "fecha":    st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                             "peso_kg":  st.column_config.NumberColumn("Peso (kg)", format="%.1f"),
                         })
        with t2:
            st.dataframe(resumen_tabla, use_container_width=True, hide_index=True,
                         column_config={
                             "Peso_Total_kg":    st.column_config.NumberColumn("Peso Total (kg)",  format="%.1f"),
                             "Peso_Promedio_kg": st.column_config.NumberColumn("Promedio (kg)",    format="%.1f"),
                             "Peso_Maximo_kg":   st.column_config.NumberColumn("Máximo (kg)",      format="%.1f"),
                         })
        with t3:
            st.dataframe(pivot_show, use_container_width=True, hide_index=True)

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

        # ── DESCARGA ─────────────────────────────────────────
        st.subheader("⬇️ Exportar Datos")
        with st.spinner("Preparando Excel..."):
            excel_bytes = generar_excel_para_descarga(df_f)
        st.download_button(
            label="📥 Descargar Excel Completo",
            data=excel_bytes,
            file_name=f"Reporte_TINTATEX_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary",
        )
