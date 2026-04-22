import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re
import io
import plotly.express as px

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────
REPO_NAME = "alejandraho013/registro-salida-residuos"

GESTORES_DATA = {
    "CORPOGESTAR":              sorted(["Cartón limpio","Cartón sucio","Papel de archivo","Pasta","PET limpio","PET sucio","Plástico","Retal de tela","Tubo plega"]),
    "Recicla Oriente":          sorted(["Cartón limpio","Cartón sucio","Papel de archivo","Pasta","PET limpio","PET sucio","Plástico","Retal de tela","Tubo plega"]),
    "Quimetales NO Peligrosos": sorted(["Algodón","Retal de tela","Tubo plega"]),
    "Quimetales Peligrosos":    sorted(["RAEE","Residuos laboratorio","CDR"]),
    "Otro": [],
}

COLORES_EMPRESA = {
    "CORPOGESTAR":              "#2196F3",
    "Recicla Oriente":          "#4CAF50",
    "Quimetales NO Peligrosos": "#FF9800",
    "Quimetales Peligrosos":    "#F44336",
}

# ─────────────────────────────────────────────────────────────
# SECRETS
# ─────────────────────────────────────────────────────────────
try:
    TOKEN     = st.secrets["TOKEN"]
    REPO_NAME = st.secrets.get("REPO_NAME", REPO_NAME)
except Exception:
    st.error("⚠️ Configura el TOKEN en los Secrets de Streamlit Cloud.")
    st.stop()

# ─────────────────────────────────────────────────────────────
# ESTILOS Y SESSION STATE
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="TINTATEX · Gestión de Residuos", layout="wide")

st.markdown("""
<style>
.main-header { background: linear-gradient(135deg,#1a237e 0%,#283593 100%); padding:1.5rem; border-radius:12px; color:white; margin-bottom:1.5rem; }
.kpi-card { background:white; border-radius:10px; padding:1rem; border-left:4px solid #1a237e; box-shadow:0 2px 5px rgba(0,0,0,.1); text-align:center; }
.kpi-value { font-size:1.8rem; font-weight:700; color:#1a237e; }
.kpi-label { font-size:0.9rem; color:#666; }
</style>
<div class="main-header"><h1>🏭 TINTATEX · Gestión de Residuos</h1><p>Registro de Pesajes y Control de Carga</p></div>
""", unsafe_allow_html=True)

for key, default in [("lista_temporal", []), ("envio_exitoso", False)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ─────────────────────────────────────────────────────────────
# FUNCIONES DE DATOS
# ─────────────────────────────────────────────────────────────
COLUMNAS_MASTER = ["fecha","mes","empresa","conductor","placa","tipo_residuo","peso_kg","novedades"]

def _get_repo():
    return Github(TOKEN).get_repo(REPO_NAME)

@st.cache_data(ttl=120)   # TTL corto para ver datos recién guardados
def cargar_datos_github() -> pd.DataFrame:
    """
    Carga database.xlsx desde GitHub.
    Retorna DataFrame vacío si el archivo no existe todavía.
    Muestra un error visible si hay otro problema (token, permisos, etc.).
    """
    try:
        repo = _get_repo()
        try:
            f = repo.get_contents("database.xlsx")
        except Exception:
            # El archivo aún no existe → DataFrame vacío, sin error ruidoso
            return pd.DataFrame(columns=COLUMNAS_MASTER)

        df = pd.read_excel(
            io.BytesIO(f.decoded_content),
            sheet_name="MASTER",
            engine="openpyxl",
        )

        # Normalizar columnas obligatorias
        df["fecha"] = pd.to_datetime(df.get("fecha"), errors="coerce")
        if "mes" not in df.columns or df["mes"].isna().all():
            df["mes"] = df["fecha"].dt.strftime("%B_%Y")
        if "peso_kg" in df.columns:
            df["peso_kg"] = pd.to_numeric(df["peso_kg"], errors="coerce").fillna(0)

        return df

    except Exception as e:
        st.error(f"❌ Error al cargar datos desde GitHub: {e}")
        return pd.DataFrame(columns=COLUMNAS_MASTER)


def guardar_datos(nuevas_filas: list[dict], empresa: str, placa: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    repo = _get_repo()

    try:
        xlsx_file = repo.get_contents("database.xlsx")
        dicc = pd.read_excel(
            io.BytesIO(xlsx_file.decoded_content),
            sheet_name=None,
            engine="openpyxl",
        )
        sha = xlsx_file.sha
    except Exception:
        # El archivo no existe → lo creamos desde cero
        dicc = {}
        sha = None

    df_nuevos = pd.DataFrame(nuevas_filas)

    # Hoja MASTER
    dicc["MASTER"] = pd.concat(
        [dicc.get("MASTER", pd.DataFrame(columns=COLUMNAS_MASTER)), df_nuevos],
        ignore_index=True,
    )

    # Hoja por empresa
    nombre_hoja = re.sub(r'[\\/*?:\[\]]', "", empresa)[:30].upper() or "OTROS"
    dicc[nombre_hoja] = pd.concat(
        [dicc.get(nombre_hoja, pd.DataFrame(columns=COLUMNAS_MASTER)), df_nuevos],
        ignore_index=True,
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for hoja, datos in dicc.items():
            datos.to_excel(writer, sheet_name=hoja, index=False)

    commit_msg = f"Reg_{empresa}_{placa}_{ts}"
    if sha:
        repo.update_file("database.xlsx", commit_msg, output.getvalue(), sha)
    else:
        repo.create_file("database.xlsx", commit_msg, output.getvalue())

    # Invalidar caché para que Reportes muestre los datos nuevos de inmediato
    cargar_datos_github.clear()

# ─────────────────────────────────────────────────────────────
# INTERFAZ
# ─────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📝 Registro", "📊 Reportes y Gráficas"])

# ── TAB 1: REGISTRO ──────────────────────────────────────────
with tab1:
    if st.session_state.envio_exitoso:
        st.success("✅ ¡Registro guardado con éxito!")
        if st.button("🔄 Nuevo Registro", type="primary"):
            st.session_state.envio_exitoso = False
            st.session_state.lista_temporal = []
            st.rerun()
    else:
        with st.expander("🚛 1. Datos del Vehículo", expanded=True):
            c1, c2 = st.columns(2)
            fecha      = c1.date_input("Fecha", datetime.now())
            emp_sel    = c1.selectbox("Empresa Gestora", options=list(GESTORES_DATA.keys()))
            empresa_final = (
                c1.text_input("Nombre Manual").upper().strip()
                if emp_sel == "Otro"
                else emp_sel
            )

            conductor = c2.text_input("Conductor")
            cp1, cp2  = c2.columns([2, 1])
            placa     = cp1.text_input("Placa (ABC123)").upper().strip()
            placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa))
            if placa:
                (cp2.success if placa_valida else cp2.error)(
                    "✔ Válida" if placa_valida else "✗ Inválida"
                )

            capacidad = c2.number_input("Capacidad Camión (kg) — Opcional", min_value=0.0, step=100.0)

        st.subheader("⚖️ 2. Pesajes")
        col_r, col_p, col_b = st.columns([3, 2, 1])
        res_opts  = GESTORES_DATA.get(emp_sel, []) + ["Otro"]
        tipo_res  = col_r.selectbox("Tipo de Residuo", options=res_opts)
        res_final = col_r.text_input("¿Cuál residuo?") if tipo_res == "Otro" else tipo_res
        peso      = col_p.number_input("Peso (kg)", min_value=0.0, step=0.1)

        if col_b.button("➕ Añadir", use_container_width=True):
            if not placa_valida:
                st.error("Placa inválida — formato esperado: ABC123")
            elif peso <= 0:
                st.error("El peso debe ser mayor a 0")
            else:
                st.session_state.lista_temporal.append(
                    {"tipo_residuo": res_final, "peso_kg": peso}
                )
                st.toast(f"✅ {res_final} — {peso} kg añadido")

        if st.session_state.lista_temporal:
            df_temp  = pd.DataFrame(st.session_state.lista_temporal)
            total_kg = df_temp["peso_kg"].sum()

            if capacidad > 0 and total_kg > capacidad:
                st.warning(f"⚠️ Carga ({total_kg:,.1f} kg) supera la capacidad ({capacidad:,.1f} kg)")

            m1, m2 = st.columns(2)
            m1.metric("Suma Carga", f"{total_kg:,.1f} kg")
            m2.metric("Registros",  len(df_temp))
            st.dataframe(df_temp, use_container_width=True, hide_index=True)

            b1, b2 = st.columns(2)
            if b1.button("⏪ Eliminar último"):
                st.session_state.lista_temporal.pop()
                st.rerun()
            if b2.button("🧹 Limpiar lista"):
                st.session_state.lista_temporal = []
                st.rerun()

        novedades = st.text_area("Observaciones")

        if st.button("📤 ENVIAR TODO", type="primary", use_container_width=True):
            if not st.session_state.lista_temporal:
                st.error("Agrega al menos un pesaje antes de enviar.")
            elif not placa_valida:
                st.error("La placa no es válida.")
            elif not empresa_final:
                st.error("Escribe el nombre de la empresa gestora.")
            else:
                with st.spinner("Guardando en GitHub…"):
                    filas = [
                        {
                            "fecha":        str(fecha),
                            "mes":          fecha.strftime("%B_%Y"),
                            "empresa":      empresa_final,
                            "conductor":    conductor,
                            "placa":        placa,
                            "tipo_residuo": x["tipo_residuo"],
                            "peso_kg":      x["peso_kg"],
                            "novedades":    novedades or "Sin novedades",
                        }
                        for x in st.session_state.lista_temporal
                    ]
                    try:
                        guardar_datos(filas, empresa_final, placa)
                        st.session_state.envio_exitoso = True
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error al guardar: {e}")

# ── TAB 2: REPORTES ──────────────────────────────────────────
with tab2:
    if st.button("🔄 Actualizar datos", key="refresh"):
        cargar_datos_github.clear()
        st.rerun()

    with st.spinner("Cargando datos…"):
        df = cargar_datos_github()

    if df.empty:
        st.info("📭 Sin datos registrados aún. Realiza tu primer registro en la pestaña anterior.")
    else:
        # Filtros
        meses_disp = sorted(df["mes"].dropna().unique().tolist())
        f_mes = st.selectbox("Filtrar por Mes", ["Todos"] + meses_disp)
        df_f  = df if f_mes == "Todos" else df[df["mes"] == f_mes]

        if df_f.empty:
            st.warning("No hay registros para el mes seleccionado.")
        else:
            # KPIs
            k1, k2, k3 = st.columns(3)
            k1.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["peso_kg"].sum():,.1f} kg</div><div class="kpi-label">Peso Total</div></div>', unsafe_allow_html=True)
            k2.markdown(f'<div class="kpi-card"><div class="kpi-value">{len(df_f)}</div><div class="kpi-label">Registros</div></div>', unsafe_allow_html=True)
            k3.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["empresa"].nunique()}</div><div class="kpi-label">Empresas</div></div>', unsafe_allow_html=True)

            st.divider()

            # Gráficas
            g1, g2 = st.columns(2)
            with g1:
                res_emp = df_f.groupby("empresa")["peso_kg"].sum().reset_index()
                fig_bar = px.bar(
                    res_emp, x="empresa", y="peso_kg",
                    title="Peso total por Empresa (kg)",
                    color="empresa",
                    color_discrete_map=COLORES_EMPRESA,
                    labels={"peso_kg": "Peso (kg)", "empresa": ""},
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            with g2:
                fig_pie = px.pie(
                    df_f, values="peso_kg", names="tipo_residuo",
                    title="Distribución por Tipo de Residuo",
                )
                st.plotly_chart(fig_pie, use_container_width=True)

            # Evolución temporal
            if "fecha" in df_f.columns and df_f["fecha"].notna().any():
                df_tiempo = (
                    df_f.groupby(df_f["fecha"].dt.date)["peso_kg"]
                    .sum()
                    .reset_index()
                    .rename(columns={"fecha": "Fecha", "peso_kg": "Peso (kg)"})
                )
                fig_line = px.line(
                    df_tiempo, x="Fecha", y="Peso (kg)",
                    title="Evolución Diaria de Peso",
                    markers=True,
                )
                st.plotly_chart(fig_line, use_container_width=True)

            # Tabla y descarga
            st.subheader("📋 Detalle de Registros")
            st.dataframe(df_f, use_container_width=True, hide_index=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_f.to_excel(writer, index=False, sheet_name="MASTER")
            st.download_button(
                "⬇️ Descargar Excel",
                data=output.getvalue(),
                file_name=f"Reporte_{f_mes}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
