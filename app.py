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
    "Recicla Oriente":          sorted(["Cartón limpio","Cartón sucio","Papel de archivo","Pasta","PET limpio","PET sucio","Plástico","Retal de tela","Tubo plega"]),
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

try:
    TOKEN     = st.secrets["TOKEN"]
    REPO_NAME = st.secrets.get("REPO_NAME", REPO_NAME)
except Exception:
    st.error("⚠️ Configura el TOKEN en 'Advanced Settings' > Secrets.")
    st.stop()

USE_ONEDRIVE = all(
    k in st.secrets for k in ("AZURE_TENANT_ID","AZURE_CLIENT_ID",
                               "AZURE_CLIENT_SECRET","ONEDRIVE_FILE_ID")
)
if USE_ONEDRIVE:
    from onedrive import cargar_datos_onedrive, append_filas_onedrive, descargar_excel_onedrive

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
    <p>Sistema de registro y trazabilidad de salida de residuos (Sin Fotos)</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
if "lista_temporal"  not in st.session_state: st.session_state.lista_temporal  = []
if "envio_exitoso"   not in st.session_state: st.session_state.envio_exitoso   = False

# ─────────────────────────────────────────────────────────────
# RECURSOS CACHEADOS
# ─────────────────────────────────────────────────────────────
@st.cache_resource
def _github_repo():
    return Github(TOKEN).get_repo(REPO_NAME)

@st.cache_data(ttl=600)
def _cargar_datos_github() -> pd.DataFrame:
    try:
        repo  = _github_repo()
        f     = repo.get_contents("database.xlsx")
        df    = pd.read_excel(io.BytesIO(f.decoded_content), sheet_name="MASTER", engine="openpyxl")
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        return df
    except Exception:
        return pd.DataFrame()

def cargar_datos() -> pd.DataFrame:
    if USE_ONEDRIVE: return cargar_datos_onedrive()
    return _cargar_datos_github()

def invalidar_cache():
    if USE_ONEDRIVE: cargar_datos_onedrive.clear()
    else: _cargar_datos_github.clear()

# ─────────────────────────────────────────────────────────────
# GUARDAR REGISTRO
# ─────────────────────────────────────────────────────────────
def _guardar_onedrive(nuevas_filas):
    append_filas_onedrive(nuevas_filas)

def _guardar_github(nuevas_filas, ts, fecha, empresa_final, placa_raw):
    repo = _github_repo()
    xlsx_file = repo.get_contents("database.xlsx")
    df_master = pd.read_excel(io.BytesIO(xlsx_file.decoded_content), sheet_name="MASTER", engine="openpyxl")
    
    df_nuevos = pd.DataFrame(nuevas_filas)
    df_master = pd.concat([df_master, df_nuevos], ignore_index=True)

    nombre_hoja = re.sub(r'[\\/*?:\[\]]', "", empresa_final)[:30].strip().upper() or "OTROS"

    try:
        df_emp = pd.read_excel(io.BytesIO(xlsx_file.decoded_content), sheet_name=nombre_hoja, engine="openpyxl")
        df_emp = pd.concat([df_emp, df_nuevos], ignore_index=True)
    except Exception:
        df_emp = df_nuevos.copy()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="MASTER", index=False)
        df_emp.to_excel(writer, sheet_name=nombre_hoja, index=False)

    repo.update_file("database.xlsx", f"Registro {empresa_final} {placa_raw} {ts}", output.getvalue(), xlsx_file.sha)

# ─────────────────────────────────────────────────────────────
# EXCEL DE DESCARGA
# ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _generar_excel_completo(df_hash: int, df_master: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    df_m = df_master.copy()
    df_m["fecha"] = df_m["fecha"].astype(str)

    resumen_empresa = (df_m.groupby("empresa").agg(Registros=("peso_kg","count"), Peso_Total_kg=("peso_kg","sum")).reset_index())

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_m.to_excel(writer, sheet_name="MASTER", index=False)
        resumen_empresa.to_excel(writer, sheet_name="Resumen", index=False)
    return output.getvalue()

def generar_excel_para_descarga(df: pd.DataFrame) -> bytes:
    if USE_ONEDRIVE: return descargar_excel_onedrive()
    return _generar_excel_completo(hash(df.to_json()), df)

# ═════════════════════════════════════════════════════════════
# TABS
# ═════════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📝 Registro de Salida", "📊 Reportes"])

with tab1:
    if st.session_state.envio_exitoso:
        st.success("✅ ¡Registro guardado correctamente!")
        if st.button("🔄 Iniciar Nuevo Registro", type="primary", use_container_width=True):
            st.session_state.envio_exitoso = False
            st.session_state.lista_temporal = []
            invalidar_cache()
            st.rerun()
    else:
        with st.expander("🚛 **1. Datos del Vehículo y Gestor**", expanded=True):
            c1, c2 = st.columns(2)
            fecha = c1.date_input("Fecha", datetime.now())
            empresa_sel = c1.selectbox("Empresa Gestora", options=list(GESTORES_DATA.keys()))
            empresa_final = c1.text_input("Nombre del Gestor").upper().strip() if empresa_sel == "Otro" else empresa_sel
            conductor = c2.text_input("Conductor")
            placa_raw = c2.text_input("Placa (Ej: ABC123)").upper().strip()
            placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa_raw)) if placa_raw else False

        st.subheader("⚖️ 2. Detalle de Pesajes")
        opciones_res = GESTORES_DATA.get(empresa_sel, []).copy() + ["Otro"]
        col_res, col_pes, col_btn = st.columns([3,2,1])
        tipo_sel = col_res.selectbox("Tipo de Residuo", options=opciones_res)
        residuo_final = col_res.text_input("Especifique residuo") if tipo_sel=="Otro" else tipo_sel
        peso = col_pes.number_input("Peso (kg)", min_value=0.0, step=0.1)

        col_btn.markdown("<br>", unsafe_allow_html=True)
        if col_btn.button("➕ Agregar", use_container_width=True):
            if not placa_valida: st.error("❌ Placa inválida.")
            elif peso <= 0: st.error("❌ Peso inválido.")
            else:
                st.session_state.lista_temporal.append({"tipo_residuo": residuo_final, "peso_kg": peso})
                st.toast(f"✅ Agregado: {peso} kg")

        if st.session_state.lista_temporal:
            df_temp = pd.DataFrame(st.session_state.lista_temporal)
            st.metric("Peso Total", f"{df_temp['peso_kg'].sum():,.1f} kg")
            st.dataframe(df_temp, use_container_width=True)
            if st.button("🧹 Limpiar todo"):
                st.session_state.lista_temporal = []
                st.rerun()

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        novedades = st.text_area("📝 Observaciones / Novedades")

        if st.button("📤 ENVIAR REGISTRO", type="primary", use_container_width=True):
            if not st.session_state.lista_temporal or not placa_valida:
                st.error("❌ Verifique los datos.")
            else:
                with st.spinner("💾 Guardando..."):
                    try:
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        nuevas_filas = [{
                            "fecha": str(fecha),
                            "mes": fecha.strftime("%B_%Y"),
                            "empresa": empresa_final,
                            "conductor": conductor.strip(),
                            "placa": placa_raw,
                            "tipo_residuo": x["tipo_residuo"],
                            "peso_kg": x["peso_kg"],
                            "novedades": novedades.strip() or "Sin observaciones"
                        } for x in st.session_state.lista_temporal]

                        if USE_ONEDRIVE: _guardar_onedrive(nuevas_filas)
                        else: _guardar_github(nuevas_filas, ts, fecha, empresa_final, placa_raw)

                        st.session_state.envio_exitoso = True
                        st.rerun()
                    except Exception as e: st.error(f"❌ Error: {e}")

with tab2:
    df_master = cargar_datos()
    if df_master.empty:
        st.info("No hay datos.")
    else:
        st.subheader("📊 Reporte Consolidado")
        st.dataframe(df_master, use_container_width=True)
        
        excel_bytes = generar_excel_para_descarga(df_master)
        st.download_button(
            label="📥 Descargar Excel",
            data=excel_bytes,
            file_name="Reporte_TINTATEX.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
            use_container_width=True, type="primary",
        )
