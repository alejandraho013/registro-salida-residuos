import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re
import io
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

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
    st.error("⚠️ Configura el TOKEN en 'Advanced Settings'.")
    st.stop()

USE_ONEDRIVE = all(k in st.secrets for k in ("AZURE_TENANT_ID","AZURE_CLIENT_ID","AZURE_CLIENT_SECRET","ONEDRIVE_FILE_ID"))
if USE_ONEDRIVE:
    from onedrive import cargar_datos_onedrive, append_filas_onedrive, descargar_excel_onedrive

# ─────────────────────────────────────────────────────────────
# ESTILOS Y CABECERA
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="TINTATEX · Gestión de Residuos", layout="wide")

st.markdown("""
<style>
.main-header { background: linear-gradient(135deg,#1a237e 0%,#283593 100%); padding:1.5rem; border-radius:12px; color:white; margin-bottom:1.5rem; }
.kpi-card { background:white; border-radius:10px; padding:1rem; border-left:4px solid #1a237e; box-shadow:0 2px 5px rgba(0,0,0,.1); text-align:center; }
.kpi-value { font-size:1.8rem; font-weight:700; color:#1a237e; }
</style>
<div class="main-header"><h1>🏭 TINTATEX · Gestión de Residuos</h1><p>Registro rápido de pesajes (Sin Fotos)</p></div>
""", unsafe_allow_html=True)

if "lista_temporal" not in st.session_state: st.session_state.lista_temporal = []
if "envio_exitoso" not in st.session_state: st.session_state.envio_exitoso = False

# ─────────────────────────────────────────────────────────────
# GESTIÓN DE DATOS (CRUD)
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=600)
def cargar_datos_github():
    try:
        repo = Github(TOKEN).get_repo(REPO_NAME)
        f = repo.get_contents("database.xlsx")
        df = pd.read_excel(io.BytesIO(f.decoded_content), sheet_name="MASTER", engine="openpyxl")
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        return df
    except Exception: return pd.DataFrame()

def guardar_datos(nuevas_filas, empresa, placa, fecha):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if USE_ONEDRIVE:
        append_filas_onedrive(nuevas_filas)
    else:
        repo = Github(TOKEN).get_repo(REPO_NAME)
        xlsx_file = repo.get_contents("database.xlsx")
        dicc = pd.read_excel(io.BytesIO(xlsx_file.decoded_content), sheet_name=None, engine="openpyxl")
        
        df_nuevos = pd.DataFrame(nuevas_filas)
        dicc["MASTER"] = pd.concat([dicc.get("MASTER", pd.DataFrame()), df_nuevos], ignore_index=True)
        
        # Hoja por empresa
        nombre_hoja = re.sub(r'[\\/*?:\[\]]', "", empresa)[:30].upper() or "OTROS"
        dicc[nombre_hoja] = pd.concat([dicc.get(nombre_hoja, pd.DataFrame()), df_nuevos], ignore_index=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for hoja, df_h in dicc.items(): df_h.to_excel(writer, sheet_name=hoja, index=False)
        
        repo.update_file("database.xlsx", f"Reg_{empresa}_{placa}_{ts}", output.getvalue(), xlsx_file.sha)

# ─────────────────────────────────────────────────────────────
# INTERFAZ
# ─────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📝 Nuevo Registro", "📊 Reportes"])

with tab1:
    if st.session_state.envio_exitoso:
        st.success("✅ ¡Registro guardado con éxito!")
        if st.button("🔄 Nuevo Registro", type="primary"):
            st.session_state.envio_exitoso = False
            st.session_state.lista_temporal = []
            cargar_datos_github.clear()
            st.rerun()
    else:
        # 1. Datos Generales
        c1, c2 = st.columns(2)
        fecha = c1.date_input("Fecha", datetime.now())
        emp_sel = c1.selectbox("Empresa Gestora", options=list(GESTORES_DATA.keys()))
        empresa_final = c1.text_input("Nombre Manual").upper() if emp_sel == "Otro" else emp_sel
        conductor = c2.text_input("Conductor")
        placa = c2.text_input("Placa (Ej: ABC123)").upper().strip()
        placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa))

        # 2. Pesajes
        st.divider()
        col_r, col_p, col_b = st.columns([3,2,1])
        res_opts = GESTORES_DATA.get(emp_sel, []) + ["Otro"]
        tipo_res = col_r.selectbox("Tipo de Residuo", options=res_opts)
        res_final = col_r.text_input("¿Cuál?") if tipo_res == "Otro" else tipo_res
        peso = col_p.number_input("Peso (kg)", min_value=0.0, step=0.1)
        
        if col_b.button("➕ Añadir", use_container_width=True):
            if placa_valida and peso > 0:
                st.session_state.lista_temporal.append({"tipo_residuo": res_final, "peso_kg": peso})
                st.toast("Añadido")
            else: st.error("Revisa placa/peso")

        if st.session_state.lista_temporal:
            df_temp = pd.DataFrame(st.session_state.lista_temporal)
            st.dataframe(df_temp, use_container_width=True)
            st.info(f"Peso Total: {df_temp['peso_kg'].sum():,.1f} kg")

        # 3. Finalizar
        novedades = st.text_area("Observaciones")
        if st.button("📤 ENVIAR TODO", type="primary", use_container_width=True):
            if not st.session_state.lista_temporal or not placa_valida:
                st.error("Datos incompletos o placa errónea.")
            else:
                with st.spinner("Guardando..."):
                    filas = [{
                        "fecha": str(fecha), "mes": fecha.strftime("%B_%Y"),
                        "empresa": empresa_final, "conductor": conductor, "placa": placa,
                        "tipo_residuo": x["tipo_residuo"], "peso_kg": x["peso_kg"],
                        "novedades": novedades or "Sin novedades"
                    } for x in st.session_state.lista_temporal]
                    
                    guardar_datos(filas, empresa_final, placa, fecha)
                    st.session_state.envio_exitoso = True
                    st.rerun()

with tab2:
    df = cargar_datos_onedrive() if USE_ONEDRIVE else cargar_datos_github()
    if df.empty: st.info("No hay datos.")
    else:
        # Filtros rápidos
        mes_f = st.selectbox("Filtrar por Mes", ["Todos"] + sorted(df["mes"].unique().tolist()))
        df_f = df if mes_f == "Todos" else df[df["mes"] == mes_f]
        
        # KPIs
        k1, k2, k3 = st.columns(3)
        k1.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["peso_kg"].sum():,.1f} kg</div>Total Peso</div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card"><div class="kpi-value">{len(df_f)}</div>Viajes</div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["empresa"].nunique()}</div>Gestores</div>', unsafe_allow_html=True)
        
        st.dataframe(df_f, use_container_width=True, hide_index=True)
        
        # Exportación
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_f.to_excel(writer, index=False, sheet_name="Reporte")
        st.download_button("⬇️ Descargar Excel", output.getvalue(), "Reporte_Residuos.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
