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

try:
    TOKEN     = st.secrets["TOKEN"]
    REPO_NAME = st.secrets.get("REPO_NAME", REPO_NAME)
except Exception:
    st.error("⚠️ Configura el TOKEN en los Secrets.")
    st.stop()

USE_ONEDRIVE = all(k in st.secrets for k in ("AZURE_TENANT_ID","AZURE_CLIENT_ID","AZURE_CLIENT_SECRET","ONEDRIVE_FILE_ID"))
if USE_ONEDRIVE:
    from onedrive import cargar_datos_onedrive, append_filas_onedrive, descargar_excel_onedrive

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

if "lista_temporal" not in st.session_state: st.session_state.lista_temporal = []
if "envio_exitoso" not in st.session_state: st.session_state.envio_exitoso = False

# ─────────────────────────────────────────────────────────────
# FUNCIONES DE DATOS
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

def guardar_datos(nuevas_filas, empresa, placa):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if USE_ONEDRIVE:
        append_filas_onedrive(nuevas_filas)
    else:
        repo = Github(TOKEN).get_repo(REPO_NAME)
        xlsx_file = repo.get_contents("database.xlsx")
        dicc = pd.read_excel(io.BytesIO(xlsx_file.decoded_content), sheet_name=None, engine="openpyxl")
        df_nuevos = pd.DataFrame(nuevas_filas)
        dicc["MASTER"] = pd.concat([dicc.get("MASTER", pd.DataFrame()), df_nuevos], ignore_index=True)
        nombre_hoja = re.sub(r'[\\/*?:\[\]]', "", empresa)[:30].upper() or "OTROS"
        dicc[nombre_hoja] = pd.concat([dicc.get(nombre_hoja, pd.DataFrame()), df_nuevos], ignore_index=True)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for h, d in dicc.items(): d.to_excel(writer, sheet_name=h, index=False)
        repo.update_file("database.xlsx", f"Reg_{empresa}_{placa}_{ts}", output.getvalue(), xlsx_file.sha)

# ─────────────────────────────────────────────────────────────
# INTERFAZ
# ─────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📝 Registro", "📊 Reportes y Gráficas"])

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
            fecha = c1.date_input("Fecha", datetime.now())
            emp_sel = c1.selectbox("Empresa Gestora", options=list(GESTORES_DATA.keys()))
            empresa_final = c1.text_input("Nombre Manual").upper() if emp_sel == "Otro" else emp_sel
            
            conductor = c2.text_input("Conductor")
            cp1, cp2 = c2.columns([2, 1])
            placa = cp1.text_input("Placa (ABC123)").upper().strip()
            placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa))
            if placa:
                if placa_valida: cp2.success("Válida")
                else: cp2.error("Inválida")
            
            capacidad = c2.number_input("Capacidad Camión (kg) - Opcional", min_value=0.0, step=100.0)

        st.subheader("⚖️ 2. Pesajes")
        col_r, col_p, col_b = st.columns([3,2,1])
        res_opts = GESTORES_DATA.get(emp_sel, []) + ["Otro"]
        tipo_res = col_r.selectbox("Tipo de Residuo", options=res_opts)
        res_final = col_r.text_input("¿Cuál?") if tipo_res == "Otro" else tipo_res
        peso = col_p.number_input("Peso (kg)", min_value=0.0, step=0.1)
        
        if col_b.button("➕ Añadir", use_container_width=True):
            if placa_valida and peso > 0:
                st.session_state.lista_temporal.append({"tipo_residuo": res_final, "peso_kg": peso})
                st.toast("Añadido")
            else: st.error("Revisa placa y peso")

        if st.session_state.lista_temporal:
            df_temp = pd.DataFrame(st.session_state.lista_temporal)
            total_kg = df_temp['peso_kg'].sum()
            
            if capacidad > 0 and total_kg > capacidad:
                st.warning(f"⚠️ Carga ({total_kg:,.1f} kg) supera la capacidad ({capacidad:,.1f} kg)")
            
            m1, m2 = st.columns(2)
            m1.metric("Suma Carga", f"{total_kg:,.1f} kg")
            m2.metric("Registros", len(df_temp))
            
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
            if not st.session_state.lista_temporal or not placa_valida:
                st.error("Datos incompletos.")
            else:
                with st.spinner("Guardando..."):
                    filas = [{
                        "fecha": str(fecha), "mes": fecha.strftime("%B_%Y"),
                        "empresa": empresa_final, "conductor": conductor, "placa": placa,
                        "tipo_residuo": x["tipo_residuo"], "peso_kg": x["peso_kg"],
                        "novedades": novedades or "Sin novedades"
                    } for x in st.session_state.lista_temporal]
                    guardar_datos(filas, empresa_final, placa)
                    st.session_state.envio_exitoso = True
                    st.rerun()

with tab2:
    df = cargar_datos_onedrive() if USE_ONEDRIVE else cargar_datos_github()
    if df.empty: st.info("Sin datos.")
    else:
        f_mes = st.selectbox("Filtrar Mes", ["Todos"] + sorted(df["mes"].unique().tolist()))
        df_f = df if f_mes == "Todos" else df[df["mes"] == f_mes]

        k1, k2, k3 = st.columns(3)
        k1.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["peso_kg"].sum():,.1f} kg</div>Peso Total</div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card"><div class="kpi-value">{len(df_f)}</div>Registros</div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["empresa"].nunique()}</div>Empresas</div>', unsafe_allow_html=True)

        st.divider()
        g1, g2 = st.columns(2)
        with g1:
            res_emp = df_f.groupby("empresa")["peso_kg"].sum().reset_index()
            st.plotly_chart(px.bar(res_emp, x="empresa", y="peso_kg", title="Peso por Empresa", color_discrete_map=COLORES_EMPRESA), use_container_width=True)
        with g2:
            st.plotly_chart(px.pie(df_f, values="peso_kg", names="tipo_residuo", title="Distribución"), use_container_width=True)

        st.dataframe(df_f, use_container_width=True, hide_index=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_f.to_excel(writer, index=False, sheet_name="MASTER")
        st.download_button("⬇️ Descargar Excel", output.getvalue(), "Reporte.xlsx", use_container_width=True)
