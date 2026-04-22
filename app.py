import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re
import io
import plotly.express as px

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN Y FACTORES AMBIENTALES
# ─────────────────────────────────────────────────────────────
REPO_NAME = "alejandraho013/registro-salida-residuos"

GESTORES_DATA = {
    "CORPOGESTAR":             sorted(["Cartón limpio","Cartón sucio","Papel de archivo","Pasta","PET limpio","PET sucio","Plástico","Retal de tela","Tubo plega"]),
    "Recicla Oriente":          sorted(["Cartón limpio","Cartón sucio","Papel de archivo","Pasta","PET limpio","PET sucio","Plástico","Retal de tela","Tubo plega"]),
    "Quimetales NO Peligrosos":sorted(["Algodón","Retal de tela","Tubo plega"]),
    "Quimetales Peligrosos":   sorted(["RAEE","Residuos laboratorio","Tela sucia"]),
    "Otro": [],
}

# Factores de emisión (kg CO2 evitados por cada kg reciclado) - Valores promedio aproximados
FACTORES_CO2 = {
    "Cartón": 0.9, "Papel": 0.9, "Plástico": 1.5, "PET": 1.2, "Pasta": 1.0, 
    "Retal de tela": 0.5, "Algodón": 0.4, "Tubo plega": 0.8
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
st.set_page_config(page_title="TINTATEX · Gestión Ambiental", layout="wide")

st.markdown("""
<style>
.main-header { background: linear-gradient(135deg,#004d40 0%,#00796b 100%); padding:1.5rem; border-radius:12px; color:white; margin-bottom:1.5rem; }
.kpi-card { background:white; border-radius:10px; padding:1rem; border-left:4px solid #00796b; box-shadow:0 2px 5px rgba(0,0,0,.1); text-align:center; }
.kpi-value { font-size:1.8rem; font-weight:700; color:#004d40; }
</style>
<div class="main-header"><h1>🏭 TINTATEX · Gestión Ambiental de Residuos</h1><p>Registro de Pesajes y Huella de Carbono</p></div>
""", unsafe_allow_html=True)

if "lista_temporal" not in st.session_state: st.session_state.lista_temporal = []
if "envio_exitoso" not in st.session_state: st.session_state.envio_exitoso = False

# ─────────────────────────────────────────────────────────────
# FUNCIONES AUXILIARES
# ─────────────────────────────────────────────────────────────
def calcular_co2(residuo, peso):
    for clave, factor in FACTORES_CO2.items():
        if clave.lower() in residuo.lower():
            return peso * factor
    return 0.0

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
tab1, tab2 = st.tabs(["📝 Registro", "📊 Reportes de Impacto"])

with tab1:
    if st.session_state.envio_exitoso:
        st.success("✅ ¡Registro guardado y reporte de impacto actualizado!")
        if st.button("🔄 Nuevo Registro", type="primary"):
            st.session_state.envio_exitoso = False
            st.session_state.lista_temporal = []
            st.rerun()
    else:
        with st.expander("🚛 1. Información del Transporte", expanded=True):
            c1, c2 = st.columns(2)
            fecha = c1.date_input("Fecha", datetime.now())
            emp_sel = c1.selectbox("Empresa Gestora", options=list(GESTORES_DATA.keys()))
            empresa_final = c1.text_input("Nombre Manual").upper() if emp_sel == "Otro" else emp_sel
            
            conductor = c2.text_input("Nombre Conductor")
            cp1, cp2 = c2.columns([2, 1])
            placa = cp1.text_input("Placa (ABC123)").upper().strip()
            placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa))
            if placa:
                if placa_valida: cp2.success("OK")
                else: cp2.error("Error")
            
            # --- NUEVA FUNCIÓN: Capacidad del camión ---
            capacidad = c2.number_input("Capacidad Camión (kg) - Opcional", min_value=0.0, step=100.0)

        st.subheader("⚖️ 2. Pesajes y Cálculo de Impacto")
        col_r, col_p, col_b = st.columns([3,2,1])
        res_opts = GESTORES_DATA.get(emp_sel, []) + ["Otro"]
        tipo_res = col_r.selectbox("Tipo de Residuo", options=res_opts)
        res_final = col_r.text_input("Especifique") if tipo_res == "Otro" else tipo_res
        peso = col_p.number_input("Peso (kg)", min_value=0.0, step=0.1)
        
        if col_b.button("➕ Añadir Pesaje", use_container_width=True):
            if placa_valida and peso > 0:
                st.session_state.lista_temporal.append({"tipo_residuo": res_final, "peso_kg": peso})
                st.toast(f"Añadido {peso} kg")
            else: st.error("Verifique placa y peso")

        if st.session_state.lista_temporal:
            df_temp = pd.DataFrame(st.session_state.lista_temporal)
            total_actual = df_temp['peso_kg'].sum()
            
            # Alerta de capacidad
            if capacidad > 0 and total_actual > capacidad:
                st.warning(f"⚠️ Carga actual ({total_actual:,.1f} kg) supera la capacidad informada ({capacidad:,.1f} kg).")
            
            # KPIs Temporales
            t1, t2, t3 = st.columns(3)
            t1.metric("Suma Carga", f"{total_actual:,.1f} kg")
            t2.metric("N° Pesajes", len(df_temp))
            # Cálculo de CO2 para la lista temporal
            co2_temp = sum([calcular_co2(x["tipo_residuo"], x["peso_kg"]) for x in st.session_state.lista_temporal])
            t3.metric("CO2 Evitado Est.", f"{co2_temp:,.1f} kg")
            
            st.dataframe(df_temp, use_container_width=True, hide_index=True)
            if st.button("⏪ Eliminar último"):
                st.session_state.lista_temporal.pop()
                st.rerun()

        novedades = st.text_area("Observaciones")
        if st.button("📤 ENVIAR REGISTRO", type="primary", use_container_width=True):
            if not st.session_state.lista_temporal or not placa_valida:
                st.error("Datos incompletos.")
            else:
                with st.spinner("💾 Sincronizando con base de datos ambiental..."):
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
    if df.empty: st.info("Esperando primer registro...")
    else:
        # Cálculo de columna de CO2 en el histórico
        df["co2_evitado"] = df.apply(lambda x: calcular_co2(x["tipo_residuo"], x["peso_kg"]), axis=1)
        
        f_mes = st.selectbox("Mes de reporte", ["Todos"] + sorted(df["mes"].unique().tolist()))
        df_f = df if f_mes == "Todos" else df[df["mes"] == f_mes]

        # KPIs de Impacto Ambiental
        k1, k2, k3 = st.columns(3)
        k1.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["peso_kg"].sum():,.1f} kg</div>Peso Total</div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["co2_evitado"].sum():,.1f} kg</div>CO2 Evitado Est.</div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-card"><div class="kpi-value">{df_f["empresa"].nunique()}</div>Gestores Activos</div>', unsafe_allow_html=True)

        # Gráficas
        st.divider()
        g1, g2 = st.columns(2)
        with g1:
            res_co2 = df_f.groupby("empresa")["co2_evitado"].sum().reset_index()
            fig = px.bar(res_co2, x="empresa", y="co2_evitado", title="🌱 CO2 Evitado por Gestor (kg)", text_auto=".1f")
            st.plotly_chart(fig, use_container_width=True)
        with g2:
            fig_p = px.pie(df_f, values="peso_kg", names="tipo_residuo", title="♻️ Distribución de Carga", hole=0.4)
            st.plotly_chart(fig_p, use_container_width=True)

        st.subheader("📋 Registro Histórico")
        st.dataframe(df_f, use_container_width=True, hide_index=True)
