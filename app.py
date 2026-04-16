import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import io

# --- CONFIGURACIÓN DE SEGURIDAD ---
try:
    TOKEN = st.secrets["TOKEN"]
    REPO_NAME = "alejandraho013/registro-salida-residuos" 
except Exception:
    st.error("⚠️ Error: Configura el TOKEN en 'Advanced Settings'.")
    st.stop()

st.set_page_config(page_title="TINTATEX - Gestión de Residuos", page_icon="♻️", layout="wide")

# --- FUNCIÓN PARA LEER DATOS ACTUALES ---
def cargar_datos():
    try:
        g = Github(TOKEN)
        repo = g.get_repo(REPO_NAME)
        contents = repo.get_contents("database.csv")
        df = pd.read_csv(io.StringIO(contents.decoded_content.decode("utf-8")))
        return df, contents.sha
    except:
        return pd.DataFrame(), None

# Cargar datos al inicio
df_existente, file_sha = cargar_datos()

st.title("🚚 Gestión de Residuos TINTATEX")

# --- SECCIÓN DE MÉTRICAS (SUMATORIAS) ---
if not df_existente.empty:
    st.subheader("📊 Resumen de Pesos Acumulados")
    col_total, col_gestor = st.columns([1, 2])
    
    with col_total:
        total_kg = df_existente['peso_kg'].sum()
        st.metric("Total General", f"{total_kg:,.1f} kg")
    
    with col_gestor:
        # Suma por empresa gestora
        resumen_empresa = df_existente.groupby('empresa')['peso_kg'].sum().reset_index()
        # Mostrar como una tablita rápida o mini gráfico
        st.dataframe(resumen_empresa, hide_index=True, use_container_width=True)
    st.markdown("---")

# --- CONFIGURACIÓN DE GESTORES Y RESIDUOS ---
residuos_corpogestar = sorted([
    "Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", 
    "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"
])

gestores_data = {
    "CORPOGESTAR": residuos_corpogestar,
    "Recicla Oriente": residuos_corpogestar,
    "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
    "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"])
}

# --- INTERFAZ DE REGISTRO ---
st.subheader("📝 Nuevo Registro de Salida")
col1, col2 = st.columns(2)

with col1:
    fecha = st.date_input("Fecha", datetime.now())
    empresa = st.selectbox("Empresa Gestora", options=list(gestores_data.keys()))
    conductor = st.text_input("Conductor")

with col2:
    placa = st.text_input("Placa")
    residuos_disponibles = gestores_data.get(empresa, []).copy()
    residuos_disponibles.append("Otro")
    tipo_residuo = st.selectbox("Tipo de Residuo", options=residuos_disponibles)
    
    residuo_manual = ""
    if tipo_residuo == "Otro":
        residuo_manual = st.text_input("Especifique residuo")
    
    peso = st.number_input("Peso (kg)", min_value=0.0, step=0.1)

novedades = st.text_area("Novedades")
foto = st.file_uploader("Evidencia Fotográfica", type=["jpg", "png", "jpeg"])

if st.button("🚀 Guardar y Actualizar Suma"):
    nombre_residuo_final = residuo_manual if tipo_residuo == "Otro" else tipo_residuo
    
    if not conductor or not placa or peso <= 0:
        st.warning("⚠️ Complete los campos obligatorios y peso > 0")
    else:
        with st.spinner("Actualizando registros..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                
                # 1. Foto
                url_foto = "Sin foto"
                if foto:
                    nombre_foto = f"fotos/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{placa}.jpg"
                    repo.create_file(nombre_foto, f"Foto {placa}", foto.getvalue())
                    url_foto = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{nombre_foto}"

                # 2. CSV
                contents = repo.get_contents("database.csv")
                db_actual = contents.decoded_content.decode("utf-8").strip()
                nueva_fila = f"\n{fecha},{empresa},{conductor},{placa},{nombre_residuo_final},{peso},\"{novedades}\",{url_foto}"
                repo.update_file("database.csv", f"Registro {placa}", db_actual + nueva_fila, contents.sha)
                
                st.success("✅ ¡Guardado! La suma se actualizará al recargar.")
                st.rerun() # Esto hace que la app se refresque y muestre la nueva suma arriba
                
            except Exception as e:
                st.error(f"❌ Error: {e}")
