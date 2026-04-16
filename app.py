import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime

# --- CONFIGURACIÓN DE SEGURIDAD ---
try:
    TOKEN = st.secrets["TOKEN"]
    REPO_NAME = "alejandraho013/registro-salida-residuos" 
except Exception:
    st.error("⚠️ Error: Configura el TOKEN en 'Advanced Settings'.")
    st.stop()

st.set_page_config(page_title="TINTATEX - Gestión de Residuos", page_icon="♻️")

st.title("🚚 Registro de Salida de Residuos")
st.markdown("---")

# --- DICCIONARIO DE GESTORES Y RESIDUOS ACTUALIZADO ---
# Definimos la lista base de Corpogestar para reutilizarla
residuos_corpogestar = sorted([
    "Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", 
    "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"
])

gestores_data = {
    "CORPOGESTAR": residuos_corpogestar,
    "Recicla Oriente": residuos_corpogestar,  # Son los mismos que Corpogestar
    "Quimetales NO Peligrosos": sorted([
        "Algodón", "Retal de tela", "Tubo plega"  # Agregado Tubo plega
    ]),
    "Quimetales Peligrosos": sorted([
        "RAEE", "Residuos laboratorio", "Tela sucia"
    ])
}

# --- INTERFAZ DINÁMICA ---
col1, col2 = st.columns(2)

with col1:
    fecha = st.date_input("Fecha de Registro", datetime.now())
    empresa = st.selectbox("Empresa Gestora (Recolector)", options=list(gestores_data.keys()))
    conductor = st.text_input("Nombre del Conductor")

with col2:
    placa = st.text_input("Placa del Vehículo")
    
    # Obtener lista según gestor y añadir "Otro"
    residuos_disponibles = gestores_data.get(empresa, []).copy()
    residuos_disponibles.append("Otro")
    
    tipo_residuo = st.selectbox("Tipo de Residuo", options=residuos_disponibles)
    
    # Campo manual si se selecciona "Otro"
    residuo_manual = ""
    if tipo_residuo == "Otro":
        residuo_manual = st.text_input("Escriba el nombre del residuo manualmente")
    
    peso = st.number_input("Peso Total (kg)", min_value=0.0, step=0.1)

novedades = st.text_area("Novedades u Observaciones")
foto = st.file_uploader("Evidencia Fotográfica", type=["jpg", "png", "jpeg"])

# Botón de envío
if st.button("🚀 Guardar Registro"):
    # Definir el nombre final que se guardará en el Excel
    nombre_residuo_final = residuo_manual if tipo_residuo == "Otro" else tipo_residuo
    
    if not empresa or not conductor or not placa or (tipo_residuo == "Otro" and not residuo_manual) or peso <= 0:
        st.warning("⚠️ Por favor, complete todos los campos obligatorios y asegúrese de que el peso sea mayor a 0.")
    else:
        with st.spinner("Guardando registro en GitHub..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                
                # 1. Procesar Foto
                url_foto = "Sin foto"
                if foto:
                    nombre_foto = f"fotos/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{placa}.jpg"
                    repo.create_file(nombre_foto, f"Evidencia {placa}", foto.getvalue())
                    url_foto = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{nombre_foto}"

                # 2. Actualizar CSV
                contents = repo.get_contents("database.csv")
                db_actual = contents.decoded_content.decode("utf-8").strip()
                
                # Formato de fila: fecha, empresa, conductor, placa, residuo, peso, novedades, url_foto
                nueva_fila = f"\n{fecha},{empresa},{conductor},{placa},{nombre_residuo_final},{peso},\"{novedades}\",{url_foto}"
                repo.update_file("database.csv", f"Registro {placa} - {datetime.now().date()}", db_actual + nueva_fila, contents.sha)
                
                st.success(f"✅ ¡Registro guardado exitosamente!")
                st.balloons()
                
            except Exception as e:
                st.error(f"❌ Error al guardar: {e}")
