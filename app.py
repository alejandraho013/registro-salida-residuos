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

# --- DICCIONARIO DE GESTORES Y RESIDUOS ---
# Definimos qué maneja cada uno en orden alfabético
gestores_data = {
    "CORPOGESTAR": sorted([
        "Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", 
        "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"
    ]),
    "Quimetales NO Peligrosos": sorted([
        "Algodón", "Retal de tela"
    ]),
    "Quimetales Peligrosos": sorted([
        "RAEE", "Residuos laboratorio", "Tela sucia"
    ]),
    "Recicla Oriente": [] # Puedes dejarlo vacío o añadir sus residuos aquí
}

# Formulario
with st.form("form_registro", clear_on_submit=True):
    col1, col2 = st.columns(2)
    
    with col1:
        fecha = st.date_input("Fecha de Registro", datetime.now())
        # Selector de Empresa Gestora
        empresa = st.selectbox("Empresa Gestora (Recolector)", options=list(gestores_data.keys()))
        conductor = st.text_input("Nombre del Conductor")
        
    with col2:
        placa = st.text_input("Placa del Vehículo")
        
        # Lógica para mostrar residuos según el gestor
        residuos_disponibles = gestores_data.get(empresa, [])
        residuos_disponibles.append("Otro") # Siempre añadimos la opción "Otro"
        
        tipo_residuo = st.selectbox("Tipo de Residuo", options=residuos_disponibles)
        
        # Campo manual si elige "Otro"
        residuo_manual = ""
        if tipo_residuo == "Otro":
            residuo_manual = st.text_input("Especifique el residuo (Manual)")
            
        peso = st.number_input("Peso Total (kg)", min_value=0.0, step=0.1)

    novedades = st.text_area("Novedades u Observaciones")
    foto = st.file_uploader("Evidencia Fotográfica", type=["jpg", "png", "jpeg"])

    boton_enviar = st.form_submit_button("🚀 Guardar Registro")

# Lógica de guardado
if boton_enviar:
    # Definir el nombre final del residuo
    residuo_final = residuo_manual if tipo_residuo == "Otro" else tipo_residuo
    
    if not empresa or not conductor or not placa or (tipo_residuo == "Otro" and not residuo_manual):
        st.warning("⚠️ Complete todos los campos, incluyendo el nombre del residuo manual si seleccionó 'Otro'.")
    else:
        with st.spinner("Guardando en base de datos..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                
                # Procesar Foto
                url_foto = "Sin foto"
                if foto:
                    nombre_foto = f"fotos/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{placa}.jpg"
                    repo.create_file(nombre_foto, f"Evidencia {placa}", foto.getvalue())
                    url_foto = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{nombre_foto}"

                # Actualizar CSV
                contents = repo.get_contents("database.csv")
                db_actual = contents.decoded_content.decode("utf-8").strip()
                nueva_fila = f"\n{fecha},{empresa},{conductor},{placa},{residuo_final},{peso},\"{novedades}\",{url_foto}"
                repo.update_file("database.csv", f"Registro {placa}", db_actual + nueva_fila, contents.sha)
                
                st.success(f"✅ ¡Éxito! Registro de {empresa} guardado.")
                st.balloons()
            except Exception as e:
                st.error(f"❌ Error: {e}")
