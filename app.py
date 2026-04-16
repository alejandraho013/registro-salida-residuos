import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime

# --- CONFIGURACIÓN DE SEGURIDAD ---
# Usamos st.secrets para no exponer el Token en el código público/privado
try:
    TOKEN = st.secrets["TOKEN"]
    # AJUSTA EL NOMBRE DEL REPO AQUÍ SI ES DIFERENTE:
    REPO_NAME = "alejandraho013/registro-salida-residuos" 
except Exception:
    st.error("⚠️ Error: No se encontró el TOKEN. Configúralo en 'Advanced Settings' de Streamlit.")
    st.stop()

# Configuración de la página
st.set_page_config(page_title="TINTATEX - Registro de Residuos", page_icon="♻️")

st.title("🚚 Registro de Salida de Residuos")
st.markdown("---")

# Formulario de entrada de datos
with st.form("form_registro", clear_on_submit=True):
    col1, col2 = st.columns(2)
    
    with col1:
        fecha = st.date_input("Fecha de Registro", datetime.now())
        empresa = st.text_input("Empresa Gestora (Recolector)")
        conductor = st.text_input("Nombre del Conductor")
        
    with col2:
        placa = st.text_input("Placa del Vehículo")
        tipo_residuo = st.selectbox("Tipo de Residuo", 
                                   ["Lodos PTARI", "Retal Textil", "Plástico", "Cartón/Papel", "Peligrosos", "Ordinarios"])
        peso = st.number_input("Peso Total (kg)", min_value=0.0, step=0.1)

    novedades = st.text_area("Novedades u Observaciones")
    foto = st.file_uploader("Evidencia Fotográfica (Carro cargado)", type=["jpg", "png", "jpeg"])

    boton_enviar = st.form_submit_button("🚀 Guardar Registro")

# Lógica de guardado
if boton_enviar:
    if not empresa or not conductor or not placa or peso == 0:
        st.warning("⚠️ Por favor complete todos los campos obligatorios y asegúrese de que el peso sea mayor a 0.")
    else:
        with st.spinner("Subiendo datos a la base de datos de TINTATEX..."):
            try:
                # 1. Conexión con GitHub
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                
                # 2. Procesar Foto si existe
                url_foto = "Sin foto"
                if foto is not None:
                    # Nombre único basado en tiempo y placa
                    nombre_archivo_foto = f"fotos/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{placa}.jpg"
                    repo.create_file(nombre_archivo_foto, f"Foto evidencia {placa}", foto.getvalue())
                    # Link directo para visualizar en Power BI
                    url_foto = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{nombre_archivo_foto}"

                # 3. Actualizar base de datos (CSV)
                contents = repo.get_contents("database.csv")
                # Decodificar y limpiar espacios
                db_actual = contents.decoded_content.decode("utf-8").strip()
                
                # Crear nueva fila (usamos comillas para evitar errores con comas en las novedades)
                nueva_fila = f"\n{fecha},{empresa},{conductor},{placa},{tipo_residuo},{peso},\"{novedades}\",{url_foto}"
                db_actualizada = db_actual + nueva_fila
                
                # Subir cambio a GitHub
                repo.update_file("database.csv", f"Registro {placa} - {datetime.now()}", db_actualizada, contents.sha)
                
                st.success(f"✅ ¡Registro de la placa {placa} guardado con éxito!")
                st.balloons()
                
            except Exception as e:
                st.error(f"❌ Error al conectar con GitHub: {e}")
