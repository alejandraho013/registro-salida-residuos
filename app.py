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

# --- FUNCIÓN PARA LEER DATOS Y CALCULAR SUMAS ---
def cargar_y_sumar():
    try:
        g = Github(TOKEN)
        repo = g.get_repo(REPO_NAME)
        contents = repo.get_contents("database.csv")
        # Leemos el CSV directamente desde GitHub
        df = pd.read_csv(io.StringIO(contents.decoded_content.decode("utf-8")))
        return df
    except Exception as e:
        # Si el archivo está vacío o no existe, devolvemos un DataFrame vacío
        return pd.DataFrame()

# 1. Cargamos los datos actuales
df_datos = cargar_y_sumar()

st.title("🚚 Sistema de Gestión de Residuos - TINTATEX")

# --- VISUALIZACIÓN DE LA SUMA EN LA APP ---
st.markdown("### 📊 Totales Acumulados en Planta")

if not df_datos.empty:
    # Calculamos la suma total
    total_general = df_datos['peso_kg'].sum()
    
    # Creamos columnas para las métricas
    m1, m2, m3 = st.columns(3)
    
    with m1:
        st.metric(label="Total General Enviado", value=f"{total_general:,.1f} kg", delta="Acumulado")
    
    with m2:
        # Suma específica para CORPOGESTAR / Recicla Oriente
        suma_corpo = df_datos[df_datos['empresa'].isin(['CORPOGESTAR', 'Recicla Oriente'])]['peso_kg'].sum()
        st.metric(label="CORPO / Recicla Oriente", value=f"{suma_corpo:,.1f} kg")
        
    with m3:
        # Suma específica para Quimetales
        suma_qui = df_datos[df_datos['empresa'].str.contains('Quimetales')]['peso_kg'].sum()
        st.metric(label="Quimetales (P y No P)", value=f"{suma_qui:,.1f} kg")

    # Tabla de resumen detallado por Residuo
    with st.expander("Ver detalle de sumas por tipo de residuo"):
        resumen_tipo = df_datos.groupby('tipo_residuo')['peso_kg'].sum().reset_index()
        st.table(resumen_tipo)
else:
    st.info("Aún no hay registros. La suma aparecerá aquí apenas guardes el primer dato.")

st.markdown("---")

# --- CONFIGURACIÓN DE GESTORES Y RESIDUOS ---
residuos_corpogestar = sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"])

gestores_data = {
    "CORPOGESTAR": residuos_corpogestar,
    "Recicla Oriente": residuos_corpogestar,
    "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
    "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"])
}

# --- FORMULARIO DE REGISTRO ---
st.subheader("📝 Nuevo Registro de Salida")
c1, c2 = st.columns(2)

with c1:
    fecha = st.date_input("Fecha", datetime.now())
    empresa = st.selectbox("Empresa Gestora", options=list(gestores_data.keys()))
    conductor = st.text_input("Nombre del Conductor")

with c2:
    placa = st.text_input("Placa del Vehículo")
    lista_residuos = gestores_data.get(empresa, []).copy()
    lista_residuos.append("Otro")
    tipo = st.selectbox("Tipo de Residuo", options=lista_residuos)
    
    res_manual = ""
    if tipo == "Otro":
        res_manual = st.text_input("¿Cuál residuo?")
        
    peso = st.number_input("Peso registrado (kg)", min_value=0.0, step=0.1)

nov = st.text_area("Novedades")
foto = st.file_uploader("Evidencia Fotográfica", type=["jpg", "png", "jpeg"])

if st.button("🚀 Guardar y Actualizar Sumatoria"):
    residuo_final = res_manual if tipo == "Otro" else tipo
    if not conductor or not placa or peso <= 0:
        st.warning("⚠️ Complete los campos obligatorios.")
    else:
        with st.spinner("Guardando y recalculando totales..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                
                # Subir Foto
                url_foto = "Sin foto"
                if foto:
                    path = f"fotos/{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                    repo.create_file(path, "Foto residuo", foto.getvalue())
                    url_foto = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{path}"

                # Actualizar CSV
                contents = repo.get_contents("database.csv")
                db_txt = contents.decoded_content.decode("utf-8").strip()
                nueva_fila = f"\n{fecha},{empresa},{conductor},{placa},{residuo_final},{peso},\"{nov}\",{url_foto}"
                repo.update_file("database.csv", f"Registro {placa}", db_txt + nueva_fila, contents.sha)
                
                st.success("✅ ¡Registro guardado!")
                st.rerun() # Esto hace que la app se recargue y la suma de arriba cambie inmediatamente
                
            except Exception as e:
                st.error(f"Error: {e}")
