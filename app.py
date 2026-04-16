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

# --- ESTADO DE LA SESIÓN (Para la lista temporal) ---
if 'lista_temporal' not in st.session_state:
    st.session_state.lista_temporal = []

st.title("🚚 Registro salida residuos - TINTATEX")

# --- DATOS DEL GESTOR (Se llenan una sola vez por camión) ---
st.subheader("1. Datos del Vehículo y Gestor")
with st.expander("Configurar datos del transportador", expanded=True):
    c1, c2, c3 = st.columns(3)
    gestores_data = {
        "CORPOGESTAR": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Recicla Oriente": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
        "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"])
    }
    
    fecha = c1.date_input("Fecha", datetime.now())
    empresa = c1.selectbox("Empresa Gestora", options=list(gestores_data.keys()))
    conductor = c2.text_input("Nombre del Conductor")
    placa = c3.text_input("Placa del Vehículo")

st.markdown("---")

# --- AGREGAR RESIDUOS UNO POR UNO ---
st.subheader("2. Cargar Residuos al Camión")
col_a, col_b, col_c = st.columns([2, 1, 1])

lista_residuos = gestores_data.get(empresa, []).copy()
lista_residuos.append("Otro")
tipo = col_a.selectbox("Tipo de Residuo", options=lista_residuos)

# Manejo de "Otro"
residuo_nombre = tipo
if tipo == "Otro":
    residuo_nombre = col_a.text_input("¿Cuál residuo?")

peso = col_b.number_input("Peso (kg)", min_value=0.0, step=0.1)

if col_c.button("➕ Agregar a la lista"):
    if peso > 0 and residuo_nombre:
        nuevo_item = {
            "fecha": fecha,
            "empresa": empresa,
            "conductor": conductor,
            "placa": placa,
            "tipo_residuo": residuo_nombre,
            "peso_kg": peso
        }
        st.session_state.lista_temporal.append(nuevo_item)
        st.toast(f"Agregado: {residuo_nombre}")
    else:
        st.warning("Ingrese peso y nombre del residuo")

# --- MOSTRAR TABLA TEMPORAL Y SUMA ---
if st.session_state.lista_temporal:
    df_temp = pd.DataFrame(st.session_state.lista_temporal)
    
    st.markdown("### 📋 Resumen de Carga Actual")
    st.dataframe(df_temp[["tipo_residuo", "peso_kg"]], use_container_width=True, hide_index=True)
    
    suma_actual = df_temp["peso_kg"].sum()
    st.info(f"⚖️ **Suma Total de este despacho: {suma_actual:,.1f} kg**")
    
    if st.button("🗑️ Borrar lista y empezar de nuevo"):
        st.session_state.lista_temporal = []
        st.rerun()

st.markdown("---")

# --- FINALIZAR Y ENVIAR A GITHUB ---
st.subheader("3. Finalizar Registro")
novedades = st.text_area("Novedades finales del despacho")
foto = st.file_uploader("Foto de evidencia del camión cargado", type=["jpg", "png", "jpeg"])

if st.button("📤 ENVIAR TODO A GITHUB"):
    if not st.session_state.lista_temporal:
        st.error("La lista está vacía. Agregue residuos primero.")
    elif not conductor or not placa:
        st.error("Faltan datos del conductor o placa.")
    else:
        with st.spinner("Guardando despacho completo..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                
                # 1. Subir Foto
                url_foto = "Sin foto"
                if foto:
                    path = f"fotos/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{placa}.jpg"
                    repo.create_file(path, f"Evidencia {placa}", foto.getvalue())
                    url_foto = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{path}"

                # 2. Leer CSV actual
                contents = repo.get_contents("database.csv")
                db_txt = contents.decoded_content.decode("utf-8").strip()
                
                # 3. Preparar todas las filas nuevas
                nuevas_filas = ""
                for item in st.session_state.lista_temporal:
                    nuevas_filas += f"\n{item['fecha']},{item['empresa']},{item['conductor']},{item['placa']},{item['tipo_residuo']},{item['peso_kg']},\"{novedades}\",{url_foto}"
                
                # 4. Guardar todo de un solo golpe
                repo.update_file("database.csv", f"Despacho {placa}", db_txt + nuevas_filas, contents.sha)
                
                st.success(f"✅ ¡Despacho de {suma_actual} kg guardado exitosamente!")
                st.balloons()
                st.session_state.lista_temporal = [] # Limpiar para el siguiente camión
                
            except Exception as e:
                st.error(f"Error al guardar: {e}")
