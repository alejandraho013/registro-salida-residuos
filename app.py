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

# --- ESTADO DE LA SESIÓN ---
if 'lista_temporal' not in st.session_state:
    st.session_state.lista_temporal = []

st.title("🚚 Registro de Despacho - TINTATEX")

# --- 1. DATOS DEL GESTOR ---
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
    placa = c3.text_input("Placa del Vehículo").upper()

st.markdown("---")

# --- 2. AGREGAR RESIDUOS ---
st.subheader("2. Cargar Residuos al Camión")
col_a, col_b, col_c = st.columns([2, 1, 1])

lista_residuos = gestores_data.get(empresa, []).copy()
lista_residuos.append("Otro")
tipo = col_a.selectbox("Tipo de Residuo", options=lista_residuos)

residuo_nombre = tipo
if tipo == "Otro":
    residuo_nombre = col_a.text_input("¿Cuál residuo?")

peso = col_b.number_input("Peso (kg)", min_value=0.0, step=0.1)

if col_c.button("➕ Agregar a la lista"):
    if peso > 0 and residuo_nombre and placa:
        nuevo_item = {
            "tipo_residuo": residuo_nombre,
            "peso_kg": peso
        }
        st.session_state.lista_temporal.append(nuevo_item)
        st.rerun()
    else:
        st.warning("Complete placa, peso y residuo")

# --- MOSTRAR TABLA CON OPCIÓN DE BORRAR INDIVIDUAL ---
if st.session_state.lista_temporal:
    st.markdown("### 📋 Resumen de Carga Actual")
    
    # Creamos una tabla con un botón de eliminar para cada fila
    for index, item in enumerate(st.session_state.lista_temporal):
        cols = st.columns([3, 1, 1])
        cols[0].write(f"**{item['tipo_residuo']}**")
        cols[1].write(f"{item['peso_kg']} kg")
        if cols[2].button("🗑️ Quitar", key=f"btn_{index}"):
            st.session_state.lista_temporal.pop(index)
            st.rerun()
    
    suma_actual = sum(item['peso_kg'] for item in st.session_state.lista_temporal)
    st.info(f"⚖️ **Suma Total del despacho: {suma_actual:,.1f} kg**")

st.markdown("---")

# --- 3. EVIDENCIAS Y ENVÍO ---
st.subheader("3. Evidencias y Finalizar")
f1, f2 = st.columns(2)

with f1:
    foto_memo = st.file_uploader("📄 Foto del Memo", type=["jpg", "png", "jpeg"])
with f2:
    foto_camion = st.file_uploader("🚛 Foto Camión Lleno (Placa visible)", type=["jpg", "png", "jpeg"])

novedades = st.text_area("Novedades finales")

if st.button("📤 ENVIAR DESPACHO COMPLETO"):
    if not st.session_state.lista_temporal:
        st.error("La lista está vacía.")
    else:
        with st.spinner("Guardando en GitHub..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # Subir Fotos
                u_memo = "Sin foto"
                if foto_memo:
                    p_memo = f"fotos/MEMO_{ts}_{placa}.jpg"
                    repo.create_file(p_memo, f"Memo {placa}", foto_memo.getvalue())
                    u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"

                u_camion = "Sin foto"
                if foto_camion:
                    p_camion = f"fotos/CAMION_{ts}_{placa}.jpg"
                    repo.create_file(p_camion, f"Camion {placa}", foto_camion.getvalue())
                    u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                # Actualizar CSV
                contents = repo.get_contents("database.csv")
                db_txt = contents.decoded_content.decode("utf-8").strip()
                
                filas_nuevas = ""
                for item in st.session_state.lista_temporal:
                    filas_nuevas += f"\n{fecha},{empresa},{conductor},{placa},{item['tipo_residuo']},{item['peso_kg']},\"{novedades}\",{u_memo},{u_camion}"
                
                repo.update_file("database.csv", f"Despacho {placa}", db_txt + filas_nuevas, contents.sha)
                
                st.success("✅ Despacho guardado exitosamente.")
                st.balloons()
                st.session_state.lista_temporal = []
                st.rerun()
                
            except Exception as e:
                st.error(f"Error: {e}")
