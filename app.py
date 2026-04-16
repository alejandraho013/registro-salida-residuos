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

# --- 1. DATOS DEL VEHÍCULO Y GESTOR ---
st.subheader("1. Datos del Transportador")
with st.container():
    c1, c2, c3 = st.columns(3)
    gestores_data = {
        "CORPOGESTAR": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Recicla Oriente": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
        "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"])
    }
    
    fecha = c1.date_input("Fecha de Salida", datetime.now())
    empresa = c1.selectbox("Empresa Gestora", options=list(gestores_data.keys()))
    conductor = c2.text_input("Nombre del Conductor")
    placa = c3.text_input("Placa del Vehículo").upper()

st.markdown("---")

# --- 2. CARGA DE RESIDUOS ---
st.subheader("2. Detalle de la Carga")
col_a, col_b, col_c = st.columns([2, 1, 1])

lista_residuos = gestores_data.get(empresa, []).copy()
lista_residuos.append("Otro")
tipo = col_a.selectbox("Seleccione Residuo", options=lista_residuos)

# Lógica para campo "Otro"
residuo_nombre = tipo
if tipo == "Otro":
    residuo_nombre = col_a.text_input("¿Cuál es el residuo?")

peso = col_b.number_input("Peso (kg)", min_value=0.0, step=0.1)

if col_c.button("➕ Agregar a la Lista"):
    if peso > 0 and residuo_nombre and placa:
        nuevo_item = {
            "tipo_residuo": residuo_nombre,
            "peso_kg": peso
        }
        st.session_state.lista_temporal.append(nuevo_item)
        st.rerun()
    else:
        st.warning("⚠️ Complete Placa, Peso y Residuo.")

# Mostrar tabla interactiva
if st.session_state.lista_temporal:
    st.markdown("#### 📋 Resumen Temporal")
    for index, item in enumerate(st.session_state.lista_temporal):
        cols = st.columns([3, 1, 1])
        cols[0].write(f"🔹 {item['tipo_residuo']}")
        cols[1].write(f"{item['peso_kg']} kg")
        if cols[2].button("🗑️ Quitar", key=f"btn_{index}"):
            st.session_state.lista_temporal.pop(index)
            st.rerun()
    
    suma_actual = sum(item['peso_kg'] for item in st.session_state.lista_temporal)
    st.info(f"⚖️ **Peso Total Acumulado: {suma_actual:,.1f} kg**")

st.markdown("---")

# --- 3. EVIDENCIAS Y ENVÍO FINAL ---
st.subheader("3. Evidencias y Cierre")
f1, f2 = st.columns(2)

with f1:
    foto_memo = st.file_uploader("📄 Subir Foto del Memo", type=["jpg", "png", "jpeg"])
with f2:
    foto_camion = st.file_uploader("🚛 Subir Foto del Camión Lleno", type=["jpg", "png", "jpeg"])

novedades = st.text_area("Novedades u Observaciones")

if st.button("📤 FINALIZAR Y GUARDAR REGISTRO"):
    if not st.session_state.lista_temporal:
        st.error("❌ La lista de carga está vacía.")
    elif not foto_memo or not foto_camion:
        st.error("❌ Es obligatorio subir ambas fotos para el despacho.")
    else:
        with st.spinner("Sincronizando con GitHub..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # --- SUBIR FOTOS ---
                # Memo
                path_memo = f"fotos/MEMO_{ts}.jpg"
                repo.create_file(path_memo, f"Memo {ts}", foto_memo.getvalue())
                url_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{path_memo}"

                # Camion
                path_camion = f"fotos/CAMION_{ts}.jpg"
                repo.create_file(path_camion, f"Camion {ts}", foto_camion.getvalue())
                url_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{path_camion}"

                # --- ACTUALIZAR CSV ---
                contents = repo.get_contents("database.csv")
                db_txt = contents.decoded_content.decode("utf-8").strip()
                
                nuevas_filas = ""
                for item in st.session_state.lista_temporal:
                    # Orden: fecha,empresa,conductor,placa,tipo_residuo,peso_kg,novedades,url_memo,url_camion
                    nuevas_filas += f"\n{fecha},{empresa},{conductor},{placa},{item['tipo_residuo']},{item['peso_kg']},\"{novedades}\",{url_memo},{url_camion}"
                
                repo.update_file("database.csv", f"Despacho {placa} - {ts}", db_txt + nuevas_filas, contents.sha)
                
                st.success(f"✅ ¡Despacho de {placa} guardado con éxito!")
                st.balloons()
                st.session_state.lista_temporal = [] # Limpiar lista
                
            except Exception as e:
                st.error(f"❌ Error al conectar con GitHub: {e}")
