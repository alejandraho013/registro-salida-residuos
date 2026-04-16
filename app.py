import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re # Para validar el formato de la placa

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

st.title("🚚 Registro de salida de Residuos - TINTATEX")

# --- 1. DATOS DEL VEHÍCULO Y GESTOR ---
st.subheader("1. Datos del Transportador")
with st.container():
    c1, c2, c3 = st.columns(3)
    
    # Diccionario de gestores
    gestores_data = {
        "CORPOGESTAR": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Recicla Oriente": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
        "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"]),
        "Otro": []
    }
    
    fecha = c1.date_input("Fecha de Salida", datetime.now())
    empresa_sel = c1.selectbox("Empresa Gestora", options=list(gestores_data.keys()))
    
    # Campo manual si el gestor es "Otro"
    empresa_final = empresa_sel
    if empresa_sel == "Otro":
        empresa_final = c1.text_input("Escriba el nombre del Gestor")

    conductor = c2.text_input("Nombre del Conductor")
    placa = c3.text_input("Placa del Vehículo (Ej: ABC123)").upper().strip()

st.markdown("---")

# --- 2. CARGA DE RESIDUOS ---
st.subheader("2. Detalle de la Carga")
col_a, col_b, col_c = st.columns([2, 1, 1])

# Obtener lista de residuos según gestor seleccionado
opciones_residuos = gestores_data.get(empresa_sel, []).copy()
opciones_residuos.append("Otro")
tipo_sel = col_a.selectbox("Seleccione Residuo", options=opciones_residuos)

residuo_final = tipo_sel
if tipo_sel == "Otro":
    residuo_final = col_a.text_input("¿Cuál es el residuo?")

peso = col_b.number_input("Peso (kg)", min_value=0.0, step=0.1)

if col_c.button("➕ Agregar a la Lista"):
    # Validar formato de placa antes de permitir agregar a la lista
    es_placa_valida = re.match(r"^[A-Z]{3}[0-9]{3}$", placa)
    
    if not es_placa_valida:
        st.error("❌ Formato de placa incorrecto. Debe ser 3 letras y 3 números (Ej: ABC123).")
    elif peso <= 0:
        st.error("❌ El peso debe ser mayor a 0.")
    elif not residuo_final or not empresa_final:
        st.error("❌ Complete los campos de texto.")
    else:
        st.session_state.lista_temporal.append({"tipo_residuo": residuo_final, "peso_kg": peso})
        st.rerun()

# Tabla de resumen
if st.session_state.lista_temporal:
    st.markdown("#### 📋 Resumen Temporal")
    for index, item in enumerate(st.session_state.lista_temporal):
        cols = st.columns([3, 1, 1])
        cols[0].write(f"🔹 {item['tipo_residuo']}")
        cols[1].write(f"{item['peso_kg']} kg")
        if cols[2].button("🗑️ Quitar", key=f"btn_{index}"):
            st.session_state.lista_temporal.pop(index)
            st.rerun()
    
    suma_total = sum(item['peso_kg'] for item in st.session_state.lista_temporal)
    st.info(f"⚖️ **Peso Total Acumulado: {suma_total:,.1f} kg**")

st.markdown("---")

# --- 3. EVIDENCIAS Y ENVÍO FINAL ---
st.subheader("3. Evidencias (Obligatorias)")
f1, f2 = st.columns(2)

with f1:
    foto_memo = st.file_uploader("📄 Subir Foto del Memo", type=["jpg", "png", "jpeg"])
with f2:
    foto_camion = st.file_uploader("🚛 Subir Foto del Camión Lleno", type=["jpg", "png", "jpeg"])

novedades = st.text_area("Novedades u Observaciones (Opcional)")

if st.button("📤 GUARDAR REGISTRO"):
    # Validaciones de seguridad finales
    es_placa_valida = re.match(r"^[A-Z]{3}[0-9]{3}$", placa)
    
    if not st.session_state.lista_temporal:
        st.error("❌ Agregue al menos un residuo.")
    elif not es_placa_valida:
        st.error("❌ Corrija la placa (3 letras y 3 números).")
    elif not foto_memo or not foto_camion:
        st.error("❌ Ambas fotos son obligatorias para TINTATEX.")
    elif not conductor or not empresa_final:
        st.error("❌ Complete los datos del transportador.")
    else:
        with st.spinner("Sincronizando con GitHub..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # Subir fotos
                path_memo = f"fotos/MEMO_{ts}.jpg"
                repo.create_file(path_memo, f"Memo {ts}", foto_memo.getvalue())
                url_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{path_memo}"

                path_camion = f"fotos/CAMION_{ts}.jpg"
                repo.create_file(path_camion, f"Camion {ts}", foto_camion.getvalue())
                url_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{path_camion}"

                # Actualizar CSV
                contents = repo.get_contents("database.csv")
                db_txt = contents.decoded_content.decode("utf-8").strip()
                
                nov_final = novedades if novedades else "Sin observaciones"
                
                nuevas_filas = ""
                for item in st.session_state.lista_temporal:
                    nuevas_filas += f"\n{fecha},{empresa_final},{conductor},{placa},{item['tipo_residuo']},{item['peso_kg']},\"{nov_final}\",{url_memo},{url_camion}"
                
                repo.update_file("database.csv", f"Registro {placa} {ts}", db_txt + nuevas_filas, contents.sha)
                
                st.success(f"✅ ¡Despacho {placa} registrado con éxito!")
                st.balloons()
                st.session_state.lista_temporal = []
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Error técnico: {e}")
