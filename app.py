import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re

# --- CONFIGURACIÓN DE SEGURIDAD ---
try:
    TOKEN = st.secrets["TOKEN"]
    REPO_NAME = "alejandraho013/registro-salida-residuos" 
except Exception:
    st.error("⚠️ Error: Configura el TOKEN en 'Advanced Settings'.")
    st.stop()

st.set_page_config(page_title="TINTATEX - Gestión de Residuos", page_icon="♻️", layout="wide")

if 'lista_temporal' not in st.session_state:
    st.session_state.lista_temporal = []

st.title("🚚 Registro de salida de Residuos - TINTATEX")

# --- 1. DATOS DEL VEHÍCULO Y GESTOR ---
st.subheader("1. Datos del Transportador")
with st.container():
    c1, c2, c3 = st.columns(3)
    
    gestores_data = {
        "CORPOGESTAR": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Recicla Oriente": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
        "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"]),
        "Otro": []
    }
    
    fecha = c1.date_input("Fecha de Salida", datetime.now())
    empresa_sel = c1.selectbox("Empresa Gestora", options=list(gestores_data.keys()))
    
    empresa_final = empresa_sel
    if empresa_sel == "Otro":
        empresa_final = c1.text_input("Nombre del Gestor Manual")

    conductor = c2.text_input("Nombre del Conductor")
    
    # Campo de placa con validación inmediata
    placa = c3.text_input("Placa del Vehículo", placeholder="Ej: ABC123").upper().strip()
    
    # LÓGICA DE VALIDACIÓN DE PLACA EN TIEMPO REAL
    placa_valida = False
    if placa:
        if re.match(r"^[A-Z]{3}[0-9]{3}$", placa):
            st.success("✅ Formato de placa correcto")
            placa_valida = True
        else:
            st.error("❌ Formato incorrecto: Use 3 letras y 3 números (Ej: ABC123)")

st.markdown("---")

# --- 2. CARGA DE RESIDUOS ---
st.subheader("2. Detalle de la Carga")
col_a, col_b, col_c = st.columns([2, 1, 1])

opciones_residuos = gestores_data.get(empresa_sel, []).copy()
opciones_residuos.append("Otro")
tipo_sel = col_a.selectbox("Seleccione Residuo", options=opciones_residuos)

residuo_final = tipo_sel
if tipo_sel == "Otro":
    residuo_final = col_a.text_input("¿Cuál es el residuo?")

peso = col_b.number_input("Peso (kg)", min_value=0.0, step=0.1)

if col_c.button("➕ Agregar a la Lista"):
    if not placa_valida:
        st.error("Corrija la placa antes de agregar residuos.")
    elif peso <= 0:
        st.error("El peso debe ser mayor a 0.")
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

# --- 3. EVIDENCIAS (OPCIONALES) Y ENVÍO ---
st.subheader("3. Evidencias y Finalizar")
f1, f2 = st.columns(2)

with f1:
    foto_memo = st.file_uploader("📄 Foto del Memo", type=["jpg", "png", "jpeg"])
with f2:
    foto_camion = st.file_uploader("🚛 Foto Camión Lleno", type=["jpg", "png", "jpeg"])

novedades = st.text_area("Novedades u Observaciones")

if st.button("📤 GUARDAR REGISTRO"):
    if not st.session_state.lista_temporal:
        st.error("❌ Agregue al menos un residuo.")
    elif not placa_valida:
        st.error("❌ No se puede guardar con una placa inválida.")
    else:
        with st.spinner("Guardando en base de datos..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # Subir fotos solo si existen
                u_memo = "Sin foto"
                if foto_memo:
                    p_memo = f"fotos/MEMO_{ts}.jpg"
                    repo.create_file(p_memo, f"Memo {ts}", foto_memo.getvalue())
                    u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"

                u_camion = "Sin foto"
                if foto_camion:
                    p_camion = f"fotos/CAMION_{ts}.jpg"
                    repo.create_file(p_camion, f"Camion {ts}", foto_camion.getvalue())
                    u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                # Actualizar CSV
                contents = repo.get_contents("database.csv")
                db_txt = contents.decoded_content.decode("utf-8").strip()
                nov_final = novedades if novedades else "Sin observaciones"
                
                filas_nuevas = ""
                for item in st.session_state.lista_temporal:
                    filas_nuevas += f"\n{fecha},{empresa_final},{conductor},{placa},{item['tipo_residuo']},{item['peso_kg']},\"{nov_final}\",{u_memo},{u_camion}"
                
                repo.update_file("database.csv", f"Registro {placa} {ts}", db_txt + filas_nuevas, contents.sha)
                
                st.success(f"✅ ¡Despacho {placa} registrado con éxito!")
                st.balloons()
                st.session_state.lista_temporal = []
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Error técnico: {e}")
