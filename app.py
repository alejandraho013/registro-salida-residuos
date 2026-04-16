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

st.set_page_config(page_title="TINTATEX - Gestión de Residuos", page_icon="♻️", layout="centered")

# Mantener la lista de pesajes en la memoria del navegador
if 'lista_temporal' not in st.session_state:
    st.session_state.lista_temporal = []

st.title("♻️ Salida de Residuos TINTATEX")

# --- 1. DATOS DEL TRANSPORTADOR (COLAPSABLE PARA AHORRAR ESPACIO) ---
with st.expander("🚛 1. Datos del Vehículo y Gestor", expanded=True):
    c1, c2 = st.columns(2)
    gestores_data = {
        "CORPOGESTAR": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Recicla Oriente": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
        "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
        "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"]),
        "Otro": []
    }
    
    fecha = c1.date_input("Fecha", datetime.now())
    empresa_sel = c1.selectbox("Gestor", options=list(gestores_data.keys()))
    
    empresa_final = empresa_sel
    if empresa_sel == "Otro":
        empresa_final = c1.text_input("Nombre Gestor Manual")

    conductor = c2.text_input("Conductor")
    placa = c2.text_input("Placa (ABC123)").upper().strip()
    
    # Validación visual de placa
    placa_valida = False
    if placa:
        if re.match(r"^[A-Z]{3}[0-9]{3}$", placa):
            st.success("Placa Correcta")
            placa_valida = True
        else:
            st.error("Formato: 3 letras y 3 números")

st.markdown("---")

# --- 2. ENTRADA DE PESAJES (ESTO ES LO QUE MÁS USARÁS) ---
st.subheader("2. Ingreso de Pesos")
col_res, col_pes = st.columns([2, 1])

opciones_residuos = gestores_data.get(empresa_sel, []).copy()
opciones_residuos.append("Otro")
tipo_sel = col_res.selectbox("Tipo Residuo", options=opciones_residuos)

residuo_final = tipo_sel
if tipo_sel == "Otro":
    residuo_final = col_res.text_input("¿Cuál residuo?")

peso = col_pes.number_input("Peso (kg)", min_value=0.0, step=0.1, format="%.1f")

if st.button("➕ AGREGAR PESAJE", use_container_width=True):
    if not placa_valida:
        st.warning("⚠️ Primero corrige la placa arriba.")
    elif peso <= 0:
        st.warning("⚠️ El peso debe ser mayor a 0.")
    else:
        st.session_state.lista_temporal.append({"Residuo": residuo_final, "Kg": peso})
        st.toast(f"Agregado: {peso}kg")
        # No usamos rerun aquí para que el teclado del celular no se cierre siempre

# --- 3. RESUMEN SINTETIZADO (PARA VER LOS 20 PESAJES SIN OCUPAR ESPACIO) ---
if st.session_state.lista_temporal:
    suma_total = sum(item['Kg'] for item in st.session_state.lista_temporal)
    conteo = len(st.session_state.lista_temporal)
    
    # Métricas rápidas
    m1, m2 = st.columns(2)
    m1.metric("Total Acumulado", f"{suma_total:,.1f} kg")
    m2.metric("N° Pesajes", conteo)

    with st.expander("🔍 Ver detalle / Borrar registros", expanded=False):
        df_temp = pd.DataFrame(st.session_state.lista_temporal)
        # Tabla compacta tipo Excel
        st.dataframe(df_temp, use_container_width=True, hide_index=True)
        
        c_del1, c_del2 = st.columns(2)
        if c_del1.button("⏪ Borrar Último", use_container_width=True):
            st.session_state.lista_temporal.pop()
            st.rerun()
        if c_del2.button("🗑️ Borrar Todo", use_container_width=True):
            st.session_state.lista_temporal = []
            st.rerun()

st.markdown("---")

# --- 4. EVIDENCIAS Y CIERRE ---
st.subheader("3. Finalizar Despacho")
f1, f2 = st.columns(2)
foto_memo = f1.file_uploader("📄 Foto Memo (Opc)", type=["jpg", "png", "jpeg"])
foto_camion = f2.file_uploader("🚛 Foto Camión (Opc)", type=["jpg", "png", "jpeg"])
novedades = st.text_area("Observaciones")

if st.button("📤 ENVIAR A GITHUB / POWER BI", type="primary", use_container_width=True):
    if not st.session_state.lista_temporal:
        st.error("Lista vacía.")
    elif not placa_valida:
        st.error("Placa inválida.")
    else:
        with st.spinner("Sincronizando..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # Fotos
                u_memo = "Sin foto"; u_camion = "Sin foto"
                if foto_memo:
                    p_memo = f"fotos/MEMO_{ts}.jpg"
                    repo.create_file(p_memo, f"Memo {ts}", foto_memo.getvalue())
                    u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"
                if foto_camion:
                    p_camion = f"fotos/CAMION_{ts}.jpg"
                    repo.create_file(p_camion, f"Camion {ts}", foto_camion.getvalue())
                    u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                # CSV
                csv = repo.get_contents("database.csv")
                data_csv = csv.decoded_content.decode("utf-8").strip()
                nov_final = novedades if novedades else "Sin observaciones"
                
                nuevas_filas = ""
                for x in st.session_state.lista_temporal:
                    nuevas_filas += f"\n{fecha},{empresa_final},{conductor},{placa},{x['Residuo']},{x['Kg']},\"{nov_final}\",{u_memo},{u_camion}"
                
                repo.update_file("database.csv", f"Carga {placa} {ts}", data_csv + nuevas_filas, csv.sha)
                
                st.success("✅ ¡Despacho guardado!")
                st.balloons()
                st.session_state.lista_temporal = []
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
