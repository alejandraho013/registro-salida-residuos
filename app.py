import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re
import io

# --- CONFIGURACIÓN DE SEGURIDAD ---
try:
    TOKEN = st.secrets["TOKEN"]
    REPO_NAME = "alejandraho013/registro-salida-residuos" 
except Exception:
    st.error("⚠️ Configura el TOKEN en 'Advanced Settings'.")
    st.stop()

st.set_page_config(page_title="TINTATEX - Registro Dual", layout="centered")

if 'lista_temporal' not in st.session_state:
    st.session_state.lista_temporal = []
if 'envio_exitoso' not in st.session_state:
    st.session_state.envio_exitoso = False

st.title("Registro de salida de Residuos TINTATEX")

# --- MENSAJE DE ÉXITO ---
if st.session_state.envio_exitoso:
    st.success("¡Registro guardado exitosamente en CSV y Excel! ✅")
    if st.button("Nuevo registro"):
        st.session_state.envio_exitoso = False
        st.session_state.lista_temporal = []
        st.rerun()
    st.stop()

# --- 1. DATOS DEL TRANSPORTADOR ---
with st.expander("🚛 1. Datos del Vehículo", expanded=True):
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
    empresa_final = c1.text_input("Nombre Gestor Manual") if empresa_sel == "Otro" else empresa_sel
    conductor = c2.text_input("Conductor")
    placa = c2.text_input("Placa (ABC123)").upper().strip()
    
    placa_valida = False
    if placa:
        if re.match(r"^[A-Z]{3}[0-9]{3}$", placa):
            st.success("Placa Válida ✅")
            placa_valida = True
        else:
            st.error("Formato: 3 letras y 3 números")

st.markdown("---")

# --- 2. INGRESO DE PESAJES ---
st.subheader("2. Detalle de Pesajes")
col_res, col_pes = st.columns([2, 1])
opciones_res = gestores_data.get(empresa_sel, []).copy()
opciones_res.append("Otro")
tipo_sel = col_res.selectbox("Residuo", options=opciones_res)
residuo_final = col_res.text_input("¿Cuál residuo?") if tipo_sel == "Otro" else tipo_sel
peso = col_pes.number_input("Peso (kg)", min_value=0.0, step=0.1)

if st.button("➕ AGREGAR PESAJE", use_container_width=True):
    if not placa_valida: st.error("Corrija la placa antes de continuar.")
    elif peso <= 0: st.error("El peso debe ser mayor a 0.")
    else:
        st.session_state.lista_temporal.append({"tipo_residuo": residuo_final, "peso_kg": peso})
        st.toast(f"Agregado: {peso}kg")

# Resumen visual compacto
if st.session_state.lista_temporal:
    df_temp = pd.DataFrame(st.session_state.lista_temporal)
    st.metric("Total Acumulado", f"{df_temp['peso_kg'].sum():,.1f} kg")
    with st.expander("🔍 Ver detalle de pesajes", expanded=False):
        st.dataframe(df_temp, use_container_width=True, hide_index=True)
        if st.button("⏪ Borrar Último"):
            st.session_state.lista_temporal.pop()
            st.rerun()

st.markdown("---")

# --- 3. EVIDENCIAS Y ENVÍO DUAL ---
st.subheader("3. Evidencias y Cierre")
f1, f2 = st.columns(2)
foto_memo = f1.file_uploader("Foto Memo (Opc)", type=["jpg", "png", "jpeg"])
foto_camion = f2.file_uploader("Foto Camión (Opc)", type=["jpg", "png", "jpeg"])
novedades = st.text_area("Observaciones")

if st.button("📤 ENVIAR REGISTRO", type="primary", use_container_width=True):
    if not st.session_state.lista_temporal or not placa_valida:
        st.error("Faltan pesajes o la placa es incorrecta.")
    else:
        with st.spinner("Guardando en CSV y Excel..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # 1. Subir Fotos
                u_memo, u_camion = "Sin foto", "Sin foto"
                if foto_memo:
                    p_memo = f"fotos/MEMO_{ts}.jpg"
                    repo.create_file(p_memo, f"Memo {ts}", foto_memo.getvalue())
                    u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"
                if foto_camion:
                    p_camion = f"fotos/CAMION_{ts}.jpg"
                    repo.create_file(p_camion, f"Camion {ts}", foto_camion.getvalue())
                    u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                nov_final = novedades if novedades else "Sin observaciones"
                
                # --- 2. ACTUALIZAR CSV ---
                csv_file = repo.get_contents("database.csv")
                csv_data = csv_file.decoded_content.decode("utf-8").strip()
                for x in st.session_state.lista_temporal:
                    csv_data += f"\n{fecha},{empresa_final},{conductor},{placa},{x['tipo_residuo']},{x['peso_kg']},\"{nov_final}\",{u_memo},{u_camion}"
                repo.update_file("database.csv", f"Update CSV {ts}", csv_data, csv_file.sha)

                # --- 3. ACTUALIZAR EXCEL (.xlsx) ---
                xlsx_file = repo.get_contents("database.xlsx")
                df_old = pd.read_excel(io.BytesIO(xlsx_file.decoded_content))
                
                nuevas_filas = []
                for x in st.session_state.lista_temporal:
                    nuevas_filas.append({
                        "fecha": fecha, "empresa": empresa_final, "conductor": conductor,
                        "placa": placa, "tipo_residuo": x['tipo_residuo'], "peso_kg": x['peso_kg'],
                        "novedades": nov_final, "url_memo": u_memo, "url_camion": u_camion
                    })
                
                df_final = pd.concat([df_old, pd.DataFrame(nuevas_filas)], ignore_index=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                
                repo.update_file("database.xlsx", f"Update XLSX {ts}", output.getvalue(), xlsx_file.sha)
                
                st.session_state.envio_exitoso = True
                st.rerun()

            except Exception as e:
                st.error(f"Error en la sincronización: {e}")
