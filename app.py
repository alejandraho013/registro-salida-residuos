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

st.title("🚚 Registro de Despacho - TINTATEX")

# --- DATOS DEL GESTOR ---
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

residuo_nombre = tipo
if tipo == "Otro":
    residuo_nombre = col_a.text_input("¿Cuál residuo?")

peso = col_b.number_input("Peso (kg)", min_value=0.0, step=0.1)

if col_c.button("➕ Agregar a la lista"):
    if peso > 0 and residuo_nombre and placa:
        nuevo_item = {
            "fecha": fecha,
            "empresa": empresa,
            "conductor": conductor,
            "placa": placa.upper(),
            "tipo_residuo": residuo_nombre,
            "peso_kg": peso
        }
        st.session_state.lista_temporal.append(nuevo_item)
        st.toast(f"Agregado: {residuo_nombre}")
    else:
        st.warning("Ingrese placa, peso y nombre del residuo")

# --- MOSTRAR TABLA Y SUMA ---
if st.session_state.lista_temporal:
    df_temp = pd.DataFrame(st.session_state.lista_temporal)
    st.markdown("### 📋 Resumen de Carga Actual")
    st.table(df_temp[["tipo_residuo", "peso_kg"]])
    
    suma_actual = df_temp["peso_kg"].sum()
    st.info(f"⚖️ **Suma Total del despacho: {suma_actual:,.1f} kg**")
    
    if st.button("🗑️ Borrar lista"):
        st.session_state.lista_temporal = []
        st.rerun()

st.markdown("---")

# --- SECCIÓN DE FOTOS Y FINALIZACIÓN ---
st.subheader("3. Evidencias y Finalizar")
f1, f2 = st.columns(2)

with f1:
    foto_memorando = st.file_uploader("📄 Foto del Memorando", type=["jpg", "png", "jpeg"])
with f2:
    foto_camion = st.file_uploader("🚛 Foto Camión Lleno (Que se vea la placa)", type=["jpg", "png", "jpeg"])

novedades = st.text_area("Novedades finales")

if st.button("📤 ENVIAR REGISTRO"):
    if not st.session_state.lista_temporal:
        st.error("Debe agregar al menos un residuo a la lista.")
    elif not foto_memorando or not foto_camion:
        st.warning("⚠️ Se recomienda subir ambas fotos (Memorando y Camión) para el registro.")
        # Si prefieres que sea OBLIGATORIO, cambia el st.warning por st.error y pon un st.stop()
    
    else:
        with st.spinner("Guardando registro y fotos en GitHub..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # 1. Subir Foto Memorando
                url_memo = "Sin foto"
                if foto_memorando:
                    path_memo = f"fotos/MEMO_{timestamp}_{placa.upper()}.jpg"
                    repo.create_file(path_memo, f"Memo {placa}", foto_memorando.getvalue())
                    url_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{path_memo}"

                # 2. Subir Foto Camión
                url_camion = "Sin foto"
                if foto_camion:
                    path_camion = f"fotos/CAMION_{timestamp}_{placa.upper()}.jpg"
                    repo.create_file(path_camion, f"Camion {placa}", foto_camion.getvalue())
                    url_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{path_camion}"

                # 3. Actualizar CSV (Incluyendo las dos URLs de fotos)
                contents = repo.get_contents("database.csv")
                db_txt = contents.decoded_content.decode("utf-8").strip()
                
                nuevas_filas = ""
                for item in st.session_state.lista_temporal:
                    # Formato: fecha, empresa, conductor, placa, residuo, peso, novedades, url_memo, url_camion
                    nuevas_filas += f"\n{item['fecha']},{item['empresa']},{item['conductor']},{item['placa']},{item['tipo_residuo']},{item['peso_kg']},\"{novedades}\",{url_memo},{url_camion}"
                
                repo.update_file("database.csv", f"Despacho {placa}", db_txt + nuevas_filas, contents.sha)
                
                st.success(f"✅ ¡Despacho guardado! Total: {suma_actual} kg.")
                st.balloons()
                st.session_state.lista_temporal = []
                
            except Exception as e:
                st.error(f"Error: {e}")
