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

st.set_page_config(page_title="TINTATEX - Gestión de Residuos", layout="centered")

if 'lista_temporal' not in st.session_state:
    st.session_state.lista_temporal = []
if 'envio_exitoso' not in st.session_state:
    st.session_state.envio_exitoso = False

st.title("♻️ Registro de Residuos TINTATEX")

# --- CONFIRMACIÓN DE ENVÍO ---
if st.session_state.envio_exitoso:
    st.success("¡Registro guardado y pestaña creada/actualizada correctamente! ✅")
    if st.button("Iniciar Nuevo Registro"):
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
    empresa_sel = c1.selectbox("Empresa Gestora", options=list(gestores_data.keys()))
    
    empresa_final = empresa_sel
    if empresa_sel == "Otro":
        empresa_final = c1.text_input("Nombre del Gestor Manual").upper().strip()

    conductor = c2.text_input("Conductor")
    placa = c2.text_input("Placa (ABC123)").upper().strip()
    placa_valida = re.match(r"^[A-Z]{3}[0-9]{3}$", placa) if placa else False

st.markdown("---")

# --- 2. INGRESO DE PESAJES ---
st.subheader("2. Detalle de Pesajes")
col_res, col_pes = st.columns([2, 1])
opciones_res = gestores_data.get(empresa_sel, []).copy()
opciones_res.append("Otro")
tipo_sel = col_res.selectbox("Residuo", options=opciones_res)
residuo_final = col_res.text_input("¿Cuál?") if tipo_sel == "Otro" else tipo_sel
peso = col_pes.number_input("Peso (kg)", min_value=0.0, step=0.1)

if st.button("➕ AGREGAR PESAJE", use_container_width=True):
    if not placa_valida: st.error("Corrija la placa.")
    elif peso <= 0: st.error("Peso inválido.")
    else:
        st.session_state.lista_temporal.append({"tipo_residuo": residuo_final, "peso_kg": peso})
        st.toast(f"Agregado: {peso}kg")

if st.session_state.lista_temporal:
    df_temp = pd.DataFrame(st.session_state.lista_temporal)
    st.metric("Total Acumulado", f"{df_temp['peso_kg'].sum():,.1f} kg", f"{len(st.session_state.lista_temporal)} pesajes")
    with st.expander("🔍 Ver detalle", expanded=False):
        st.dataframe(df_temp, use_container_width=True, hide_index=True)

st.markdown("---")

# --- 3. EVIDENCIAS Y ENVÍO AUTOMÁTICO ---
st.subheader("3. Finalizar y Clasificar")
f1, f2 = st.columns(2)
foto_memo = f1.file_uploader("Foto Memo (Opc)", type=["jpg", "png", "jpeg"])
foto_camion = f2.file_uploader("Foto Camión (Opc)", type=["jpg", "png", "jpeg"])
novedades = st.text_area("Observaciones")

if st.button("📤 ENVIAR REGISTRO", type="primary", use_container_width=True):
    if not st.session_state.lista_temporal or not placa_valida:
        st.error("Datos incompletos.")
    else:
        with st.spinner("Creando pestañas y guardando..."):
            try:
                g = Github(TOKEN); repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                mes_actual = fecha.strftime('%B_%Y')

                # Fotos
                u_memo, u_camion = "Sin foto", "Sin foto"
                if foto_memo:
                    p_memo = f"fotos/MEMO_{ts}.jpg"
                    repo.create_file(p_memo, f"Memo {ts}", foto_memo.getvalue())
                    u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"
                if foto_camion:
                    p_camion = f"fotos/CAMION_{ts}.jpg"
                    repo.create_file(p_camion, f"Camion {ts}", foto_camion.getvalue())
                    u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                # --- LÓGICA DE EXCEL ---
                xlsx_file = repo.get_contents("database.xlsx")
                diccionario_hojas = pd.read_excel(io.BytesIO(xlsx_file.decoded_content), sheet_name=None, engine='openpyxl')
                
                nov_final = novedades if novedades else "Sin observaciones"
                nuevas_datas = []
                for x in st.session_state.lista_temporal:
                    nuevas_datas.append({
                        "fecha": str(fecha), "mes": mes_actual, "empresa": empresa_final, 
                        "conductor": conductor, "placa": placa, "tipo_residuo": x['tipo_residuo'], 
                        "peso_kg": x['peso_kg'], "novedades": nov_final, 
                        "url_memo": u_memo, "url_camion": u_camion
                    })
                df_nuevos = pd.DataFrame(nuevas_datas)

                # 1. Actualizar MASTER
                if "MASTER" in diccionario_hojas:
                    diccionario_hojas["MASTER"] = pd.concat([diccionario_hojas["MASTER"], df_nuevos], ignore_index=True)
                else:
                    diccionario_hojas["MASTER"] = df_nuevos

                # 2. Lógica para crear pestañas automáticas
                # Limpiamos el nombre para que Excel no se bloquee (máx 31 caracteres, sin / \ ? * [ ])
                nombre_hoja = re.sub(r'[\\/*?:\[\]]', '', empresa_final)[:30].strip().upper()
                if not nombre_hoja: nombre_hoja = "OTROS"

                if nombre_hoja in diccionario_hojas:
                    diccionario_hojas[nombre_hoja] = pd.concat([diccionario_hojas[nombre_hoja], df_nuevos], ignore_index=True)
                else:
                    diccionario_hojas[nombre_hoja] = df_nuevos

                # Guardar todo
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for hoja, df_hoja in diccionario_hojas.items():
                        df_hoja.to_excel(writer, sheet_name=hoja, index=False)
                
                repo.update_file("database.xlsx", f"Update {placa}", output.getvalue(), xlsx_file.sha)
                
                st.session_state.envio_exitoso = True
                st.rerun()

            except Exception as e:
                st.error(f"Error técnico: {e}")
