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

st.set_page_config(page_title="TINTATEX - Fuente Organizada", layout="centered")

if 'lista_temporal' not in st.session_state:
    st.session_state.lista_temporal = []
if 'envio_exitoso' not in st.session_state:
    st.session_state.envio_exitoso = False

st.title("Registro salida de Residuos TINTATEX")

# --- MENSAJE DE ÉXITO ---
if st.session_state.envio_exitoso:
    st.success("¡Registro guardado y clasificado por gestor exitosamente! ✅")
    if st.button("Iniciar Nuevo Camión"):
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
        if st.button("⏪ Borrar Último"):
            st.session_state.lista_temporal.pop(); st.rerun()

st.markdown("---")

# --- 3. EVIDENCIAS Y ENVÍO CLASIFICADO ---
st.subheader("3. Finalizar y Clasificar")
f1, f2 = st.columns(2)
foto_memo = f1.file_uploader("Foto Memo (Opc)", type=["jpg", "png", "jpeg"])
foto_camion = f2.file_uploader("Foto Camión (Opc)", type=["jpg", "png", "jpeg"])
novedades = st.text_area("Observaciones")

if st.button("📤 ENVIAR REGISTRO", type="primary", use_container_width=True):
    if not st.session_state.lista_temporal or not placa_valida:
        st.error("Datos incompletos.")
    else:
        with st.spinner("Clasificando datos en Excel..."):
            try:
                g = Github(TOKEN); repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                mes_actual = fecha.strftime('%B_%Y') # Ejemplo: April_2026

                # Manejo de Fotos
                u_memo, u_camion = "Sin foto", "Sin foto"
                if foto_memo:
                    p_memo = f"fotos/MEMO_{ts}.jpg"
                    repo.create_file(p_memo, f"Memo {ts}", foto_memo.getvalue())
                    u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"
                if foto_camion:
                    p_camion = f"fotos/CAMION_{ts}.jpg"
                    repo.create_file(p_camion, f"Camion {ts}", foto_camion.getvalue())
                    u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                # --- PROCESO DE EXCEL MULTI-PESTAÑA ---
                xlsx_file = repo.get_contents("database.xlsx")
                # Cargamos todas las pestañas existentes
                diccionario_hojas = pd.read_excel(io.BytesIO(xlsx_file.decoded_content), sheet_name=None, engine='openpyxl')
                
                # Preparamos los nuevos datos
                nov_final = novedades if novedades else "Sin observaciones"
                nuevas_datas = []
                for x in st.session_state.lista_temporal:
                    nuevas_datas.append({
                        "fecha": fecha, "mes": mes_actual, "empresa": empresa_final, 
                        "conductor": conductor, "placa": placa, "tipo_residuo": x['tipo_residuo'], 
                        "peso_kg": x['peso_kg'], "novedades": nov_final, 
                        "url_memo": u_memo, "url_camion": u_camion
                    })
                df_nuevos = pd.DataFrame(nuevas_datas)

                # 1. Actualizar Hoja MASTER
                if "MASTER" in diccionario_hojas:
                    diccionario_hojas["MASTER"] = pd.concat([diccionario_hojas["MASTER"], df_nuevos], ignore_index=True)
                else:
                    diccionario_hojas["MASTER"] = df_nuevos

                # 2. Actualizar Hoja por GESTOR
                nombre_hoja_gestor = "".join(re.findall(r"[\w\s]", empresa_final))[:30] # Limpiar nombre para Excel
                if nombre_hoja_gestor in diccionario_hojas:
                    diccionario_hojas[nombre_hoja_gestor] = pd.concat([diccionario_hojas[nombre_hoja_gestor], df_nuevos], ignore_index=True)
                else:
                    diccionario_hojas[nombre_hoja_gestor] = df_nuevos

                # Guardar todas las hojas de vuelta al Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for nombre_hoja, df_hoja in diccionario_hojas.items():
                        df_hoja.to_excel(writer, sheet_name=nombre_hoja, index=False)
                
                repo.update_file("database.xlsx", f"Update Clasificado {placa}", output.getvalue(), xlsx_file.sha)
                
                # --- ACTUALIZAR CSV (Opcional, como respaldo plano) ---
                csv_file = repo.get_contents("database.csv")
                csv_data = csv_file.decoded_content.decode("utf-8").strip()
                for row in nuevas_datas:
                    csv_data += f"\n{row['fecha']},{row['empresa']},{row['conductor']},{row['placa']},{row['tipo_residuo']},{row['peso_kg']},\"{row['novedades']}\",{u_memo},{u_camion}"
                repo.update_file("database.csv", f"Update CSV {ts}", csv_data, csv_file.sha)

                st.session_state.envio_exitoso = True
                st.rerun()

            except Exception as e:
                st.error(f"Error en la fuente: {e}")
