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
    st.error("⚠️ Configura el TOKEN en 'Advanced Settings' de Streamlit Cloud.")
    st.stop()

st.set_page_config(page_title="TINTATEX - Gestión de Residuos", layout="centered")

# Inicialización de estados de sesión
if 'lista_temporal' not in st.session_state:
    st.session_state.lista_temporal = []
if 'envio_exitoso' not in st.session_state:
    st.session_state.envio_exitoso = False

st.title("Registro de salida de Residuos TINTATEX")

# --- PANTALLA DE CONFIRMACIÓN DE ENVÍO ---
if st.session_state.envio_exitoso:
    st.success("El registro se guardó y clasificó correctamente ✅")
    if st.button("Iniciar Nuevo Registro"):
        st.session_state.envio_exitoso = False
        st.session_state.lista_temporal = []
        st.rerun()
    st.stop()

# --- 1. DATOS DEL TRANSPORTADOR (COLAPSABLE) ---
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
    empresa_sel = c1.selectbox("Empresa Gestora", options=list(gestores_data.keys()))
    
    empresa_final = empresa_sel
    if empresa_sel == "Otro":
        empresa_final = c1.text_input("Escriba el nombre del Gestor")

    conductor = c2.text_input("Nombre del Conductor")
    placa = c2.text_input("Placa (ABC123)").upper().strip()
    
    # Validación de placa en tiempo real
    placa_valida = False
    if placa:
        if re.match(r"^[A-Z]{3}[0-9]{3}$", placa):
            st.success("✅ Formato de placa correcto")
            placa_valida = True
        else:
            st.error("❌ Formato incorrecto: Use 3 letras y 3 números (Ej: ABC123)")

st.markdown("---")

# --- 2. INGRESO DE PESAJES ---
st.subheader("2. Detalle de Pesajes")
col_res, col_pes = st.columns([2, 1])

opciones_residuos = gestores_data.get(empresa_sel, []).copy()
opciones_residuos.append("Otro")
tipo_sel = col_res.selectbox("Seleccione Residuo", options=opciones_residuos)

residuo_final = tipo_sel
if tipo_sel == "Otro":
    residuo_final = col_res.text_input("¿Cuál es el residuo?")

peso = col_pes.number_input("Peso (kg)", min_value=0.0, step=0.1, format="%.1f")

if st.button("➕ AGREGAR PESAJE", use_container_width=True):
    if not placa_valida:
        st.error("⚠️ Corrija la placa en la sección superior.")
    elif peso <= 0:
        st.error("⚠️ El peso debe ser mayor a 0.")
    else:
        st.session_state.lista_temporal.append({"tipo_residuo": residuo_final, "peso_kg": peso})
        st.toast(f"Agregado: {peso}kg")

# Resumen de métricas y tabla
if st.session_state.lista_temporal:
    df_temp = pd.DataFrame(st.session_state.lista_temporal)
    
    m1, m2 = st.columns(2)
    m1.metric("Total Acumulado", f"{df_temp['peso_kg'].sum():,.1f} kg")
    m2.metric("N° de Pesajes", len(st.session_state.lista_temporal))
    
    with st.expander("🔍 Ver / Editar Pesajes Registrados", expanded=False):
        st.dataframe(df_temp, use_container_width=True, hide_index=True)
        if st.button("⏪ Borrar último pesaje"):
            st.session_state.lista_temporal.pop()
            st.rerun()

st.markdown("---")

# --- 3. EVIDENCIAS Y ENVÍO ---
st.subheader("3. Evidencias y Cierre")
f1, f2 = st.columns(2)
foto_memo = f1.file_uploader("📄 Foto del Memo (Opcional)", type=["jpg", "png", "jpeg"])
foto_camion = f2.file_uploader("🚛 Foto del Camión (Opcional)", type=["jpg", "png", "jpeg"])
novedades = st.text_area("Novedades u Observaciones")

if st.button("📤 ENVIAR REGISTRO", type="primary", use_container_width=True):
    if not st.session_state.lista_temporal or not placa_valida:
        st.error("❌ Verifique la placa y que haya ingresado al menos un pesaje.")
    else:
        with st.spinner("⏳ Clasificando y sincronizando datos..."):
            try:
                g = Github(TOKEN)
                repo = g.get_repo(REPO_NAME)
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                mes_actual = fecha.strftime('%B_%Y')

                # Proceso de Fotos
                u_memo, u_camion = "Sin foto", "Sin foto"
                if foto_memo:
                    p_memo = f"fotos/MEMO_{ts}.jpg"
                    repo.create_file(p_memo, f"Memo {ts}", foto_memo.getvalue())
                    u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"
                if foto_camion:
                    p_camion = f"fotos/CAMION_{ts}.jpg"
                    repo.create_file(p_camion, f"Camion {ts}", foto_camion.getvalue())
                    u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                # --- ACTUALIZACIÓN DE EXCEL MULTI-PESTAÑA ---
                xlsx_content = repo.get_contents("database.xlsx")
                diccionario_hojas = pd.read_excel(io.BytesIO(xlsx_content.decoded_content), sheet_name=None, engine='openpyxl')
                
                nov_final = novedades if novedades else "Sin observaciones"
                nuevas_filas = []
                for x in st.session_state.lista_temporal:
                    nuevas_filas.append({
                        "fecha": str(fecha), "mes": mes_actual, "empresa": empresa_final, 
                        "conductor": conductor, "placa": placa, "tipo_residuo": x['tipo_residuo'], 
                        "peso_kg": x['peso_kg'], "novedades": nov_final, 
                        "url_memo": u_memo, "url_camion": u_camion
                    })
                df_nuevos = pd.DataFrame(nuevas_filas)

                # Actualizar MASTER
                if "MASTER" in diccionario_hojas:
                    diccionario_hojas["MASTER"] = pd.concat([diccionario_hojas["MASTER"], df_nuevos], ignore_index=True)
                else:
                    diccionario_hojas["MASTER"] = df_nuevos

                # Actualizar hoja por GESTOR (Limpiando nombre)
                nombre_gestor = "".join(re.findall(r"[\w\s]", empresa_final))[:30].strip().upper()
                if nombre_gestor in diccionario_hojas:
                    diccionario_hojas[nombre_gestor] = pd.concat([diccionario_hojas[nombre_gestor], df_nuevos], ignore_index=True)
                else:
                    diccionario_hojas[nombre_gestor] = df_nuevos

                # Guardar todas las pestañas de vuelta
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for nombre, df in diccionario_hojas.items():
                        df.to_excel(writer, sheet_name=nombre, index=False)
                
                repo.update_file("database.xlsx", f"Update Clasificado {placa}", output.getvalue(), xlsx_content.sha)
                
                # --- ACTUALIZAR CSV (Respaldo plano) ---
                csv_content = repo.get_contents("database.csv")
                csv_text = csv_content.decoded_content.decode("utf-8").strip()
                for row in nuevas_filas:
                    csv_text += f"\n{row['fecha']},{row['mes']},{row['empresa']},{row['conductor']},{row['placa']},{row['tipo_residuo']},{row['peso_kg']},\"{row['novedades']}\",{u_memo},{u_camion}"
                repo.update_file("database.csv", f"Update CSV {ts}", csv_text, csv_content.sha)

                st.session_state.envio_exitoso = True
                st.rerun()

            except Exception as e:
                st.error(f"❌ Error en la fuente de datos: {e}")
