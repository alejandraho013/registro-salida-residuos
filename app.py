import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re
import io

REPO_NAME = "alejandraho013/registro-salida-residuos"

GESTORES_DATA = {
    "CORPOGESTAR": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
    "Recicla Oriente": sorted(["Cartón limpio", "Cartón sucio", "Papel de archivo", "Pasta", "PET limpio", "PET sucio", "Plástico", "Retal de tela", "Tubo plega"]),
    "Quimetales NO Peligrosos": sorted(["Algodón", "Retal de tela", "Tubo plega"]),
    "Quimetales Peligrosos": sorted(["RAEE", "Residuos laboratorio", "Tela sucia"]),
    "Otro": []
}

try:
    TOKEN = st.secrets["TOKEN"]
    REPO_NAME = st.secrets.get("REPO_NAME", REPO_NAME)
except Exception:
    st.error("⚠️ Configura el TOKEN en 'Advanced Settings'.")
    st.stop()

st.set_page_config(page_title="TINTATEX - Gestión Dual", layout="wide")
st.title("Registro de Residuos TINTATEX")

if "lista_temporal" not in st.session_state:
    st.session_state.lista_temporal = []
if "envio_exitoso" not in st.session_state:
    st.session_state.envio_exitoso = False

@st.cache_data(ttl=300)
def cargar_datos():
    try:
        g = Github(TOKEN)
        repo = g.get_repo(REPO_NAME)
        xlsx_file = repo.get_contents("database.xlsx")
        df = pd.read_excel(io.BytesIO(xlsx_file.decoded_content), sheet_name="MASTER", engine="openpyxl")
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        return df
    except Exception:
        return pd.DataFrame()

# Función para crear el Excel con los resultados optimizados
def generar_excel_dual(df_detalle, df_resumen):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_detalle.to_excel(writer, sheet_name='Detalle_Filtrado', index=False)
        df_resumen.to_excel(writer, sheet_name='Resumen_Empresas', index=False)
    return output.getvalue()

tab1, tab2 = st.tabs(["📝 Registro", "📊 Reportes"])

# ─────────────────────────────────────────────────────────────
# TAB 1: REGISTRO (CON DOBLE ACTUALIZACIÓN GITHUB)
# ─────────────────────────────────────────────────────────────
with tab1:
    if st.session_state.envio_exitoso:
        st.success("¡Registro guardado en Excel y CSV correctamente! ✅")
        if st.button("Iniciar Nuevo Registro"):
            st.session_state.envio_exitoso = False
            st.session_state.lista_temporal = []
            cargar_datos.clear()
            st.rerun()
    else:
        with st.expander("🚛 1. Datos del Vehículo", expanded=True):
            c1, c2 = st.columns(2)
            fecha = c1.date_input("Fecha", datetime.now())
            empresa_sel = c1.selectbox("Empresa Gestora", options=list(GESTORES_DATA.keys()))
            empresa_final = c1.text_input("Nombre Manual").upper().strip() if empresa_sel == "Otro" else empresa_sel
            conductor = c2.text_input("Conductor")
            col_p, col_v = c2.columns([2,1])
            placa = col_p.text_input("Placa (ABC123)").upper().strip()
            placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa)) if placa else False
            if placa: 
                if placa_valida: col_v.success("✅")
                else: col_v.error("❌")

        st.subheader("2. Detalle de Pesajes")
        col_res, col_pes = st.columns([2, 1])
        opciones_res = GESTORES_DATA.get(empresa_sel, []).copy()
        opciones_res.append("Otro")
        tipo_sel = col_res.selectbox("Residuo", options=opciones_res)
        residuo_final = col_res.text_input("Especifique") if tipo_sel == "Otro" else tipo_sel
        peso = col_pes.number_input("Peso (kg)", min_value=0.0, step=0.1)

        if st.button("➕ AGREGAR PESAJE", use_container_width=True):
            if not placa_valida: st.error("Placa incorrecta.")
            elif peso <= 0: st.error("Peso inválido.")
            else:
                st.session_state.lista_temporal.append({"tipo_residuo": residuo_final, "peso_kg": peso})
                st.toast(f"✅ {residuo_final} agregado")

        if st.session_state.lista_temporal:
            df_temp = pd.DataFrame(st.session_state.lista_temporal)
            st.metric("Total Acumulado", f"{df_temp['peso_kg'].sum():,.1f} kg", f"{len(st.session_state.lista_temporal)} pesajes")
            with st.expander("🔍 Ver detalle temporal"):
                st.dataframe(df_temp, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("3. Finalizar")
        f1, f2 = st.columns(2)
        foto_memo = f1.file_uploader("Foto Memo", type=["jpg", "png", "jpeg"])
        foto_camion = f2.file_uploader("Foto Camión", type=["jpg", "png", "jpeg"])
        novedades = st.text_area("Observaciones")

        if st.button("📤 ENVIAR REGISTRO", type="primary", use_container_width=True):
            if not st.session_state.lista_temporal or not placa_valida:
                st.error("Verifique los datos.")
            else:
                with st.spinner("⏳ Procesando y clasificando la información..."):
                    try:
                        g = Github(TOKEN); repo = g.get_repo(REPO_NAME)
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        mes_actual = fecha.strftime("%B_%Y")
                        
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

                        # --- 1. ACTUALIZAR EXCEL (CON PESTAÑAS) ---
                        xlsx_file = repo.get_contents("database.xlsx")
                        diccionario_hojas = pd.read_excel(io.BytesIO(xlsx_file.decoded_content), sheet_name=None, engine="openpyxl")
                        
                        nov_final = novedades if novedades.strip() else "Sin observaciones"
                        nuevas_filas = [{"fecha": str(fecha), "mes": mes_actual, "empresa": empresa_final, "conductor": conductor, "placa": placa, "tipo_residuo": x["tipo_residuo"], "peso_kg": x["peso_kg"], "novedades": nov_final, "url_memo": u_memo, "url_camion": u_camion} for x in st.session_state.lista_temporal]
                        df_nuevos = pd.DataFrame(nuevas_filas)
                        
                        diccionario_hojas["MASTER"] = pd.concat([diccionario_hojas.get("MASTER", pd.DataFrame()), df_nuevos], ignore_index=True)
                        
                        nombre_hoja = re.sub(r'[\\/*?:\[\]]', "", empresa_final)[:30].strip().upper()
                        diccionario_hojas[nombre_hoja] = pd.concat([diccionario_hojas.get(nombre_hoja, pd.DataFrame()), df_nuevos], ignore_index=True)
                        
                        output_xlsx = io.BytesIO()
                        with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
                            for hoja, df_h in diccionario_hojas.items(): df_h.to_excel(writer, sheet_name=hoja, index=False)
                        repo.update_file("database.xlsx", f"Excel Update {ts}", output_xlsx.getvalue(), xlsx_file.sha)
                        
                        # --- 2. ACTUALIZAR CSV (ARCHIVO PLANO) ---
                        csv_file = repo.get_contents("database.csv")
                        csv_data = csv_file.decoded_content.decode("utf-8").strip()
                        for row in nuevas_filas:
                            csv_data += f"\n{row['fecha']},{row['mes']},{row['empresa']},{row['conductor']},{row['placa']},{row['tipo_residuo']},{row['peso_kg']},\"{row['novedades']}\",{row['url_memo']},{row['url_camion']}"
                        repo.update_file("database.csv", f"CSV Update {ts}", csv_data, csv_file.sha)
                        
                        st.session_state.envio_exitoso = True
                        st.rerun()
                    except Exception as e: st.error(f"Error técnico: {e}")

# ─────────────────────────────────────────────────────────────
# TAB 2: REPORTES (CON DOBLE FORMATO DE DESCARGA)
# ─────────────────────────────────────────────────────────────
with tab2:
    st.header("📊 Reportes y Estadísticas")
    df_master = cargar_datos()

    if df_master.empty:
        st.info("No hay datos.")
    else:
        with st.expander("🔍 Filtros", expanded=True):
            fc1, fc2, fc3 = st.columns(3)
            emp_f = fc1.selectbox("Empresa", ["Todas"] + sorted(df_master["empresa"].unique().tolist()))
            res_f = fc2.selectbox("Residuo", ["Todos"] + sorted(df_master["tipo_residuo"].unique().tolist()))
            mes_f = fc3.selectbox("Mes", ["Todos"] + sorted(df_master["mes"].unique().tolist()))

        df_f = df_master.copy()
        if emp_f != "Todas": df_f = df_f[df_f["empresa"] == emp_f]
        if res_f != "Todos": df_f = df_f[df_f["tipo_residuo"] == res_f]
        if mes_f != "Todos": df_f = df_f[df_f["mes"] == mes_f]

        # KPIs
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Registros", len(df_f))
        k2.metric("Total Kilos", f"{df_f['peso_kg'].sum():,.1f}")
        k3.metric("Gestores", df_f["empresa"].nunique())
        k4.metric("Promedio", f"{df_f['peso_kg'].mean():,.1f}" if not df_f.empty else "0")

        # Tabla y Resumen
        st.subheader("📋 Detalle de Movimientos")
        st.dataframe(df_f, use_container_width=True, hide_index=True)

        st.subheader("🏢 Resumen por Gestor")
        resumen_emp = df_f.groupby("empresa").agg(
            Viajes=("peso_kg", "count"),
            Kilos_Totales=("peso_kg", "sum"),
            Promedio_Kilos=("peso_kg", "mean")
        ).round(1).reset_index()
        st.dataframe(resumen_emp, use_container_width=True, hide_index=True)

        # ── EXPORTACIÓN DUAL ──
        st.markdown("---")
        st.subheader("💾 Descargar Resultados")
        d1, d2 = st.columns(2)
        
        # Botón CSV
        csv_b = df_f.to_csv(index=False).encode("utf-8")
        d1.download_button("⬇️ Descargar CSV (Plano)", csv_b, "datos.csv", "text/csv", use_container_width=True)
        
        # Botón Excel con Resumen
        xlsx_b = generar_excel_dual(df_f, resumen_emp)
        d2.download_button("⬇️ Descargar EXCEL (Detalle + Resumen)", xlsx_b, "reporte_completo.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
