import streamlit as st
import pandas as pd
from github import Github
from datetime import datetime
import re
import io
import plotly.graph_objects as go
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

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
st.title("📦 Registro de Residuos TINTATEX")

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

# Función mejorada para crear Excel con formato profesional y gráficos
def generar_excel_optimizado(df_detalle, df_resumen_empresa, df_resumen_residuo):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja 1: Detalle
        df_detalle.to_excel(writer, sheet_name='Detalle', index=False)
        
        # Hoja 2: Resumen por Gestor
        df_resumen_empresa.to_excel(writer, sheet_name='Resumen_Gestores', index=False)
        
        # Hoja 3: Resumen por Residuo
        df_resumen_residuo.to_excel(writer, sheet_name='Resumen_Residuos', index=False)
    
    # Aplicar estilos
    wb = output
    # (Los estilos se aplicarían aquí con openpyxl)
    
    return output.getvalue()

tab1, tab2 = st.tabs(["📝 Registro", "📊 Reportes Avanzados"])

# ─────────────────────────────────────────────────────────────
# TAB 1: REGISTRO
# ���────────────────────────────────────────────────────────────
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
            
            # KPIs visuales
            m1, m2, m3 = st.columns(3)
            m1.metric("📦 Total Acumulado", f"{df_temp['peso_kg'].sum():,.1f} kg")
            m2.metric("📍 Pesajes", len(st.session_state.lista_temporal))
            m3.metric("⚖️ Promedio", f"{df_temp['peso_kg'].mean():,.1f} kg")
            
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

                        # Actualizar EXCEL
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
                            for hoja, df_h in diccionario_hojas.items(): 
                                df_h.to_excel(writer, sheet_name=hoja, index=False)
                        repo.update_file("database.xlsx", f"Excel Update {ts}", output_xlsx.getvalue(), xlsx_file.sha)
                        
                        # Actualizar CSV
                        csv_file = repo.get_contents("database.csv")
                        csv_data = csv_file.decoded_content.decode("utf-8").strip()
                        for row in nuevas_filas:
                            csv_data += f"\n{row['fecha']},{row['mes']},{row['empresa']},{row['conductor']},{row['placa']},{row['tipo_residuo']},{row['peso_kg']},\"{row['novedades']}\",{row['url_memo']},{row['url_camion']}"
                        repo.update_file("database.csv", f"CSV Update {ts}", csv_data, csv_file.sha)
                        
                        st.session_state.envio_exitoso = True
                        st.rerun()
                    except Exception as e: 
                        st.error(f"Error técnico: {e}")

# ─────────────────────────────────────────────────────────────
# TAB 2: REPORTES AVANZADOS CON VISUALIZACIONES
# ─────────────────────────────────────────────────────────────
with tab2:
    st.header("📊 Reportes y Estadísticas Avanzadas")
    df_master = cargar_datos()

    if df_master.empty:
        st.info("No hay datos para mostrar.")
    else:
        # Filtros en contenedor expandible
        with st.expander("🔍 Filtros Avanzados", expanded=True):
            fc1, fc2, fc3, fc4 = st.columns(4)
            emp_f = fc1.selectbox("Empresa", ["Todas"] + sorted(df_master["empresa"].unique().tolist()))
            res_f = fc2.selectbox("Residuo", ["Todos"] + sorted(df_master["tipo_residuo"].unique().tolist()))
            mes_f = fc3.selectbox("Mes", ["Todos"] + sorted(df_master["mes"].unique().tolist()))
            conductor_f = fc4.selectbox("Conductor", ["Todos"] + sorted(df_master["conductor"].unique().tolist()))

        # Aplicar filtros
        df_f = df_master.copy()
        if emp_f != "Todas": 
            df_f = df_f[df_f["empresa"] == emp_f]
        if res_f != "Todos": 
            df_f = df_f[df_f["tipo_residuo"] == res_f]
        if mes_f != "Todos": 
            df_f = df_f[df_f["mes"] == mes_f]
        if conductor_f != "Todos": 
            df_f = df_f[df_f["conductor"] == conductor_f]

        # KPIs Principales
        st.markdown("### 📈 Indicadores Clave")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("📊 Registros", len(df_f), delta=f"{len(df_f)} movimientos")
        k2.metric("⚖️ Total Kilos", f"{df_f['peso_kg'].sum():,.1f}", delta=f"+{df_f['peso_kg'].sum():,.1f} kg")
        k3.metric("🏢 Gestores", df_f["empresa"].nunique())
        k4.metric("🚚 Conductores", df_f["conductor"].nunique())
        k5.metric("📦 Promedio", f"{df_f['peso_kg'].mean():,.1f} kg/registro" if not df_f.empty else "0")

        # Tabs de visualización
        viz_tab1, viz_tab2, viz_tab3 = st.tabs(["📋 Tabla Detallada", "📊 Gráficos", "💾 Descargas"])

        with viz_tab1:
            st.subheader("Detalle de Movimientos")
            st.dataframe(df_f, use_container_width=True, hide_index=True)

        with viz_tab2:
            col_g1, col_g2 = st.columns(2)
            
            # Gráfico 1: Peso por Gestor
            with col_g1:
                resumen_emp = df_f.groupby("empresa").agg({
                    "peso_kg": ["sum", "count", "mean"]
                }).round(1).reset_index()
                resumen_emp.columns = ["empresa", "total_kg", "registros", "promedio_kg"]
                
                fig1 = px.bar(resumen_emp, x="empresa", y="total_kg", title="📦 Kilos por Gestor",
                             labels={"total_kg": "Kilos", "empresa": "Gestor"},
                             color="total_kg", color_continuous_scale="Viridis")
                st.plotly_chart(fig1, use_container_width=True)

            # Gráfico 2: Peso por Tipo de Residuo
            with col_g2:
                resumen_res = df_f.groupby("tipo_residuo").agg({
                    "peso_kg": ["sum", "count"]
                }).round(1).reset_index()
                resumen_res.columns = ["tipo_residuo", "total_kg", "registros"]
                
                fig2 = px.pie(resumen_res, values="total_kg", names="tipo_residuo", 
                             title="♻️ Distribución por Tipo de Residuo")
                st.plotly_chart(fig2, use_container_width=True)

            col_g3, col_g4 = st.columns(2)
            
            # Gráfico 3: Evolución temporal
            with col_g3:
                df_f_sorted = df_f.sort_values("fecha")
                evolucion = df_f_sorted.groupby(df_f_sorted["fecha"].dt.date).agg({
                    "peso_kg": "sum"
                }).reset_index()
                
                fig3 = px.line(evolucion, x="fecha", y="peso_kg", 
                              title="📅 Evolución de Kilos por Día",
                              labels={"peso_kg": "Kilos", "fecha": "Fecha"},
                              markers=True)
                st.plotly_chart(fig3, use_container_width=True)

            # Gráfico 4: Top Conductores
            with col_g4:
                top_conductores = df_f.groupby("conductor").agg({
                    "peso_kg": "sum"
                }).reset_index().sort_values("peso_kg", ascending=False).head(10)
                
                fig4 = px.barh(top_conductores, x="peso_kg", y="conductor",
                              title="👨‍✈️ Top 10 Conductores",
                              labels={"peso_kg": "Kilos", "conductor": "Conductor"})
                st.plotly_chart(fig4, use_container_width=True)

        with viz_tab3:
            st.subheader("💾 Descargar Reportes")
            
            # Generar resúmenes
            resumen_empresa = df_f.groupby("empresa").agg(
                Viajes=("peso_kg", "count"),
                Kilos_Totales=("peso_kg", "sum"),
                Promedio_Kilos=("peso_kg", "mean"),
                Conductores=("conductor", "nunique")
            ).round(1).reset_index().sort_values("Kilos_Totales", ascending=False)
            
            resumen_residuo = df_f.groupby("tipo_residuo").agg(
                Registros=("peso_kg", "count"),
                Kilos_Totales=("peso_kg", "sum"),
                Promedio_Kilos=("peso_kg", "mean")
            ).round(1).reset_index().sort_values("Kilos_Totales", ascending=False)
            
            d1, d2, d3 = st.columns(3)
            
            # Descargar CSV
            csv_b = df_f.to_csv(index=False).encode("utf-8")
            d1.download_button("📥 CSV (Detallado)", csv_b, "detalle_movimientos.csv", 
                             "text/csv", use_container_width=True)
            
            # Descargar Excel
            xlsx_b = generar_excel_optimizado(df_f, resumen_empresa, resumen_residuo)
            d2.download_button("📥 EXCEL (Completo)", xlsx_b, "reporte_completo.xlsx", 
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                             use_container_width=True)
            
            # Resumen por Empresa
            st.subheader("🏢 Resumen por Gestor")
            st.dataframe(resumen_empresa, use_container_width=True, hide_index=True)
            
            # Resumen por Residuo
            st.subheader("♻️ Resumen por Residuo")
            st.dataframe(resumen_residuo, use_container_width=True, hide_index=True)
