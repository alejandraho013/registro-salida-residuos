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

st.set_page_config(page_title="TINTATEX - Gestión de Residuos", layout="wide")
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


tab1, tab2 = st.tabs(["📝 Registro", "📊 Reportes"])

# ─────────────────────────────────────────────────────────────
# TAB 1: REGISTRO
# ─────────────────────────────────────────────────────────────
with tab1:
    if st.session_state.envio_exitoso:
        st.success("¡Registro guardado correctamente! ✅")
        if st.button("Iniciar Nuevo Registro"):
            st.session_state.envio_exitoso = False
            st.session_state.lista_temporal = []
            cargar_datos.clear()
            st.rerun()
    else:
        # --- 1. DATOS DEL VEHÍCULO ---
        with st.expander("🚛 1. Datos del Vehículo", expanded=True):
            c1, c2 = st.columns(2)
            fecha = c1.date_input("Fecha", datetime.now())
            empresa_sel = c1.selectbox("Empresa Gestora", options=list(GESTORES_DATA.keys()))

            empresa_final = empresa_sel
            if empresa_sel == "Otro":
                empresa_final = c1.text_input("Nombre del Gestor Manual").upper().strip()

            conductor = c2.text_input("Conductor")
            placa = c2.text_input("Placa (ABC123)").upper().strip()
            placa_valida = bool(re.match(r"^[A-Z]{3}[0-9]{3}$", placa)) if placa else False
            if placa and not placa_valida:
                c2.warning("Formato incorrecto. Ejemplo válido: ABC123")

        st.markdown("---")

        # --- 2. PESAJES ---
        st.subheader("2. Detalle de Pesajes")
        col_res, col_pes = st.columns([2, 1])
        opciones_res = GESTORES_DATA.get(empresa_sel, []).copy()
        opciones_res.append("Otro")
        tipo_sel = col_res.selectbox("Residuo", options=opciones_res)
        residuo_final = col_res.text_input("Especifique el residuo") if tipo_sel == "Otro" else tipo_sel
        peso = col_pes.number_input("Peso (kg)", min_value=0.0, max_value=50000.0, step=0.1)

        if st.button("➕ AGREGAR PESAJE", use_container_width=True):
            if not placa_valida:
                st.error("Corrija el formato de la placa antes de agregar pesajes.")
            elif peso <= 0:
                st.error("El peso debe ser mayor a 0.")
            elif not residuo_final:
                st.error("Especifique el tipo de residuo.")
            else:
                st.session_state.lista_temporal.append({"tipo_residuo": residuo_final, "peso_kg": peso})
                st.toast(f"✅ Agregado: {residuo_final} — {peso} kg")

        if st.session_state.lista_temporal:
            df_temp = pd.DataFrame(st.session_state.lista_temporal)
            col_m1, col_m2 = st.columns(2)
            col_m1.metric("Total Acumulado", f"{df_temp['peso_kg'].sum():,.1f} kg")
            col_m2.metric("Pesajes Ingresados", len(st.session_state.lista_temporal))
            with st.expander("🔍 Ver detalle de pesajes", expanded=False):
                st.dataframe(df_temp, use_container_width=True, hide_index=True)
                if st.button("🗑️ Limpiar lista", type="secondary"):
                    st.session_state.lista_temporal = []
                    st.rerun()

        st.markdown("---")

        # --- 3. EVIDENCIAS Y ENVÍO ---
        st.subheader("3. Finalizar y Clasificar")
        f1, f2 = st.columns(2)
        foto_memo = f1.file_uploader("Foto Memo (Opcional)", type=["jpg", "png", "jpeg"])
        foto_camion = f2.file_uploader("Foto Camión (Opcional)", type=["jpg", "png", "jpeg"])
        novedades = st.text_area("Observaciones")

        if st.button("📤 ENVIAR REGISTRO", type="primary", use_container_width=True):
            if not st.session_state.lista_temporal:
                st.error("Agregue al menos un pesaje antes de enviar.")
            elif not placa_valida:
                st.error("Corrija el formato de la placa (Ejemplo: ABC123).")
            elif empresa_sel == "Otro" and not empresa_final:
                st.error("Especifique el nombre de la empresa gestora.")
            else:
                with st.spinner("Guardando registro..."):
                    try:
                        g = Github(TOKEN)
                        repo = g.get_repo(REPO_NAME)
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        mes_actual = fecha.strftime("%B_%Y")

                        u_memo, u_camion = "Sin foto", "Sin foto"
                        if foto_memo:
                            p_memo = f"fotos/MEMO_{ts}.jpg"
                            repo.create_file(p_memo, f"Memo {ts}", foto_memo.getvalue())
                            u_memo = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_memo}"
                        if foto_camion:
                            p_camion = f"fotos/CAMION_{ts}.jpg"
                            repo.create_file(p_camion, f"Camion {ts}", foto_camion.getvalue())
                            u_camion = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{p_camion}"

                        xlsx_file = repo.get_contents("database.xlsx")
                        diccionario_hojas = pd.read_excel(
                            io.BytesIO(xlsx_file.decoded_content), sheet_name=None, engine="openpyxl"
                        )

                        nov_final = novedades.strip() if novedades.strip() else "Sin observaciones"
                        nuevas_filas = [
                            {
                                "fecha": str(fecha),
                                "mes": mes_actual,
                                "empresa": empresa_final,
                                "conductor": conductor,
                                "placa": placa,
                                "tipo_residuo": x["tipo_residuo"],
                                "peso_kg": x["peso_kg"],
                                "novedades": nov_final,
                                "url_memo": u_memo,
                                "url_camion": u_camion,
                            }
                            for x in st.session_state.lista_temporal
                        ]
                        df_nuevos = pd.DataFrame(nuevas_filas)

                        if "MASTER" in diccionario_hojas:
                            diccionario_hojas["MASTER"] = pd.concat(
                                [diccionario_hojas["MASTER"], df_nuevos], ignore_index=True
                            )
                        else:
                            diccionario_hojas["MASTER"] = df_nuevos

                        nombre_hoja = re.sub(r'[\\/*?:\[\]]', "", empresa_final)[:30].strip().upper()
                        if not nombre_hoja:
                            nombre_hoja = "OTROS"

                        if nombre_hoja in diccionario_hojas:
                            diccionario_hojas[nombre_hoja] = pd.concat(
                                [diccionario_hojas[nombre_hoja], df_nuevos], ignore_index=True
                            )
                        else:
                            diccionario_hojas[nombre_hoja] = df_nuevos

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            for hoja, df_hoja in diccionario_hojas.items():
                                df_hoja.to_excel(writer, sheet_name=hoja, index=False)

                        repo.update_file(
                            "database.xlsx",
                            f"Registro {empresa_final} {placa} {ts}",
                            output.getvalue(),
                            xlsx_file.sha,
                        )

                        st.session_state.envio_exitoso = True
                        st.rerun()

                    except Exception as e:
                        st.error("Error al guardar el registro. Intente de nuevo o contacte soporte.")
                        with st.expander("Detalle del error (soporte técnico)"):
                            st.exception(e)

# ─────────────────────────────────────────────────────────────
# TAB 2: REPORTES
# ─────────────────────────────────────────────────────────────
with tab2:
    st.header("📊 Reportes y Estadísticas")

    col_ref, _ = st.columns([1, 4])
    if col_ref.button("🔄 Actualizar datos"):
        cargar_datos.clear()
        st.rerun()

    df_master = cargar_datos()

    if df_master.empty:
        st.info("No hay datos disponibles o no se pudo conectar con la base de datos.")
    else:
        # ── FILTROS ──
        with st.expander("🔍 Filtros", expanded=True):
            fc1, fc2, fc3 = st.columns(3)
            empresas = ["Todas"] + sorted(df_master["empresa"].dropna().unique().tolist())
            residuos = ["Todos"] + sorted(df_master["tipo_residuo"].dropna().unique().tolist())
            meses = ["Todos"] + sorted(df_master["mes"].dropna().unique().tolist())

            empresa_filtro = fc1.selectbox("Empresa", empresas)
            residuo_filtro = fc2.selectbox("Tipo de Residuo", residuos)
            mes_filtro = fc3.selectbox("Mes", meses)

        df_f = df_master.copy()
        if empresa_filtro != "Todas":
            df_f = df_f[df_f["empresa"] == empresa_filtro]
        if residuo_filtro != "Todos":
            df_f = df_f[df_f["tipo_residuo"] == residuo_filtro]
        if mes_filtro != "Todos":
            df_f = df_f[df_f["mes"] == mes_filtro]

        # ── KPIs ──
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Registros", f"{len(df_f):,}")
        k2.metric("Peso Total (kg)", f"{df_f['peso_kg'].sum():,.1f}")
        k3.metric("Empresas Activas", df_f["empresa"].nunique())
        promedio = df_f["peso_kg"].mean() if len(df_f) > 0 else 0
        k4.metric("Promedio por Registro", f"{promedio:,.1f} kg")

        st.markdown("---")

        # ── GRÁFICAS ──
        col_g1, col_g2 = st.columns(2)

        with col_g1:
            st.subheader("Peso por Empresa (kg)")
            if not df_f.empty:
                peso_empresa = (
                    df_f.groupby("empresa")["peso_kg"]
                    .sum()
                    .sort_values(ascending=False)
                    .rename("kg")
                )
                st.bar_chart(peso_empresa)
            else:
                st.info("Sin datos para mostrar.")

        with col_g2:
            st.subheader("Peso por Tipo de Residuo (kg)")
            if not df_f.empty:
                peso_residuo = (
                    df_f.groupby("tipo_residuo")["peso_kg"]
                    .sum()
                    .sort_values(ascending=False)
                    .rename("kg")
                )
                st.bar_chart(peso_residuo)
            else:
                st.info("Sin datos para mostrar.")

        st.subheader("Tendencia Mensual (kg totales)")
        if not df_f.empty:
            tendencia = df_f.groupby("mes")["peso_kg"].sum().rename("kg")
            st.line_chart(tendencia)
        else:
            st.info("Sin datos para mostrar.")

        st.markdown("---")

        # ── TABLA DETALLADA ──
        st.subheader("Datos Detallados")
        cols_display = ["fecha", "empresa", "conductor", "placa", "tipo_residuo", "peso_kg", "mes", "novedades"]
        cols_display = [c for c in cols_display if c in df_f.columns]
        df_tabla = df_f[cols_display].sort_values("fecha", ascending=False)
        st.dataframe(df_tabla, use_container_width=True, hide_index=True)

        # ── RESUMEN POR EMPRESA ──
        st.subheader("Resumen por Empresa")
        resumen = (
            df_f.groupby("empresa")
            .agg(registros=("peso_kg", "count"), peso_total_kg=("peso_kg", "sum"), peso_promedio_kg=("peso_kg", "mean"))
            .round(1)
            .sort_values("peso_total_kg", ascending=False)
            .reset_index()
        )
        st.dataframe(resumen, use_container_width=True, hide_index=True)

        # ── EXPORTAR ──
        st.markdown("---")
        csv_bytes = df_tabla.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️ Descargar datos filtrados (CSV)",
            data=csv_bytes,
            file_name=f"residuos_TINTATEX_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

