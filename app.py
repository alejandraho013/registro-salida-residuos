with tab2:
    col_ref, _ = st.columns([1, 5])
    if col_ref.button("🔄 Actualizar", key="refresh"):
        cargar_datos_github.clear()
        st.rerun()

    with st.spinner("Cargando datos…"):
        df = cargar_datos_github()

    if df.empty:
        st.info("📭 Sin datos registrados aún.")
    else:
        # ── FILTROS ──────────────────────────────────────────
        with st.expander("🔍 Filtros", expanded=True):
            tipo_rango = st.radio(
                "Período",
                ["Rango de fechas", "Día exacto", "Mes", "Todo"],
                horizontal=True,
            )

            f1, f2, f3, f4 = st.columns([2, 2, 2, 2])

            # Filtro de fechas
            fecha_min = df["fecha"].min().date()
            fecha_max = df["fecha"].max().date()

            if tipo_rango == "Rango de fechas":
                fecha_ini = f1.date_input("Desde", value=fecha_min, min_value=fecha_min, max_value=fecha_max)
                fecha_fin = f2.date_input("Hasta", value=fecha_max, min_value=fecha_min, max_value=fecha_max)
                mask_fecha = (df["fecha"].dt.date >= fecha_ini) & (df["fecha"].dt.date <= fecha_fin)
            elif tipo_rango == "Día exacto":
                dia = f1.date_input("Día", value=fecha_max, min_value=fecha_min, max_value=fecha_max)
                mask_fecha = df["fecha"].dt.date == dia
            elif tipo_rango == "Mes":
                meses = sorted(df["mes"].dropna().unique().tolist())
                mes_sel = f1.selectbox("Mes", meses)
                mask_fecha = df["mes"] == mes_sel
            else:
                mask_fecha = pd.Series([True] * len(df), index=df.index)

            # Filtro empresa
            empresas = ["Todas"] + sorted(df["empresa"].dropna().unique().tolist())
            emp_sel = f3.selectbox("Empresa", empresas)
            mask_emp = df["empresa"] == emp_sel if emp_sel != "Todas" else pd.Series([True] * len(df), index=df.index)

            # Filtro residuo
            residuos = ["Todos"] + sorted(df["tipo_residuo"].dropna().unique().tolist())
            res_sel = f4.selectbox("Tipo de residuo", residuos)
            mask_res = df["tipo_residuo"] == res_sel if res_sel != "Todos" else pd.Series([True] * len(df), index=df.index)

        df_f = df[mask_fecha & mask_emp & mask_res].copy()

        if df_f.empty:
            st.warning("⚠️ No hay registros con los filtros seleccionados.")
        else:
            # ── KPIs ────────────────────────────────────────
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Peso total", f"{df_f['peso_kg'].sum():,.1f} kg")
            k2.metric("Registros",  len(df_f))
            k3.metric("Empresas",   df_f["empresa"].nunique())
            k4.metric("Promedio / registro", f"{df_f['peso_kg'].mean():,.1f} kg")

            st.divider()

            # ── GRÁFICAS ────────────────────────────────────
            g1, g2 = st.columns(2)

            with g1:
                emp_agg = df_f.groupby("empresa")["peso_kg"].sum().reset_index()
                fig_bar = px.bar(
                    emp_agg, x="empresa", y="peso_kg",
                    title="Peso total por empresa",
                    color="empresa",
                    color_discrete_map=COLORES_EMPRESA,
                    labels={"peso_kg": "Peso (kg)", "empresa": ""},
                    text_auto=".0f",
                )
                fig_bar.update_traces(textposition="outside")
                fig_bar.update_layout(showlegend=False, margin=dict(t=40, b=0))
                st.plotly_chart(fig_bar, use_container_width=True)

            with g2:
                fig_pie = px.pie(
                    df_f, values="peso_kg", names="tipo_residuo",
                    title="Distribución por tipo de residuo",
                )
                fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                fig_pie.update_layout(margin=dict(t=40, b=0), showlegend=False)
                st.plotly_chart(fig_pie, use_container_width=True)

            # Línea temporal (solo si hay más de 1 día distinto)
            if df_f["fecha"].dt.date.nunique() > 1:
                df_tiempo = (
                    df_f.groupby(df_f["fecha"].dt.date)["peso_kg"]
                    .sum()
                    .reset_index()
                    .rename(columns={"fecha": "Fecha", "peso_kg": "Peso (kg)"})
                )
                fig_line = px.line(
                    df_tiempo, x="Fecha", y="Peso (kg)",
                    title="Evolución diaria de peso",
                    markers=True,
                )
                fig_line.update_layout(margin=dict(t=40, b=0))
                st.plotly_chart(fig_line, use_container_width=True)

            # ── TABLA DETALLE ────────────────────────────────
            st.subheader("Detalle de registros")
            st.dataframe(
                df_f.sort_values("fecha", ascending=False).reset_index(drop=True),
                use_container_width=True,
                hide_index=True,
            )

            # ── DESCARGA EXCEL MULTI-HOJA ────────────────────
            def construir_excel(df_filtrado: pd.DataFrame) -> bytes:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:

                    # Hoja 1: todos los registros filtrados
                    df_filtrado.to_excel(writer, sheet_name="Registros", index=False)

                    # Hoja 2: resumen por empresa
                    resumen_emp = (
                        df_filtrado.groupby("empresa")
                        .agg(
                            registros=("peso_kg", "count"),
                            peso_total_kg=("peso_kg", "sum"),
                            peso_promedio_kg=("peso_kg", "mean"),
                        )
                        .round(2)
                        .reset_index()
                        .sort_values("peso_total_kg", ascending=False)
                    )
                    resumen_emp.to_excel(writer, sheet_name="Por empresa", index=False)

                    # Hoja 3: resumen por tipo de residuo
                    resumen_res = (
                        df_filtrado.groupby("tipo_residuo")
                        .agg(
                            registros=("peso_kg", "count"),
                            peso_total_kg=("peso_kg", "sum"),
                            peso_promedio_kg=("peso_kg", "mean"),
                        )
                        .round(2)
                        .reset_index()
                        .sort_values("peso_total_kg", ascending=False)
                    )
                    resumen_res.to_excel(writer, sheet_name="Por residuo", index=False)

                    # Hoja 4: resumen por fecha
                    resumen_fecha = (
                        df_filtrado.groupby(df_filtrado["fecha"].dt.date)
                        .agg(
                            registros=("peso_kg", "count"),
                            peso_total_kg=("peso_kg", "sum"),
                        )
                        .round(2)
                        .reset_index()
                        .rename(columns={"fecha": "Fecha"})
                        .sort_values("Fecha", ascending=False)
                    )
                    resumen_fecha.to_excel(writer, sheet_name="Por fecha", index=False)

                    # Hoja 5: resumen por empresa × residuo (tabla cruzada)
                    pivot = (
                        df_filtrado.pivot_table(
                            index="empresa",
                            columns="tipo_residuo",
                            values="peso_kg",
                            aggfunc="sum",
                            fill_value=0,
                        )
                        .round(2)
                        .reset_index()
                    )
                    pivot.to_excel(writer, sheet_name="Cruce empresa-residuo", index=False)

                return output.getvalue()

            nombre_archivo = (
                f"Reporte_{emp_sel}_{res_sel}_{tipo_rango.replace(' ', '_')}.xlsx"
                .replace("Todas", "todas_empresas")
                .replace("Todos", "todos_residuos")
            )

            st.download_button(
                label="⬇️ Descargar Excel (5 hojas organizadas)",
                data=construir_excel(df_f),
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
