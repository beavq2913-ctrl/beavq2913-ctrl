import streamlit as st
import pandas as pd

from liquidacion_core import (
    MONTO_FIJO_ALBERTON_DEFAULT,
    HOJAS_VALIDAS,
    procesar,
    construir_excel_general,
    construir_zip_patologo,
)

# ─── STREAMLIT UI ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="Liquidación Patólogos", page_icon="🔬", layout="wide")

st.markdown("""
<style>
.block-container { padding-top: 2rem; }
.stButton>button {
    background-color: #1F3864; color: white;
    border-radius: 8px; padding: 0.5rem 2rem;
    font-size: 1rem; font-weight: bold; width: 100%;
}
.stButton>button:hover { background-color: #2e4f8a; }
.stDownloadButton>button {
    background-color: #1a7a4a; color: white;
    border-radius: 8px; padding: 0.6rem 2rem;
    font-size: 1rem; font-weight: bold; width: 100%;
}
.stDownloadButton>button:hover { background-color: #145e38; }
.metric-card {
    background: white; border-radius: 12px;
    padding: 1.2rem 1rem; text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-top: 4px solid #1F3864;
}
.metric-value { font-size: 1.5rem; font-weight: 800; color: #1F3864; }
.metric-label { font-size: 0.8rem; color: #888; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.05em; }
.alerta { background:#fff3cd; border-left:4px solid #ffc107; padding:0.8rem 1rem; border-radius:6px; margin:0.5rem 0; }
.descarga-box { background:#f0f7f0; border:2px solid #1a7a4a; border-radius:10px; padding:1.2rem; margin:0.5rem 0; }
</style>
""", unsafe_allow_html=True)

col_logo, col_titulo = st.columns([1, 8])
with col_logo:
    st.markdown("<div style='font-size:3rem;text-align:center'>🔬</div>", unsafe_allow_html=True)
with col_titulo:
    st.markdown("## Liquidación de Patólogos")
    st.markdown("<span style='color:#666;font-size:0.95rem'>Subí el archivo Excel consolidado y descargá la liquidación calculada automáticamente</span>", unsafe_allow_html=True)

st.divider()

uploaded = st.file_uploader(
    "📂 Seleccioná el archivo Excel",
    type=["xlsx", "xls"],
    help=f"El archivo debe tener una hoja llamada: {', '.join(HOJAS_VALIDAS)}"
)

if uploaded:
    try:
        xf         = pd.ExcelFile(uploaded)
        hojas      = xf.sheet_names
        hoja_datos = next((h for h in HOJAS_VALIDAS if h in hojas), None)

        if hoja_datos is None:
            st.error(f"❌ No se encontró una hoja válida. Hojas disponibles: {', '.join(hojas)}")
            st.stop()

        with st.spinner("Leyendo archivo..."):
            df = pd.read_excel(uploaded, sheet_name=hoja_datos, header=0, dtype=str)
            df.columns = [str(c).replace("\xa0", " ").strip() for c in df.columns]
            df["Subtotal"] = pd.to_numeric(df["Subtotal"], errors="coerce").fillna(0)
            filas_antes = len(df)
            df = df[~df["Cod. OS"].fillna("").str.strip().isin(["0", "0.0", ""])]
            df = df.reset_index(drop=True)
            filas_eliminadas = filas_antes - len(df)

        nombre_base = uploaded.name.replace(".xlsx", "").replace(".xls", "")
        msg = f"✅ Archivo cargado — hoja **{hoja_datos}** — {len(df):,} estudios"
        if filas_eliminadas > 0:
            msg += f" ({filas_eliminadas} filas con Cod. OS=0 eliminadas)"
        st.success(msg)

    except Exception as e:
        st.error(f"❌ Error al leer el archivo: {e}")
        st.stop()

    st.divider()
    st.subheader("⚙️ Parámetros del mes")
    col_alb, _ = st.columns([1, 2])
    with col_alb:
        monto_alberton = st.number_input(
            "Monto Valeria Alberton ($)",
            min_value=0.0,
            value=MONTO_FIJO_ALBERTON_DEFAULT,
            step=100.0,
            format="%.2f",
            help=f"Monto fijo por ingreso para este mes. Valor base: ${MONTO_FIJO_ALBERTON_DEFAULT:,.2f}"
        )
    st.divider()

    if st.button("⚡ Calcular liquidación"):
        prog = st.progress(0, text="Procesando estudios...")
        with st.spinner(""):
            df_result = procesar(df, monto_alberton=monto_alberton)
            prog.progress(50, text="Generando Excel general...")
            excel_buf, n_inc = construir_excel_general(df_result)
            prog.progress(80, text="Generando archivos por patólogo...")
            zip_buf, n_patologos = construir_zip_patologo(df_result, nombre_base)
            prog.progress(100, text="¡Listo!")

        st.divider()
        st.subheader("📊 Resumen")

        total_liq = df_result["Total Liquidado"].sum()
        total_sub = df_result["Subtotal"].sum()
        total_1ra = df_result["Importe Primera Firma"].sum()
        total_2da = df_result["Importe Segunda Firma"].sum()

        c1, c2, c3, c4, c5 = st.columns(5)
        for col, label, value in [
            (c1, "Estudios",          f"{len(df_result):,}"),
            (c2, "Subtotal facturado", f"${total_sub:,.0f}"),
            (c3, "Total 1ra firma",   f"${total_1ra:,.0f}"),
            (c4, "Total 2da firma",   f"${total_2da:,.0f}"),
            (c5, "TOTAL LIQUIDADO",   f"${total_liq:,.0f}"),
        ]:
            col.markdown(f"""
                <div class='metric-card'>
                    <div class='metric-value'>{value}</div>
                    <div class='metric-label'>{label}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("")

        if n_inc > 0:
            subtotal_cero = df_result[df_result["Observaciones"].str.contains("Subtotal = 0", na=False)]
            cod_esp_2da   = df_result[df_result["Observaciones"].str.contains("Código especial", na=False)]
            sin_pat       = df_result[df_result["Observaciones"].str.contains("Sin patólogo", na=False)]
            st.markdown(f"<div class='alerta'>⚠️ Se detectaron <b>{n_inc} inconsistencias</b> — revisalas en la hoja <b>Inconsistencias</b> del Excel general.</div>", unsafe_allow_html=True)
            cols_inc = st.columns(3)
            if len(subtotal_cero): cols_inc[0].metric("Subtotal = 0", len(subtotal_cero))
            if len(cod_esp_2da):   cols_inc[1].metric("Códigos especiales con 2da firma", len(cod_esp_2da))
            if len(sin_pat):       cols_inc[2].metric("Sin patólogo", len(sin_pat))

        st.divider()
        st.subheader("👩‍⚕️ Totales por patólogo")

        g1_st = df_result[df_result["Patólogo Primera Firma"] != ""].groupby("Patólogo Primera Firma").agg(
            pf=("Importe Primera Firma", "sum"), pres=("Presencias", "sum")
        ).reset_index().rename(columns={"Patólogo Primera Firma": "Patólogo"})
        g2_st = df_result[df_result["Patólogo Segunda Firma"] != ""].groupby("Patólogo Segunda Firma").agg(
            sf=("Importe Segunda Firma", "sum")
        ).reset_index().rename(columns={"Patólogo Segunda Firma": "Patólogo"})
        resumen = pd.merge(g1_st, g2_st, on="Patólogo", how="outer").fillna(0)
        resumen["Total ($)"] = resumen["pf"] + resumen["pres"] + resumen["sf"]
        resumen = resumen.sort_values("Total ($)", ascending=False).reset_index(drop=True)
        resumen = resumen.rename(columns={"pf": "1ra Firma ($)", "pres": "Presencias ($)", "sf": "2da Firma ($)"})
        resumen = resumen[["Patólogo", "1ra Firma ($)", "Presencias ($)", "2da Firma ($)", "Total ($)"]]
        fila_total = pd.DataFrame([{c: resumen[c].sum() if c != "Patólogo" else "TOTAL GENERAL" for c in resumen.columns}])
        resumen_display = pd.concat([resumen, fila_total], ignore_index=True)
        for col in resumen.columns[1:]:
            resumen_display[col] = resumen_display[col].apply(lambda x: f"${float(x):,.2f}" if x != "" else "")
        st.dataframe(resumen_display, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("⬇️ Descargas")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.markdown("<div class='descarga-box'>", unsafe_allow_html=True)
            st.markdown("**📋 Excel general completo**")
            st.markdown("Incluye todas las hojas: Detalle, Resumen, Inconsistencias y Totales.")
            st.download_button(
                label="⬇️ Descargar Excel general",
                data=excel_buf,
                file_name=f"liquidacion_{nombre_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.markdown("</div>", unsafe_allow_html=True)

        with col_d2:
            st.markdown("<div class='descarga-box'>", unsafe_allow_html=True)
            st.markdown(f"**🗂 ZIP con {n_patologos} archivos individuales**")
            st.markdown("Un Excel por patólogo con sus estudios de 1ra y 2da firma + resumen.")
            st.download_button(
                label="⬇️ Descargar ZIP por patólogo",
                data=zip_buf,
                file_name=f"liquidacion_por_patologo_{nombre_base}.zip",
                mime="application/zip",
                use_container_width=True,
            )
            st.markdown("</div>", unsafe_allow_html=True)

else:
    st.info("👆 Subí el archivo Excel para comenzar.")
    with st.expander("ℹ️ ¿Cómo funciona?"):
        st.markdown("""
        1. **Subís** el archivo Excel consolidado (hoja: LIQUIDACION, BASE o CONSOLIDADO)
        2. **Configurás** el monto fijo de Valeria Alberton si cambió
        3. **Hacés clic** en *Calcular liquidación*
        4. **Descargás** dos archivos:
           - **Excel general**: Detalle, Resumen, Inconsistencias y Totales
           - **ZIP individual**: un Excel por patólogo con sus estudios y total a cobrar
        """)

st.divider()
st.caption("Liquidación de Patólogos · Biogenar · Desarrollado con Streamlit")
