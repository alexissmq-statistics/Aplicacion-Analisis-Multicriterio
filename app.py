"""
App Streamlit para el Sistema de Soporte a la Decisión Multicriterio (AHP).
Toda la lógica matemática vive en ahp_logic.py; este archivo solo arma la interfaz.

Soporta dos modos (AHP objetivo / AHP2 subjetivo) y dos orígenes de datos
(formularios manuales o carga de un archivo CSV con la estructura esperada).
"""

import io

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

import ahp_logic as ahp

st.set_page_config(page_title="Sistema AHP", layout="wide")
st.title("Sistema de Soporte a la Decisión Multicriterio (AHP)")

ESCALA_SAATY = {
    "9 : Extrema (A)": 9, "8 : Intermedio (A)": 8, "7 : Muy fuerte (A)": 7,
    "6 : Intermedio (A)": 6, "5 : Fuerte (A)": 5, "4 : Intermedio (A)": 4,
    "3 : Moderada (A)": 3, "2 : Intermedio (A)": 2, "1 : Igual": 1,
    "2 : Intermedio 1/2 (B)": 1/2, "3 : Moderada 1/3 (B)": 1/3, "4 : Intermedio 1/4 (B)": 1/4,
    "5 : Fuerte 1/5 (B)": 1/5, "6 : Intermedio 1/6 (B)": 1/6, "7 : Muy fuerte 1/7 (B)": 1/7,
    "8 : Intermedio 1/8 (B)": 1/8, "9 : Extrema 1/9 (B)": 1/9,
}


def _formatear_valor_bloqueado(v):
    """Etiqueta corta y legible para las celdas bloqueadas (diagonal y triángulo inferior)."""
    if abs(v - 1.0) < 1e-9:
        return "1"
    if v < 1:
        entero = round(1 / v)
        return f"1/{entero}"
    return f"{v:.2f}"

PLANTILLA_AHP = """n_criterios,4
n_alternativas,3
Criterios,Precio,Calidad,Tiempo,Certificación
Tipos,C,B,C,B
Matriz_Criterios,1,3,5,4
Matriz_Criterios,1/3,1,3,2
Matriz_Criterios,1/5,1/3,1,0.5
Matriz_Criterios,1/4,0.5,2,1
P1,85,96,5,9
P2,78,92,8,7
P3,90,98,3,10
"""

PLANTILLA_AHP2 = """n_criterios,4
n_alternativas,3
Criterios,Precio,Calidad,Tiempo,Certificación
Alternativas,P1,P2,P3

Matriz_Criterios
1,3,5,4
0.333333333,1,3,2
0.2,0.333333333,1,0.5
0.25,0.5,2,1

Matriz_Alternativas_Precio
1,0.917647059,1.058823529
1.08974359,1,1.153846154
0.944444444,0.866666667,1

Matriz_Alternativas_Calidad
1,1.043478261,0.979591837
0.958333333,1,0.93877551
1.020833333,1.065217391,1

Matriz_Alternativas_Tiempo
1,1.6,0.6
0.625,1,0.375
1.666666667,2.666666667,1

Matriz_Alternativas_Certificación
1,1.285714286,0.9
0.777777778,1,0.7
1.111111111,1.428571429,1
"""


def widget_matriz_pareada(nombres, key_prefix):
    """
    Dibuja la matriz de comparación pareada completa (como en el modelo original en Excel):
    diagonal = 1 (bloqueada), triángulo superior editable (selectbox con escala de Saaty),
    triángulo inferior = 1/valor, bloqueado y calculado automáticamente.

    Retorna la matriz (n x n) completa como np.ndarray.
    """
    n = len(nombres)
    matriz = np.ones((n, n))

    # Encabezado de columnas
    cols_header = st.columns([1.6] + [1] * n)
    cols_header[0].markdown("**1 / k**")
    for j in range(n):
        cols_header[j + 1].markdown(f"**{nombres[j]}**")

    for i in range(n):
        fila = st.columns([1.6] + [1] * n)
        fila[0].markdown(f"**{nombres[i]}**")
        for j in range(n):
            with fila[j + 1]:
                if i == j:
                    st.text_input(
                        " ", value="1", key=f"{key_prefix}_diag_{i}", disabled=True,
                        label_visibility="collapsed",
                    )
                elif j > i:
                    etiquetas = list(ESCALA_SAATY.keys())
                    elegido = st.selectbox(
                        " ", etiquetas, index=8,  # índice de "1 : Igual"
                        key=f"{key_prefix}_{i}_{j}", label_visibility="collapsed",
                    )
                    matriz[i, j] = ESCALA_SAATY[elegido]
                else:
                    # Triángulo inferior: recíproco calculado, solo lectura.
                    valor_reciproco = 1.0 / matriz[j, i]
                    matriz[i, j] = valor_reciproco
                    st.text_input(
                        " ", value=_formatear_valor_bloqueado(valor_reciproco),
                        key=f"{key_prefix}_lock_{i}_{j}", disabled=True,
                        label_visibility="collapsed",
                    )
    return matriz


def mostrar_consistencia(matriz, vector_pesos, titulo):
    lambda_max, ci, cr, estado, _, _ = ahp.calcular_consistencia(matriz, vector_pesos)
    c1, c2, c3 = st.columns(3)
    c1.metric("λ máx", f"{lambda_max:.4f}")
    c2.metric("CI", f"{ci:.4f}")
    c3.metric("CR", f"{cr:.4f}")
    if cr <= 0.10:
        st.success(f"✅ {titulo}: {estado}")
    else:
        st.warning(f"⚠️ {titulo}: {estado}")


def leer_csv_subido(archivo_subido):
    """Decodifica el archivo subido por st.file_uploader a texto plano (utf-8)."""
    return archivo_subido.getvalue().decode("utf-8")


# =======================================================================================
# BARRA LATERAL: CONFIGURACIÓN GENERAL
# =======================================================================================
if st.sidebar.button("🔄 Nuevo análisis (borrar todo)", width="stretch"):
    st.session_state.clear()
    st.rerun()

st.sidebar.caption(
    "Usa este botón si ya descargaste tus resultados y quieres empezar un análisis "
    "nuevo desde cero (borra todos los datos ingresados en el formulario)."
)
st.sidebar.divider()

st.sidebar.header("Configuración del problema")
modo = st.sidebar.radio(
    "Modo de análisis",
    ["Modo AHP (datos objetivos costo/beneficio)", "Modo AHP2 (comparación subjetiva)"],
)
es_modo_objetivo = modo.startswith("Modo AHP (")

origen_datos = st.sidebar.radio(
    "Origen de los datos",
    ["Formulario manual", "Cargar archivo CSV"],
)

st.divider()

# =======================================================================================
# CARGA / CAPTURA DE DATOS SEGÚN EL ORIGEN ELEGIDO
# =======================================================================================
criterios = None
alternativas = None
matriz_crit = None
tipos_criterios = None
df_rendimiento = None
matrices_alt_csv = None

if origen_datos == "Cargar archivo CSV":
    st.header("📄 Carga de datos por archivo CSV")

    plantilla = PLANTILLA_AHP if es_modo_objetivo else PLANTILLA_AHP2
    nombre_archivo = "plantilla_ahp.csv" if es_modo_objetivo else "plantilla_ahp2.csv"

    with st.expander("Ver / descargar formato esperado para este modo"):
        st.code(plantilla, language="text")
        st.download_button(
            "⬇️ Descargar plantilla de ejemplo",
            data=plantilla,
            file_name=nombre_archivo,
            mime="text/csv",
        )

    archivo = st.file_uploader(
        f"Sube tu CSV en formato {'AHP objetivo' if es_modo_objetivo else 'AHP2 subjetivo'}",
        type=["csv"],
    )

    if archivo is None:
        st.info("Sube un archivo CSV con el formato indicado arriba para continuar.")
        st.stop()

    try:
        contenido = leer_csv_subido(archivo)
        if es_modo_objetivo:
            criterios, tipos_criterios, matriz_crit, alternativas, df_rendimiento = ahp.cargar_datos_ahp_csv(contenido)
        else:
            criterios, alternativas, matriz_crit, matrices_alt_csv = ahp.cargar_datos_ahp2_csv(contenido)
    except Exception as e:
        st.error(f"❌ No se pudo leer el archivo: {e}")
        st.stop()

    st.success(f"Archivo cargado: {len(criterios)} criterios, {len(alternativas)} alternativas.")
    n_crit, n_alt = len(criterios), len(alternativas)

else:
    n_crit = st.sidebar.number_input("Número de criterios", min_value=2, max_value=8, value=3)
    criterios = [st.sidebar.text_input(f"Nombre criterio {i+1}", value=f"Criterio {i+1}") for i in range(n_crit)]

    n_alt = st.sidebar.number_input("Número de alternativas", min_value=2, max_value=8, value=3)
    alternativas = [st.sidebar.text_input(f"Nombre alternativa {i+1}", value=f"Alternativa {i+1}") for i in range(n_alt)]

    if es_modo_objetivo:
        tipos_criterios = [
            st.sidebar.selectbox(f"Tipo de '{criterios[i]}'", ["Beneficio", "Costo"], key=f"tipo_{i}")
            for i in range(n_crit)
        ]

# =======================================================================================
# PASO 1: MATRIZ DE COMPARACIÓN DE CRITERIOS
# =======================================================================================
st.header("1️⃣ Comparación pareada de criterios")

if matriz_crit is None:  # aún no viene de un CSV -> pedir por formulario
    st.caption(
        "Completa solo el triángulo **superior**. El triángulo inferior se calcula "
        "automáticamente como el recíproco (1/k) y queda bloqueado."
    )
    matriz_crit = widget_matriz_pareada(criterios, "crit")

valida, msg = ahp.validar_matriz_cuadrada_y_reciproca(matriz_crit)
if not valida:
    st.error(msg)
    st.stop()

resultado_crit = ahp.procesar_matriz_prioridades(matriz_crit, criterios, criterios)
pesos_criterios = resultado_crit["vector_pesos"]

with st.expander("Ver desarrollo paso a paso (criterios)"):
    st.write("**Matriz original**")
    st.dataframe(resultado_crit["matriz_original"].round(4))
    st.write("**Sumas por columna**")
    st.dataframe(resultado_crit["sumas_columnas"].round(4))
    st.write("**Matriz normalizada**")
    st.dataframe(resultado_crit["matriz_normalizada"].round(4))

st.write("**Vector de pesos de los criterios**")
st.dataframe(resultado_crit["vector_pesos_df"].round(4))
mostrar_consistencia(matriz_crit, pesos_criterios, "Consistencia de criterios")

st.divider()

# =======================================================================================
# PASO 2: MATRICES DE ALTERNATIVAS (según el modo)
# =======================================================================================
st.header("2️⃣ Evaluación de alternativas por criterio")

pesos_locales_dict = {}

if es_modo_objetivo:
    if df_rendimiento is None:  # capturar manualmente
        st.write(
            "Ingresa el valor **real** de desempeño de cada alternativa en cada criterio "
            "(igual que en la matriz de datos originales: filas = alternativas, columnas = criterios)."
        )
        etiquetas_columna = [
            f"{c} ({'Beneficio' if tipos_criterios[k] == 'Beneficio' else 'Costo'})"
            for k, c in enumerate(criterios)
        ]
        df_base = pd.DataFrame(1.0, index=alternativas, columns=etiquetas_columna)
        df_editado = st.data_editor(
            df_base,
            key="editor_rendimiento",
            width="stretch",
            column_config={
                col: st.column_config.NumberColumn(col, min_value=0.0, step=0.1, format="%.4f")
                for col in etiquetas_columna
            },
        )
        df_rendimiento = df_editado.copy()
        df_rendimiento.columns = criterios  # quitar la etiqueta (Beneficio)/(Costo) para el resto del pipeline
    else:
        st.write("Datos de desempeño cargados desde el CSV:")
        st.dataframe(df_rendimiento)

    matrices_alt = ahp.generar_matrices_alternativas_objetivas(df_rendimiento, tipos_criterios)

    for criterio in criterios:
        resultado = ahp.procesar_matriz_prioridades(matrices_alt[criterio], alternativas, alternativas)
        pesos_locales_dict[criterio] = resultado["vector_pesos"]
        with st.expander(f"Detalle: alternativas bajo '{criterio}'"):
            st.dataframe(resultado["matriz_original"].round(4))
            st.dataframe(resultado["vector_pesos_df"].round(4))
            mostrar_consistencia(matrices_alt[criterio], resultado["vector_pesos"], f"Consistencia '{criterio}'")

else:
    for criterio in criterios:
        st.subheader(f"Alternativas bajo el criterio: {criterio}")
        if matrices_alt_csv is None:  # capturar manualmente
            st.caption(
                "Completa solo el triángulo **superior**. El triángulo inferior se calcula "
                "automáticamente como el recíproco (1/k) y queda bloqueado."
            )
            matriz_alt = widget_matriz_pareada(alternativas, f"alt_{criterio}")
        else:
            matriz_alt = matrices_alt_csv[criterio]

        valida_alt, msg_alt = ahp.validar_matriz_cuadrada_y_reciproca(matriz_alt)
        if not valida_alt:
            st.error(f"Error en la matriz de '{criterio}': {msg_alt}")
            st.stop()

        resultado = ahp.procesar_matriz_prioridades(matriz_alt, alternativas, alternativas)
        pesos_locales_dict[criterio] = resultado["vector_pesos"]
        with st.expander(f"Detalle: alternativas bajo '{criterio}'"):
            st.dataframe(resultado["matriz_original"].round(4))
            st.dataframe(resultado["vector_pesos_df"].round(4))
            mostrar_consistencia(matriz_alt, resultado["vector_pesos"], f"Consistencia '{criterio}'")

df_prioridades_locales = pd.DataFrame(pesos_locales_dict, index=alternativas)

st.divider()

# =======================================================================================
# PASO 3: SÍNTESIS Y RANKING
# =======================================================================================
st.header("3️⃣ Síntesis global y ranking")
df_ranking, mejor_opcion = ahp.calcular_sintesis_y_ranking(pesos_criterios, df_prioridades_locales)

st.dataframe(df_ranking.round(4))
st.success(f"🏆 Mejor opción: **{mejor_opcion}** con puntaje {df_ranking.iloc[0]['Puntaje Global']:.4f}")

fig, ax = plt.subplots(1, 2, figsize=(12, 5))
df_c = pd.DataFrame({"Criterio": criterios, "Peso": pesos_criterios}).sort_values("Peso")
ax[0].barh(df_c["Criterio"], df_c["Peso"], color="#4A90E2")
ax[0].set_title("Pesos de criterios")

df_r = df_ranking.sort_values("Puntaje Global")
colores = ["#2ECC71" if x == mejor_opcion else "#BDC3C7" for x in df_r.index]
ax[1].barh(df_r.index, df_r["Puntaje Global"], color=colores)
ax[1].set_title("Ranking de alternativas")

plt.tight_layout()
st.pyplot(fig)

# Guardamos la figura como PNG en memoria para poder incrustarla en el Excel más abajo
buffer_fig = io.BytesIO()
fig.savefig(buffer_fig, format="png", dpi=150, bbox_inches="tight")
fig_bytes = buffer_fig.getvalue()

st.divider()

# =======================================================================================
# PASO 4: ANÁLISIS DE SENSIBILIDAD
# =======================================================================================
st.header("4️⃣ Análisis de sensibilidad")
st.write("Ajusta manualmente el peso de uno o más criterios; el resto se redistribuye proporcionalmente.")

criterios_a_modificar = st.multiselect("Criterios a modificar", criterios)
cambios_objetivo = {}
for c in criterios_a_modificar:
    idx = criterios.index(c)
    nuevo_peso = st.slider(f"Nuevo peso para '{c}'", 0.0, 1.0, float(pesos_criterios[idx]), 0.01)
    cambios_objetivo[idx] = nuevo_peso

resultado_sens = None
resultado_sens = None
if cambios_objetivo:
    try:
        resultado_sens = ahp.ejecutar_analisis_sensibilidad(pesos_criterios, df_prioridades_locales, cambios_objetivo)
        st.write("**Comparativa de pesos**")
        st.dataframe(resultado_sens["df_comparativo_criterios"].round(4))
        st.write("**Impacto en el ranking**")
        st.dataframe(resultado_sens["df_resultados"])
        if resultado_sens["ranking_estable"]:
            st.info("El ranking se mantiene ESTABLE ante este cambio.")
        else:
            st.warning("⚡ El ranking CAMBIÓ ante esta modificación de pesos.")
    except ValueError as e:
        st.error(str(e))

st.divider()

# =======================================================================================
# PASO 5: DESCARGA DE RESULTADOS
# =======================================================================================
st.header("⬇️ Descargar resultados")
st.write(
    "Descarga el ranking final, la mejor alternativa y (si lo hiciste) el análisis de "
    "sensibilidad, en el formato que prefieras."
)

col_csv, col_xlsx, col_txt = st.columns(3)

with col_csv:
    csv_bytes = df_ranking.round(4).to_csv().encode("utf-8-sig")
    st.download_button(
        "📄 Ranking (CSV)",
        data=csv_bytes,
        file_name="ranking_ahp.csv",
        mime="text/csv",
        width="stretch",
    )

with col_xlsx:
    reporte_excel = ahp.generar_reporte_excel(
        pesos_criterios=pesos_criterios,
        criterios=criterios,
        df_ranking=df_ranking,
        mejor_opcion=mejor_opcion,
        modo=modo,
        fig_bytes=fig_bytes,
        resultado_sens=resultado_sens,
    )
    st.download_button(
        "📊 Reporte completo (Excel)",
        data=reporte_excel,
        file_name="reporte_ahp.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

with col_txt:
    reporte_texto = ahp.generar_reporte_texto(
        pesos_criterios, criterios, df_ranking, mejor_opcion, resultado_sens
    )
    st.download_button(
        "📝 Reporte (TXT)",
        data=reporte_texto,
        file_name="reporte_ahp.txt",
        mime="text/plain",
        width="stretch",
    )

st.caption(
    "El reporte **Excel** incluye: resumen ejecutivo con la conclusión, el ranking final, "
    "los pesos de los criterios, el gráfico comparativo y — si se ejecutó — la hoja de "
    "análisis de sensibilidad. El **CSV** solo trae la tabla de ranking, y el **TXT** un "
    "resumen en texto plano de todo lo anterior."
)
if resultado_sens is None:
    st.caption("ℹ️ Aún no ejecutaste un análisis de sensibilidad; los reportes solo incluirán ranking y mejor alternativa.")
