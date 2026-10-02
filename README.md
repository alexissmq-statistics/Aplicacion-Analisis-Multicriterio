# Aplicación de Análisis Multicriterio (AHP)

Herramienta interactiva construida con Streamlit para la toma de decisiones multicriterio usando el **Proceso Analítico Jerárquico (AHP)**

🔗 **Aplicación en línea:** [multicriterio-saty.streamlit.app](https://multicriterio-saty.streamlit.app/)

## ¿Qué hace la aplicación?

Permite comparar alternativas frente a varios criterios y obtener una priorización final, soportando dos modos de trabajo:

- **Modo objetivo (costo/beneficio):** ingreso de datos cuantitativos directamente en una grilla editable (`st.data_editor`), útil cuando ya se cuenta con valores medibles por criterio.
- **Modo subjetivo (AHP clásico):** comparación por pares con la **escala de Saaty**, mediante una interfaz de matriz interactiva, para cuando los criterios se evalúan de forma cualitativa/experta.

La interfaz incluye un botón de reinicio en la barra lateral para limpiar la sesión y empezar una nueva evaluación sin recargar la página.

## Estructura del proyecto

```
.
├── app.py              # Interfaz principal de Streamlit
├── ahp_logic.py         # Lógica del método AHP (matrices, pesos, consistencia)
├── requirements.txt     # Dependencias del proyecto
```

## Ejecutar localmente

```bash
# 1. Clonar el repositorio
git clone <url-de-tu-repo>
cd <nombre-del-repo>

# 2. Crear y activar un entorno virtual
python -m venv venv
venv\Scripts\activate      # En Windows
source venv/bin/activate   # En Linux/Mac

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar la aplicación
streamlit run app.py
```

En Windows también puedes usar directamente `run_app.bat` para lanzar la app sin escribir el comando manualmente.

## Despliegue en Streamlit Cloud

La app está desplegada en Streamlit Community Cloud apuntando a este repositorio. La versión de Python usada en producción se fija explícitamente (Python 3.11) para evitar incompatibilidades con versiones muy recientes de Python en las que algunas dependencias científicas (como `matplotlib`, `numpy` o `pyarrow`) aún no tienen soporte estable.

## Metodología

El Proceso Analítico Jerárquico (AHP) estructura la decisión en: definición de criterios y alternativas, comparación por pares (escala 1-9 de Saaty), cálculo de pesos relativos mediante el autovector principal de la matriz de comparación, y verificación de la **razón de consistencia (CR)** para validar que los juicios sean coherentes.
