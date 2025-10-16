# ===============================================
# PASO 1: SUBIDA Y VALIDACIÓN DEL ARCHIVO 1
# ===============================================

import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------------------------------------
# 1. Configuración inicial de la aplicación
# ------------------------------------------------
st.set_page_config(page_title="Validador de Archivos", page_icon="🗂️", layout="wide")
st.title("Validación del Archivo 1 (Nómina de Alumnos)")

# ------------------------------------------------
# 2. Definición de columnas requeridas y orden correcto
# ------------------------------------------------
COLUMNAS_REQUERIDAS = [
    "Nombre", "Paterno", "Materno", "Fecha", "Sexo",
    "Grado", "Sección", "Correo", "Neurodiversidad", "DNI"
]

# ------------------------------------------------
# 3. Inicializar el estado de paso (para flujo dinámico)
# ------------------------------------------------
if "paso" not in st.session_state:
    st.session_state.paso = 1
if "archivo1_df" not in st.session_state:
    st.session_state.archivo1_df = None

# Campo de entrada para el nombre del colegio (usado en los nombres de descarga)
if "nombre_colegio" not in st.session_state:
    st.session_state["nombre_colegio"] = ""
st.session_state["nombre_colegio"] = st.text_input(
    "Nombre del colegio",
    value=st.session_state.get("nombre_colegio", "").strip()
).strip()

# ------------------------------------------------
# 4. Función auxiliar: detección automática de cabecera
# ------------------------------------------------
def detectar_cabecera_automatica(df: pd.DataFrame, columnas_objetivo: list):
    """
    Busca entre las primeras 15 filas y 15 columnas
    la fila que contiene todas las columnas requeridas
    y en el orden correcto.
    """
    max_filas, max_cols = min(15, len(df)), min(15, len(df.columns))
    subset = df.iloc[:max_filas, :max_cols]

    columnas_objetivo_norm = [c.strip().lower() for c in columnas_objetivo]

    for idx in range(max_filas):
        fila = subset.iloc[idx].astype(str).str.strip().str.lower().tolist()
        # La fila debe contener todos los nombres requeridos
        if all(col in fila for col in columnas_objetivo_norm):
            return idx
    return None

# ------------------------------------------------
# 5. Interfaz del Paso 1: subida y validación del archivo Excel
# ------------------------------------------------
contenedor_paso1 = st.empty()

with contenedor_paso1.container():
    st.subheader("1️) Cargar archivo Excel")
    archivo_subido = st.file_uploader("Sube el archivo 1 (Excel)", type=["xls", "xlsx"])

    if archivo_subido is not None:
        try:
            df_original = pd.read_excel(archivo_subido, header=None) 
            st.info("Archivo cargado correctamente. Analizando cabecera...")

            # Intentar detección automática
            fila_detectada = detectar_cabecera_automatica(df_original, COLUMNAS_REQUERIDAS)

            if fila_detectada is not None:
                # Éxito en la detección automática
                st.success(f"Cabecera detectada automáticamente en la fila {fila_detectada + 1}.")
                df = pd.read_excel(archivo_subido, header=fila_detectada)
                
                # Normalizar columnas para coincidir y conservar sólo las requeridas
                columnas_norm = {c.strip().lower(): c for c in df.columns}
                cols_a_usar = []
                for col_req in COLUMNAS_REQUERIDAS:
                    col_norm = col_req.strip().lower()
                    if col_norm in columnas_norm:
                        cols_a_usar.append(columnas_norm[col_norm])

                df = df[cols_a_usar]
                df.columns = [col.upper() for col in COLUMNAS_REQUERIDAS]


                st.session_state.archivo1_df = df
                # Implementación de columna "IDENTIFICADOR" (para uso interno, no visible en descarga)
                df["IDENTIFICADOR"] = (
                    df["NOMBRE"].astype(str).str.strip() + " " +
                    df["PATERNO"].astype(str).str.strip() + " " +
                    df["MATERNO"].astype(str).str.strip()
                )
                st.session_state.archivo1_df = df


                st.write("Vista previa de los datos válidos:")
                st.dataframe(df.head(10), use_container_width=True)

            else:
                # Falla en la detección automática, se pasa a reconocimiento manual
                st.warning("No se encontró la cabecera automáticamente. Selecciona la fila manualmente.")
                vista_previa = df_original.iloc[:15, :15]
                st.dataframe(vista_previa, use_container_width=True)

                fila_manual = st.number_input(
                    "Indica el número de fila que contiene la cabecera (1 a 15):",
                    min_value=1, max_value=15, step=1
                )

                if st.button("Validar fila seleccionada"):
                    fila_idx = fila_manual - 1
                    fila = df_original.iloc[fila_idx].astype(str).str.strip().str.lower().tolist()
                    columnas_requeridas_norm = [c.lower() for c in COLUMNAS_REQUERIDAS]

                    if all(col in fila for col in columnas_requeridas_norm):
                        st.success("Cabecera válida encontrada manualmente.")
                        df = pd.read_excel(archivo_subido, header=fila_idx)

                        # Filtrar y renombrar igual que en la detección automática
                        columnas_norm = {c.strip().lower(): c for c in df.columns}
                        cols_a_usar = []
                        for col_req in COLUMNAS_REQUERIDAS:
                            col_norm = col_req.strip().lower()
                            if col_norm in columnas_norm:
                                cols_a_usar.append(columnas_norm[col_norm])

                        df = df[cols_a_usar]
                        df.columns = COLUMNAS_REQUERIDAS
                        st.session_state.archivo1_df = df
                        st.write("Vista previa de los datos válidos:")
                        st.dataframe(df.head(10), use_container_width=True)
                    else:
                        st.error("La fila seleccionada no contiene todas las columnas requeridas.")
                        st.info("Por favor, modifica el archivo Excel para que coincida con la estructura esperada:")
                        st.code(", ".join(COLUMNAS_REQUERIDAS), language="text")

        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo: {e}")

# ------------------------------------------------
# 6. (Opcional) Botón de descarga del archivo homologado
# ------------------------------------------------
if st.session_state.archivo1_df is not None:
    # Botón de descarga del archivo homologado
    df_descarga = st.session_state.archivo1_df.drop(columns=["IDENTIFICADOR"], errors="ignore")
    buffer = BytesIO()
    df_descarga.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    st.download_button(
        label="⬇️ Descargar Archivo 1 Homologado",
        data=buffer,
        file_name=(
            f"{st.session_state.get('nombre_colegio','').strip()}_nomina_RV.xlsx"
            if st.session_state.get('nombre_colegio','').strip()
            else "archivo1_homologado.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------------------------------------
# 7. (Opcional) Botón de continuar al siguiente paso
# ------------------------------------------------
if st.session_state.archivo1_df is not None:
    st.success("Archivo 1 validado correctamente. Puedes continuar al siguiente paso.")
    if st.button("➡️ Continuar al Paso 2"):
        st.session_state.paso = 2  # Avanza el flujo

st.divider()

# ===============================================
# PASO 2: SUBIDA Y VALIDACIÓN DEL ARCHIVO 2
# ===============================================
if st.session_state.paso == 2:
    st.title("1️) Cargar archivo Excel (Notas de Alumnos)")

    # ------------------------------------------------
    # 1. Columnas requeridas
    # ------------------------------------------------
    COLUMNAS_REQUERIDAS_2 = [
        "Nombre", "Paterno", "Materno", "Sexo", "Grado", "Curso", "Nota Vigesimal"
    ]

    # ------------------------------------------------
    # 2. Carga opcional del archivo de equivalencias (.txt)
    # ------------------------------------------------
    st.subheader("Equivalencias de Cursos (opcional)")
    
    cursos_equivalentes = [
            "ADOBE ILLUSTRATOR", "ADOBE INDESIGN", "ADOBE PHOTOSHOP PROFICIENT",
            "COACHING PERSONAL Y VOCACIONAL", "DE LA IDEA AL EMPRENDIMIENTO",
            "DESARROLLO DE APLICACIONES MÓVILES", "DESARROLLO WEB", "DISEÑO WEB",
            "EDICIÓN DE AUDIO", "EDICIÓN DE VIDEO", "EXCEL EXPERT SPECIALIST",
            "EXCEL INTERMEDIATE SPECIALIST", "EXCEL PROFICIENT SPECIALIST",
            "FINANZAS PERSONALES", "GESTIÓN DE DATA CON MS EXCEL Y POWER BI",
            "GESTIÓN EMPRESARIAL", "HABILIDADES BLANDAS", "MARKETING DIGITAL",
            "MARKETING PERSONAL", "PRESENTACIONES DE IMPACTO",
            "PROGRAMACIÓN VISUAL KODU PLANET I", "PROGRAMACIÓN VISUAL KODU PLANET II",
            "PROGRAMACIÓN VISUAL KODU PLANET III", "ROBLOX FOR TEENS", "ROBÓTICA",
            "SCRATCH", "WORD EXPERT SPECIALIST", "WORD PROFICIENT SPECIALIST",
            "LEARNING FOR BEGINNERS 1", "LEARNING FOR BEGINNERS 2", "CODE FOR KIDS"
    ]

    archivo_txt = st.file_uploader("Sube el archivo de equivalencias (.txt)", type=["txt"])

    if archivo_txt is not None:
        contenido_txt = archivo_txt.getvalue().decode("utf-8", errors="ignore")
        nuevos_cursos = [linea.strip().upper() for linea in contenido_txt.splitlines() if linea.strip()]

        cursos_equivalentes = sorted(list(set(cursos_equivalentes + nuevos_cursos)))

        st.success(f"Se agregaron {len(nuevos_cursos)} nuevos cursos. Total: {len(cursos_equivalentes)} cursos disponibles.")
    else:
        st.info("No se subió archivo .txt, se usará la lista de cursos por defecto.")

    # ------------------------------------------------
    # 3. Subida y validación del archivo Excel
    # ------------------------------------------------
    st.subheader("Cargar Archivo 2 (Notas de Cursos)")
    archivo2 = st.file_uploader("Sube el archivo 2 (Excel)", type=["xls", "xlsx"])

    if archivo2 is not None:
        try:
            df_original2 = pd.read_excel(archivo2, header=None)
            st.info("Archivo 2 cargado correctamente. Analizando cabecera...")

            fila_detectada2 = detectar_cabecera_automatica(df_original2, COLUMNAS_REQUERIDAS_2)

            if fila_detectada2 is not None:
                st.success(f"Cabecera detectada automáticamente en la fila {fila_detectada2 + 1}.")
                df2 = pd.read_excel(archivo2, header=fila_detectada2)

                # Mantener sólo las columnas requeridas
                columnas_norm = {c.strip().lower(): c for c in df2.columns}
                cols_a_usar = []
                for col_req in COLUMNAS_REQUERIDAS_2:
                    col_norm = col_req.strip().lower()
                    if col_norm in columnas_norm:
                        cols_a_usar.append(columnas_norm[col_norm])

                df2 = df2[cols_a_usar]
                df2.columns = [col.upper() for col in COLUMNAS_REQUERIDAS_2]

                # ------------------------------------------------
                # 4. Normalización de columnas según reglas del Word
                # ------------------------------------------------
                # Limpieza básica: espacios, mayúsculas, acentos
                for col in ["NOMBRE", "PATERNO", "MATERNO"]:
                    df2[col] = (
                        df2[col].astype(str)
                        .str.upper()
                        .str.normalize("NFKD")
                        .str.encode("ascii", errors="ignore")
                        .str.decode("utf-8")
                        .str.strip()
                    )

                df2["SEXO"] = df2["SEXO"].astype(str).str.upper().str.strip()
                df2["GRADO"] = df2["GRADO"].astype(str).str.upper().str.strip()

                # Curso: estandarizar a mayúsculas
                df2["CURSO"] = df2["CURSO"].astype(str).str.upper().str.strip()

                # ------------------------------------------------
                # 4. Validaciones previas a la homologación
                # ------------------------------------------------

                # 4.1 Validar si hay campos vacíos en columnas críticas
                columnas_obligatorias = ["NOMBRE", "PATERNO", "MATERNO", "SEXO", "GRADO", "CURSO", "NOTA VIGESIMAL"]
                filas_vacias = df2[df2[columnas_obligatorias].isnull().any(axis=1)]

                if not filas_vacias.empty:
                    st.error("Se detectaron filas con campos vacíos. El administrador debe completar estos campos antes de continuar.")
                    st.dataframe(filas_vacias, use_container_width=True)
                    st.stop()  # termina el flujo inmediatamente

                # 4.2 Validación de cursos no reconocidos
                cursos_invalidos = sorted(df2.loc[~df2["CURSO"].isin(cursos_equivalentes), "CURSO"].unique())

                if len(cursos_invalidos) > 0:
                    st.warning(f"⚠️ Se detectaron {len(cursos_invalidos)} cursos no reconocidos.")
                    st.info("Por favor, selecciona a qué curso oficial corresponde cada uno antes de continuar:")

                    equivalencias_usuario = {}
                    for curso in cursos_invalidos:
                        equivalencias_usuario[curso] = st.selectbox(
                            f"🔹 {curso}",
                            options=["-- Seleccionar --"] + cursos_equivalentes,
                            key=f"sel_{curso}"
                        )

                    if any(valor == "-- Seleccionar --" for valor in equivalencias_usuario.values()):
                        st.error("⚠️ Debes seleccionar un curso oficial para cada curso no reconocido antes de continuar.")
                        st.stop()
                    else:
                        # Aplicar las equivalencias seleccionadas
                        for curso_erroneo, curso_correcto in equivalencias_usuario.items():
                            df2.loc[df2["CURSO"] == curso_erroneo, "CURSO"] = curso_correcto
                        st.success("Los cursos fueron homologados correctamente según las selecciones del usuario.")

                # ------------------------------------------------
                # 5. Crear columna IDENTIFICADOR (para uso interno)
                # ------------------------------------------------
                df2["IDENTIFICADOR"] = (
                    df2["NOMBRE"].astype(str).str.strip() + " " +
                    df2["PATERNO"].astype(str).str.strip() + " " +
                    df2["MATERNO"].astype(str).str.strip()
                )

                # Crear columnas adicionales antes del IDENTIFICADOR
                df2["NOTAS 01"] = ""
                df2["NOTAS 02"] = ""

                # Eliminar la columna de IDENTIFICADOR para la descarga
                columnas = [col for col in df2.columns if col != "IDENTIFICADOR"]
                columnas.append("IDENTIFICADOR")
                df2 = df2[columnas]

                st.session_state.archivo2_df = df2
                st.success("Archivo 2 validado y normalizado correctamente.")
                st.dataframe(df2.head(10), use_container_width=True)

                # ------------------------------------------------
                # 6. Botón de descarga
                # ------------------------------------------------
                # ARCHIVO 2: sin columnas de notas
                from io import BytesIO
                df_sin_notas = df2.drop(columns=["IDENTIFICADOR", "NOTAS 01", "NOTAS 02"], errors="ignore")
                buffer2_sin = BytesIO()
                df_sin_notas.to_excel(buffer2_sin, index=False, engine="openpyxl")
                buffer2_sin.seek(0)
                st.download_button(
                    label="⬇️ Descargar Archivo 2 Homologado",
                    data=buffer2_sin,
                    file_name=(
                        f"{st.session_state.get('nombre_colegio','').strip()}_notas_RV.xlsx"
                        if st.session_state.get('nombre_colegio','').strip()
                        else "archivo2_homologado.xlsx"
                    ),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # ARCHIVO 3: con columnas de notas
                df_con_notas = df2.drop(columns=["IDENTIFICADOR"], errors="ignore")
                buffer_con = BytesIO()
                df_con_notas.to_excel(buffer_con, index=False, engine="openpyxl")
                buffer_con.seek(0)
                st.download_button(
                    label="⬇️ Descargar Archivo 3 (Evaluador)",
                    data=buffer_con,
                    file_name=(
                        f"{st.session_state.get('nombre_colegio','').strip()}_evaluador.xlsx"
                        if st.session_state.get('nombre_colegio','').strip()
                        else "archivo3_evaluador.xlsx"
                    ),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.warning("No se detectó la cabecera automáticamente. Selección manual disponible.")
                vista_previa = df_original2.iloc[:15, :15]
                st.dataframe(vista_previa, use_container_width=True)

                fila_manual2 = st.number_input(
                    "Indica el número de fila que contiene la cabecera (1 a 15):",
                    min_value=1, max_value=15, step=1, key="fila_manual_2"
                )

                if st.button("Validar fila seleccionada del archivo 2"):
                    fila_idx2 = fila_manual2 - 1
                    fila = df_original2.iloc[fila_idx2].astype(str).str.strip().str.lower().tolist()
                    columnas_requeridas_norm = [c.lower() for c in COLUMNAS_REQUERIDAS_2]

                    if all(col in fila for col in columnas_requeridas_norm):
                        st.success("✅ Cabecera válida encontrada manualmente.")
                        df2 = pd.read_excel(archivo2, header=fila_idx2)

                        # Mantener sólo las columnas requeridas
                        columnas_norm = {c.strip().lower(): c for c in df2.columns}
                        cols_a_usar = []
                        for col_req in COLUMNAS_REQUERIDAS_2:
                            col_norm = col_req.strip().lower()
                            if col_norm in columnas_norm:
                                cols_a_usar.append(columnas_norm[col_norm])

                        df2 = df2[cols_a_usar]
                        df2.columns = [col.upper() for col in COLUMNAS_REQUERIDAS_2]

                        # ------------------------------------------------
                        # Normalización de columnas
                        # ------------------------------------------------
                        for col in ["NOMBRE", "PATERNO", "MATERNO"]:
                            df2[col] = (
                                df2[col].astype(str)
                                .str.upper()
                                .str.normalize("NFKD")
                                .str.encode("ascii", errors="ignore")
                                .str.decode("utf-8")
                                .str.strip()
                            )

                        df2["SEXO"] = df2["SEXO"].astype(str).str.upper().str.strip()
                        df2["GRADO"] = df2["GRADO"].astype(str).str.upper().str.strip()
                        df2["CURSO"] = df2["CURSO"].astype(str).str.upper().str.strip()

                        # ------------------------------------------------
                        # 4.1 Validación interactiva de cursos no reconocidos
                        # ------------------------------------------------
                        cursos_invalidos = sorted(df2.loc[~df2["CURSO"].isin(cursos_equivalentes), "CURSO"].unique())

                        if len(cursos_invalidos) > 0:
                            st.warning(f"⚠️ Se detectaron {len(cursos_invalidos)} cursos no reconocidos.")
                            st.info("Por favor, selecciona a qué curso oficial corresponde cada uno antes de continuar:")

                            # Diccionario temporal para almacenar las equivalencias elegidas
                            equivalencias_usuario = {}

                            for curso in cursos_invalidos:
                                equivalencias_usuario[curso] = st.selectbox(
                                    f"🔹 {curso}",
                                    options=["-- Seleccionar --"] + cursos_equivalentes,
                                    key=f"sel_{curso}"
                                )

                            # Verificar si el usuario ha completado todas las selecciones
                            if any(valor == "-- Seleccionar --" for valor in equivalencias_usuario.values()):
                                st.error("⚠️ Debes seleccionar un curso oficial para cada curso no reconocido antes de continuar.")
                                st.stop()
                            else:
                                # Aplicar las equivalencias seleccionadas
                                for curso_erroneo, curso_correcto in equivalencias_usuario.items():
                                    df2.loc[df2["CURSO"] == curso_erroneo, "CURSO"] = curso_correcto
                                st.success("Los cursos fueron homologados correctamente según las selecciones del usuario.")

                        # Nota vigesimal: convertir a número y validar rango
                        df2["NOTA VIGESIMAL"] = pd.to_numeric(df2["NOTA VIGESIMAL"], errors="coerce")
                        df2.loc[df2["NOTA VIGESIMAL"].isna(), "NOTA VIGESIMAL"] = "NP"
                        df2.loc[
                            (df2["NOTA VIGESIMAL"] != "NP")
                            & ((df2["NOTA VIGESIMAL"] < 0) | (df2["NOTA VIGESIMAL"] > 20)),
                            "NOTA VIGESIMAL"
                        ] = "NP"

                        # ------------------------------------------------
                        # Agregar columna IDENTIFICADOR (solo para sistema)
                        # ------------------------------------------------
                        df2["IDENTIFICADOR"] = (
                            df2["NOMBRE"].astype(str).str.strip() + " " +
                            df2["PATERNO"].astype(str).str.strip() + " " +
                            df2["MATERNO"].astype(str).str.strip()
                        )

                        # Crear columnas adicionales antes del IDENTIFICADOR
                        df2["NOTAS 01"] = ""
                        df2["NOTAS 02"] = ""

                        # Eliminar la columna de IDENTIFICADOR para la descarga
                        columnas = [col for col in df2.columns if col != "IDENTIFICADOR"]
                        columnas.append("IDENTIFICADOR")
                        df2 = df2[columnas]

                        st.session_state.archivo2_df = df2
                        st.success("Archivo 2 validado y normalizado correctamente (modo manual).")
                        st.dataframe(df2.head(10), use_container_width=True)

                        # ------------------------------------------------
                        # Botón de descarga
                        # ------------------------------------------------
                        # ARCHIVO 2: sin columnas de notas
                        from io import BytesIO
                        df_sin_notas = df2.drop(columns=["IDENTIFICADOR", "NOTAS 01", "NOTAS 02"], errors="ignore")
                        buffer2_sin = BytesIO()
                        df_sin_notas.to_excel(buffer2_sin, index=False, engine="openpyxl")
                        buffer2_sin.seek(0)
                        st.download_button(
                            label="⬇️ Descargar Archivo 2 Homologado",
                            data=buffer2_sin,
                            file_name=(
                                f"{st.session_state.get('nombre_colegio','').strip()}_notas_RV.xlsx"
                                if st.session_state.get('nombre_colegio','').strip()
                                else "archivo2_homologado.xlsx"
                            ),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                        # ARCHIVO 3: con columnas de notas
                        df_con_notas = df2.drop(columns=["IDENTIFICADOR"], errors="ignore")
                        buffer_con = BytesIO()
                        df_con_notas.to_excel(buffer_con, index=False, engine="openpyxl")
                        buffer_con.seek(0)
                        st.download_button(
                            label="⬇️ Descargar Archivo 3 (Evaluador)",
                            data=buffer_con,
                            file_name=(
                                f"{st.session_state.get('nombre_colegio','').strip()}_evaluador.xlsx"
                                if st.session_state.get('nombre_colegio','').strip()
                                else "archivo3_evaluador.xlsx"
                            ),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                    else:
                        st.error("La fila seleccionada no contiene todas las columnas requeridas.")
                        st.info("Por favor, modifica el archivo Excel para que coincida con esta estructura:")
                        st.code(", ".join(COLUMNAS_REQUERIDAS_2), language="text")

        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo 2: {e}")
