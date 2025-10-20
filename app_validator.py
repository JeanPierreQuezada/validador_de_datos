# SN: Sin nomina
# NA: Nomina ausente
# NP: No se presento
# Nota laboratorio (25)
# Nota evaluador (75)
# Promedio Final

# ===============================================
# PASO 1: SUBIDA Y VALIDACIÓN DEL ARCHIVO 1
# ===============================================

import streamlit as st
import pandas as pd
from io import BytesIO

# ================================================
# CONFIGURACIÓN INICIAL
# ================================================
st.set_page_config(
    page_title="Validador de Archivos",
    page_icon="📊",
    layout="wide"
)

# ================================================
# INICIALIZACIÓN DE ESTADOS
# ================================================
if "paso_actual" not in st.session_state:
    st.session_state.paso_actual = 0
if "nombre_colegio" not in st.session_state:
    st.session_state.nombre_colegio = ""
if "archivo1_df" not in st.session_state:
    st.session_state.archivo1_df = None
if "archivo2_df" not in st.session_state:
    st.session_state.archivo2_df = None
if "archivo2_1p3p_df" not in st.session_state:
    st.session_state.archivo2_1p3p_df = None
if "archivo2_4p5s_df" not in st.session_state:
    st.session_state.archivo2_4p5s_df = None
if "cursos_equivalentes" not in st.session_state:
    st.session_state.cursos_equivalentes = [
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

# ================================================
# CONSTANTES
# ================================================
COLUMNAS_ARCHIVO1 = [
    "Nombre", "Paterno", "Materno", "Fecha", "Sexo",
    "Grado", "Sección", "Correo", "Neurodiversidad", "DNI"
]

COLUMNAS_ARCHIVO2 = [
    "Nombre", "Paterno", "Materno", "Sexo", "Grado", "Curso", "Nota Vigesimal"
]

# ================================================
# FUNCIONES AUXILIARES
# ================================================
def detectar_cabecera_automatica(df: pd.DataFrame, columnas_objetivo: list):
    """Detecta automáticamente la fila de cabecera"""
    max_filas, max_cols = min(15, len(df)), min(15, len(df.columns))
    subset = df.iloc[:max_filas, :max_cols]
    columnas_objetivo_norm = [c.strip().lower() for c in columnas_objetivo]

    for idx in range(max_filas):
        fila = subset.iloc[idx].astype(str).str.strip().str.lower().tolist()
        if all(col in fila for col in columnas_objetivo_norm):
            return idx
    return None

def crear_identificador(df, col_nombre, col_paterno, col_materno):
    """Crea columna identificador"""
    return (
        df[col_nombre].astype(str).str.strip() + " " +
        df[col_paterno].astype(str).str.strip() + " " +
        df[col_materno].astype(str).str.strip()
    )

def normalizar_columnas_texto(df, columnas):
    """Normaliza columnas de texto (mayúsculas, sin acentos)"""
    for col in columnas:
        df[col] = (
            df[col].astype(str)
            .str.upper()
            .str.normalize("NFKD")
            .str.encode("ascii", errors="ignore")
            .str.decode("utf-8")
            .str.strip()
        )
    return df

def mostrar_stepper(paso_actual):
    """Muestra el indicador de progreso visual"""
    pasos = [
        {"num": 0, "titulo": "Nombre del Colegio", "icono": "🏫"},
        {"num": 1, "titulo": "Archivo 1: Nómina", "icono": "📋"},
        {"num": 2, "titulo": "Archivo 2: Notas", "icono": "📊"},
        {"num": 3, "titulo": "Descarga Final", "icono": "⬇️"}
    ]
    
    cols = st.columns(len(pasos))
    for i, paso in enumerate(pasos):
        with cols[i]:
            if paso["num"] < paso_actual:
                st.markdown(f"### ✅ {paso['icono']}")
                st.markdown(f"**{paso['titulo']}**")
                st.markdown("*Completado*")
            elif paso["num"] == paso_actual:
                st.markdown(f"### 🔵 {paso['icono']}")
                st.markdown(f"**{paso['titulo']}**")
                st.markdown("*En progreso*")
            else:
                st.markdown(f"### ⚪ {paso['icono']}")
                st.markdown(f"<span style='color: gray;'>{paso['titulo']}</span>", unsafe_allow_html=True)
                st.markdown("*Pendiente*")
    
    st.divider()

# ================================================
# INTERFAZ PRINCIPAL
# ================================================
st.title("🎯 Validador de Archivos Escolares")
st.markdown("### Sistema de Homologación de Datos")

# Mostrar stepper
mostrar_stepper(st.session_state.paso_actual)

# ================================================
# PASO 0: NOMBRE DEL COLEGIO
# ================================================
if st.session_state.paso_actual == 0:
    st.header("🏫 Paso 1: Información del Colegio")
    
    with st.container():
        st.markdown("""
        <div style='background-color: #78808C; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
            <h4>Bienvenido al sistema de validación</h4>
            <p>Para comenzar, ingresa el nombre del colegio. Este nombre se usará para identificar los archivos descargables.</p>
        </div>
        """, unsafe_allow_html=True)
        
        nombre = st.text_input(
            "Nombre del Colegio",
            value=st.session_state.nombre_colegio,
            placeholder="Ejemplo: Colegio San Martín de Porres",
            help="Este nombre aparecerá en los archivos descargados"
        )
        
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            if st.button("➡️ Continuar", type="primary", use_container_width=True):
                if nombre.strip():
                    st.session_state.nombre_colegio = nombre.strip()
                    st.session_state.paso_actual = 1
                    st.rerun()
                else:
                    st.error("Por favor, ingresa el nombre del colegio")

# ================================================
# PASO 1: ARCHIVO 1 (NÓMINA)
# ================================================
elif st.session_state.paso_actual == 1:
    # Mostrar resumen del paso anterior
    with st.expander("✅ Paso 1 completado: Nombre del Colegio", expanded=False):
        st.info(f"**Colegio:** {st.session_state.nombre_colegio}")
        if st.button("🔄 Cambiar nombre", key="cambiar_nombre"):
            st.session_state.paso_actual = 0
            st.rerun()
    
    st.header("📋 Paso 2: Archivo de Nómina de Alumnos")
    
    with st.container():
        st.markdown("""
        <div style='background-color: #78808C; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
            <h4>📄 Instrucciones</h4>
            <p>Sube el archivo Excel que contiene la nómina de alumnos.</p>
            <p><strong>Columnas requeridas:</strong></p>
            <code>Nombre, Paterno, Materno, Fecha, Sexo, Grado, Sección, Correo, Neurodiversidad, DNI</code>
        </div>
        """, unsafe_allow_html=True)
        
        archivo = st.file_uploader(
            "Selecciona el archivo Excel",
            type=["xls", "xlsx"],
            help="El sistema detectará automáticamente la fila de cabecera"
        )
        
        if archivo is not None:
            with st.spinner("🔍 Analizando archivo..."):
                try:
                    df_original = pd.read_excel(archivo, header=None)
                    fila_detectada = detectar_cabecera_automatica(df_original, COLUMNAS_ARCHIVO1)
                    
                    if fila_detectada is not None:
                        st.success(f"✅ Cabecera detectada automáticamente en la fila {fila_detectada + 1}")
                        
                        df = pd.read_excel(archivo, header=fila_detectada)
                        
                        # Procesar columnas
                        columnas_norm = {c.strip().lower(): c for c in df.columns}
                        cols_a_usar = []
                        for col_req in COLUMNAS_ARCHIVO1:
                            col_norm = col_req.strip().lower()
                            if col_norm in columnas_norm:
                                cols_a_usar.append(columnas_norm[col_norm])
                        
                        df = df[cols_a_usar]
                        df.columns = [col.upper() for col in COLUMNAS_ARCHIVO1]
                        df["IDENTIFICADOR"] = crear_identificador(df, "NOMBRE", "PATERNO", "MATERNO")
                        
                        st.session_state.archivo1_df = df
                        
                        # Mostrar preview
                        st.markdown("### 📊 Vista Previa de Datos")
                        st.info(f"Total de registros: {len(df)}")
                        st.dataframe(df.head(10), use_container_width=True)
                        
                        # Botones de acción
                        col1, col2 = st.columns(2)
                        with col1:
                            df_descarga = df.drop(columns=["IDENTIFICADOR"], errors="ignore")
                            buffer = BytesIO()
                            df_descarga.to_excel(buffer, index=False, engine="openpyxl")
                            buffer.seek(0)
                            st.download_button(
                                label="💾 Descargar Archivo Homologado",
                                data=buffer,
                                file_name=f"{st.session_state.nombre_colegio}_nomina_RV.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        with col2:
                            if st.button("➡️ Continuar al Paso 3", type="primary", use_container_width=True):
                                st.session_state.paso_actual = 2
                                st.rerun()
                    
                    else:
                        st.warning("⚠️ No se pudo detectar la cabecera automáticamente")
                        st.markdown("### 🔍 Detección Manual")
                        st.dataframe(df_original.iloc[:15, :15], use_container_width=True)
                        
                        fila_manual = st.number_input(
                            "Indica el número de fila que contiene la cabecera:",
                            min_value=1, max_value=15, step=1
                        )
                        
                        if st.button("✔️ Validar Fila Seleccionada", type="primary"):
                            fila_idx = fila_manual - 1
                            fila = df_original.iloc[fila_idx].astype(str).str.strip().str.lower().tolist()
                            columnas_req_norm = [c.lower() for c in COLUMNAS_ARCHIVO1]
                            
                            if all(col in fila for col in columnas_req_norm):
                                st.success("✅ Cabecera válida")
                                df = pd.read_excel(archivo, header=fila_idx)
                                
                                columnas_norm = {c.strip().lower(): c for c in df.columns}
                                cols_a_usar = []
                                for col_req in COLUMNAS_ARCHIVO1:
                                    col_norm = col_req.strip().lower()
                                    if col_norm in columnas_norm:
                                        cols_a_usar.append(columnas_norm[col_norm])
                                
                                df = df[cols_a_usar]
                                df.columns = [col.upper() for col in COLUMNAS_ARCHIVO1]
                                df["IDENTIFICADOR"] = crear_identificador(df, "NOMBRE", "PATERNO", "MATERNO")
                                
                                st.session_state.archivo1_df = df
                                st.rerun()
                            else:
                                st.error("❌ La fila seleccionada no contiene todas las columnas requeridas")
                
                except Exception as e:
                    st.error(f"❌ Error al procesar el archivo: {e}")

# ================================================
# PASO 2: ARCHIVO 2 (NOTAS)
# ================================================
elif st.session_state.paso_actual == 2:
    # Mostrar resumen de pasos anteriores
    with st.expander("✅ Pasos completados", expanded=False):
        st.success(f"**Colegio:** {st.session_state.nombre_colegio}")
        st.success(f"**Archivo 1:** {len(st.session_state.archivo1_df)} registros cargados")
        if st.button("🔙 Volver al Paso 2", key="volver_paso2"):
            st.session_state.paso_actual = 1
            st.rerun()
    
    st.header("📊 Paso 3: Archivo de Notas de Cursos")
    
    # Equivalencias de cursos
    with st.expander("⚙️ Configuración de Cursos Equivalentes", expanded=False):
        st.markdown("""
        <div style='background-color: #78808C; padding: 15px; border-radius: 10px;'>
            <p>Opcionalmente, puedes cargar un archivo .txt con cursos adicionales para reconocimiento automático.</p>
        </div>
        """, unsafe_allow_html=True)
        
        archivo_txt = st.file_uploader("Archivo de equivalencias (.txt)", type=["txt"])
        if archivo_txt:
            contenido = archivo_txt.getvalue().decode("utf-8", errors="ignore")
            nuevos = [l.strip().upper() for l in contenido.splitlines() if l.strip()]
            st.session_state.cursos_equivalentes = sorted(list(set(st.session_state.cursos_equivalentes + nuevos)))
            st.success(f"✅ {len(nuevos)} cursos agregados. Total: {len(st.session_state.cursos_equivalentes)}")
    
    # Carga del archivo
    st.markdown("""
    <div style='background-color: #78808C; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
        <h4>📄 Instrucciones</h4>
        <p>Sube el archivo Excel con las notas de los cursos.</p>
        <p><strong>Columnas requeridas:</strong></p>
        <code>Nombre, Paterno, Materno, Sexo, Grado, Curso, Nota Vigesimal</code>
    </div>
    """, unsafe_allow_html=True)
    
    archivo2 = st.file_uploader("Selecciona el archivo Excel de notas", type=["xls", "xlsx"])
    
    if archivo2 is not None:
        with st.spinner("🔍 Analizando archivo y hojas disponibles..."):
            try:
                # Leer el archivo para detectar hojas
                xls_file = pd.ExcelFile(archivo2)
                hojas_disponibles = xls_file.sheet_names
                
                # Detectar qué hojas existen
                tiene_1p3p = "1P-3P" in hojas_disponibles
                tiene_4p5s = "4P-5S" in hojas_disponibles
                
                if not tiene_1p3p and not tiene_4p5s:
                    st.error("❌ El archivo no contiene ninguna de las hojas requeridas: '1P-3P' o '4P-5S'")
                    st.info(f"Hojas encontradas: {', '.join(hojas_disponibles)}")
                    st.stop()
                
                # Mostrar información de hojas detectadas
                st.success(f"✅ Hojas detectadas en el archivo:")
                cols_info = st.columns(2)
                with cols_info[0]:
                    if tiene_1p3p:
                        st.info("📘 **1P-3P** encontrada (se convertirá a mayúsculas)")
                with cols_info[1]:
                    if tiene_4p5s:
                        st.info("📗 **4P-5S** encontrada (se homologará completamente)")
                
                st.divider()
                
                # ====================================
                # PROCESAR HOJA 1P-3P (Solo mayúsculas)
                # ====================================
                df_1p3p_procesado = None
                if tiene_1p3p:
                    st.markdown("### 📘 Procesando Hoja: 1P-3P")
                    
                    df_1p3p_original = pd.read_excel(archivo2, sheet_name="1P-3P", header=None)
                    fila_detectada_1p3p = detectar_cabecera_automatica(df_1p3p_original, COLUMNAS_ARCHIVO2)
                    
                    if fila_detectada_1p3p is not None:
                        st.success(f"✅ Cabecera detectada en la fila {fila_detectada_1p3p + 1}")
                        
                        df_1p3p = pd.read_excel(archivo2, sheet_name="1P-3P", header=fila_detectada_1p3p)
                        
                        # Procesar columnas
                        columnas_norm = {c.strip().lower(): c for c in df_1p3p.columns}
                        cols_a_usar = []
                        for col_req in COLUMNAS_ARCHIVO2:
                            col_norm = col_req.strip().lower()
                            if col_norm in columnas_norm:
                                cols_a_usar.append(columnas_norm[col_norm])
                        
                        df_1p3p = df_1p3p[cols_a_usar]
                        df_1p3p.columns = [col.upper() for col in COLUMNAS_ARCHIVO2]
                        
                        # Solo convertir todo a mayúsculas
                        for col in df_1p3p.columns:
                            df_1p3p[col] = df_1p3p[col].astype(str).str.upper().str.strip()
                        
                        # Validar cursos en 1P-3P
                        cursos_invalidos_1p3p = sorted(df_1p3p.loc[~df_1p3p["CURSO"].isin(st.session_state.cursos_equivalentes), "CURSO"].unique())
                        
                        if len(cursos_invalidos_1p3p) > 0:
                            st.warning(f"⚠️ Se detectaron {len(cursos_invalidos_1p3p)} cursos no reconocidos en 1P-3P")
                            
                            with st.form("equivalencias_form_1p3p"):
                                st.markdown("### 🔄 Homologación de Cursos (1P-3P)")
                                st.info("Selecciona el curso oficial correspondiente para cada curso no reconocido:")
                                
                                equivalencias_1p3p = {}
                                for curso in cursos_invalidos_1p3p:
                                    equivalencias_1p3p[curso] = st.selectbox(
                                        f"📌 **{curso}**",
                                        options=["-- Seleccionar --"] + st.session_state.cursos_equivalentes,
                                        key=f"eq_1p3p_{curso}"
                                    )
                                
                                submitted_1p3p = st.form_submit_button("✔️ Aplicar Equivalencias (1P-3P)", type="primary")
                                
                                if submitted_1p3p:
                                    if any(v == "-- Seleccionar --" for v in equivalencias_1p3p.values()):
                                        st.error("❌ Debes seleccionar un curso para todos los campos")
                                    else:
                                        # Aplicar equivalencias
                                        for curso_err, curso_ok in equivalencias_1p3p.items():
                                            df_1p3p.loc[df_1p3p["CURSO"] == curso_err, "CURSO"] = curso_ok
                                        
                                        # Agregar solo IDENTIFICADOR (NO agregar NOTAS VIGESIMALES 75% y PROMEDIO)
                                        df_1p3p["IDENTIFICADOR"] = crear_identificador(df_1p3p, "NOMBRE", "PATERNO", "MATERNO")
                                        
                                        # Reordenar
                                        cols_orden = [c for c in df_1p3p.columns if c != "IDENTIFICADOR"]
                                        cols_orden.append("IDENTIFICADOR")
                                        df_1p3p = df_1p3p[cols_orden]
                                        
                                        # Guardar en session_state
                                        st.session_state.archivo2_1p3p_df = df_1p3p
                                        
                                        st.success("✅ Cursos homologados correctamente en 1P-3P")
                                        st.rerun()
                        
                        # Si no hay cursos inválidos o ya fueron procesados
                        if len(cursos_invalidos_1p3p) == 0 or st.session_state.archivo2_1p3p_df is not None:
                            # Usar el DataFrame guardado si existe, sino usar el actual
                            if st.session_state.archivo2_1p3p_df is not None:
                                df_1p3p = st.session_state.archivo2_1p3p_df
                            else:
                                # Agregar solo IDENTIFICADOR (NO agregar NOTAS VIGESIMALES 75% y PROMEDIO)
                                df_1p3p["IDENTIFICADOR"] = crear_identificador(df_1p3p, "NOMBRE", "PATERNO", "MATERNO")
                                
                                # Reordenar
                                cols_orden = [c for c in df_1p3p.columns if c != "IDENTIFICADOR"]
                                cols_orden.append("IDENTIFICADOR")
                                df_1p3p = df_1p3p[cols_orden]
                                
                                st.session_state.archivo2_1p3p_df = df_1p3p
                            
                            df_1p3p_procesado = df_1p3p
                            
                            with st.expander("Vista previa 1P-3P", expanded=False):
                                st.info(f"Total de registros: {len(df_1p3p)}")
                                st.dataframe(df_1p3p, use_container_width=True)
                    else:
                        st.warning("⚠️ No se pudo detectar cabecera automáticamente en 1P-3P")
                        st.info("Por favor, verifica que la hoja tenga las columnas correctas")
                
                # ====================================
                # PROCESAR HOJA 4P-5S (Homologación completa)
                # ====================================
                df_4p5s_procesado = None
                if tiene_4p5s:
                    st.markdown("### 📗 Procesando Hoja: 4P-5S")
                    
                    df_original2 = pd.read_excel(archivo2, sheet_name="4P-5S", header=None)
                    fila_detectada2 = detectar_cabecera_automatica(df_original2, COLUMNAS_ARCHIVO2)
                    
                    if fila_detectada2 is not None:
                        st.success(f"✅ Cabecera detectada en la fila {fila_detectada2 + 1}")
                        
                        df2 = pd.read_excel(archivo2, sheet_name="4P-5S", header=fila_detectada2)
                    
                    # Procesar columnas
                    columnas_norm = {c.strip().lower(): c for c in df2.columns}
                    cols_a_usar = []
                    for col_req in COLUMNAS_ARCHIVO2:
                        col_norm = col_req.strip().lower()
                        if col_norm in columnas_norm:
                            cols_a_usar.append(columnas_norm[col_norm])
                    
                    df2 = df2[cols_a_usar]
                    df2.columns = [col.upper() for col in COLUMNAS_ARCHIVO2]
                    
                    # Normalizar
                    df2 = normalizar_columnas_texto(df2, ["NOMBRE", "PATERNO", "MATERNO"])
                    df2["SEXO"] = df2["SEXO"].astype(str).str.upper().str.strip()
                    df2["GRADO"] = df2["GRADO"].astype(str).str.upper().str.strip()
                    df2["CURSO"] = df2["CURSO"].astype(str).str.upper().str.strip()
                    
                    # Validar campos vacíos
                    columnas_oblig = ["NOMBRE", "PATERNO", "MATERNO", "SEXO", "GRADO", "CURSO", "NOTA VIGESIMAL"]
                    filas_vacias = df2[df2[columnas_oblig].isnull().any(axis=1)]
                    
                    if not filas_vacias.empty:
                        st.error("❌ Se detectaron campos vacíos")
                        st.dataframe(filas_vacias, use_container_width=True)
                        st.stop()
                    
                    # Validar cursos
                    cursos_invalidos = sorted(df2.loc[~df2["CURSO"].isin(st.session_state.cursos_equivalentes), "CURSO"].unique())
                    
                    if len(cursos_invalidos) > 0:
                        st.warning(f"⚠️ Se detectaron {len(cursos_invalidos)} cursos no reconocidos")
                        
                        with st.form("equivalencias_form"):
                            st.markdown("### 🔄 Homologación de Cursos")
                            st.info("Selecciona el curso oficial correspondiente para cada curso no reconocido:")
                            
                            equivalencias = {}
                            for curso in cursos_invalidos:
                                equivalencias[curso] = st.selectbox(
                                    f"📌 **{curso}**",
                                    options=["-- Seleccionar --"] + st.session_state.cursos_equivalentes,
                                    key=f"eq_{curso}"
                                )
                            
                            submitted = st.form_submit_button("✔️ Aplicar Equivalencias", type="primary")
                            
                            if submitted:
                                if any(v == "-- Seleccionar --" for v in equivalencias.values()):
                                    st.error("❌ Debes seleccionar un curso para todos los campos")
                                else:
                                    # Aplicar equivalencias
                                    for curso_err, curso_ok in equivalencias.items():
                                        df2.loc[df2["CURSO"] == curso_err, "CURSO"] = curso_ok
                                    
                                    # CRÍTICO: Guardar en session_state ANTES del rerun
                                    df2["IDENTIFICADOR"] = crear_identificador(df2, "NOMBRE", "PATERNO", "MATERNO")
                                    df2["NOTAS VIGESIMALES 75%"] = ""
                                    df2["PROMEDIO"] = ""
                                    
                                    # Reordenar columnas
                                    cols_orden = [c for c in df2.columns if c != "IDENTIFICADOR"]
                                    cols_orden.append("IDENTIFICADOR")
                                    df2 = df2[cols_orden]
                                    
                                    # Guardar en session_state
                                    st.session_state.archivo2_df = df2
                                    
                                    st.success("✅ Cursos homologados correctamente")
                                    st.rerun()
                    
                    # Si no hay cursos inválidos o ya fueron procesados
                    if len(cursos_invalidos) == 0 or st.session_state.archivo2_df is not None:
                        df2["IDENTIFICADOR"] = crear_identificador(df2, "NOMBRE", "PATERNO", "MATERNO")
                        df2["NOTAS VIGESIMALES 75%"] = ""
                        df2["PROMEDIO"] = ""
                        
                        # Reordenar columnas
                        cols_orden = [c for c in df2.columns if c != "IDENTIFICADOR"]
                        cols_orden.append("IDENTIFICADOR")
                        df2 = df2[cols_orden]
                        
                        st.session_state.archivo2_df = df2
                        
                        df_4p5s_procesado = df2
                        
                    else:
                        st.warning("⚠️ No se pudo detectar cabecera automáticamente en 4P-5S")
                        st.info("Por favor, verifica que la hoja tenga las columnas correctas")
                
                # ====================================
                # SECCIÓN DE DESCARGA
                # ====================================
                if df_1p3p_procesado is not None or df_4p5s_procesado is not None:
                    st.divider()
                    st.markdown("### 💾 Archivos Listos para Descargar")
                    
                    # Crear columnas dinámicamente según hojas disponibles
                    num_descargas = 0
                    if df_1p3p_procesado is not None:
                        num_descargas += 1  # solo 1 archivo para 1P-3P
                    if df_4p5s_procesado is not None:
                        num_descargas += 2  # archivo sin notas + evaluador para 4P-5S
                    
                    # Diseño dinámico
                    if num_descargas == 1:
                        cols_descarga = st.columns([1, 1])  # descarga + paso siguiente
                    else:
                        cols_descarga = st.columns(min(num_descargas, 3))

                    col_idx = 0
                    
                    # Descargas para 1P-3P
                    if df_1p3p_procesado is not None:
                        with cols_descarga[col_idx]:
                            # Para 1P-3P (no hay NOTAS VIGESIMALES 75% ni PROMEDIO)
                            df_sin_notas_1p3p = df_1p3p_procesado.drop(columns=["IDENTIFICADOR"], errors="ignore")
                            buffer_1p3p = BytesIO()
                            df_sin_notas_1p3p.to_excel(buffer_1p3p, index=False, engine="openpyxl")
                            buffer_1p3p.seek(0)
                            st.download_button(
                                label="📥 1P-3P Homologado",
                                data=buffer_1p3p,
                                file_name=f"{st.session_state.nombre_colegio}_1P-3P_RV.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        col_idx += 1
                    
                    # Descargas para 4P-5S
                    if df_4p5s_procesado is not None:
                        if col_idx < len(cols_descarga):
                            with cols_descarga[col_idx]:
                                df_sin_notas_4p5s = df_4p5s_procesado.drop(columns=["IDENTIFICADOR", "NOTAS VIGESIMALES 75%", "PROMEDIO"], errors="ignore")
                                buffer_4p5s = BytesIO()
                                df_sin_notas_4p5s.to_excel(buffer_4p5s, index=False, engine="openpyxl")
                                buffer_4p5s.seek(0)
                                st.download_button(
                                    label="📥 4P-5S Homologado",
                                    data=buffer_4p5s,
                                    file_name=f"{st.session_state.nombre_colegio}_4P-5S_RV.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        col_idx += 1
                        
                        if col_idx < len(cols_descarga):
                            with cols_descarga[col_idx]:
                                df_eval_4p5s = df_4p5s_procesado.drop(columns=["IDENTIFICADOR"], errors="ignore")
                                buffer_eval_4p5s = BytesIO()
                                df_eval_4p5s.to_excel(buffer_eval_4p5s, index=False, engine="openpyxl")
                                buffer_eval_4p5s.seek(0)
                                st.download_button(
                                    label="📥 4P-5S Evaluador",
                                    data=buffer_eval_4p5s,
                                    file_name=f"{st.session_state.nombre_colegio}_4P-5S_evaluador.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                    
                    # --- Si solo hay un botón, muestra "Paso siguiente" al lado ---
                    if num_descargas == 1:
                        with cols_descarga[1]:
                            if st.button("✅ Finalizar Proceso", type="primary", use_container_width=True):
                                st.session_state.paso_actual = 3
                                st.rerun()

                    # --- Si hay varios botones, poner el botón finalizar en la sección 3 ---
                    else:
                        st.divider()
                        col1, col2, col3 = st.columns([1, 1, 2])
                        with col1:
                            if st.button("✅ Finalizar Proceso", type="primary", use_container_width=True):
                                st.session_state.paso_actual = 3
                                st.rerun()
                
                else:
                    st.warning("⚠️ Detección manual necesaria")
            
            except Exception as e:
                st.error(f"❌ Error: {e}")

# ================================================
# PASO 3: FINALIZACIÓN
# ================================================
elif st.session_state.paso_actual == 3:
    st.balloons()
    
    st.markdown("""
    <div style='background-color: #78808C; padding: 30px; border-radius: 15px; text-align: center;'>
        <h1>🎉 ¡Proceso Completado!</h1>
        <p style='font-size: 18px;'>Todos los archivos han sido procesados y homologados correctamente.</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📋 Resumen del Proceso")
        st.success(f"**Colegio:** {st.session_state.nombre_colegio}")
        st.success(f"**Archivo 1:** {len(st.session_state.archivo1_df)} registros")
        
        if st.session_state.archivo2_1p3p_df is not None:
            st.success(f"**Hoja 1P-3P:** {len(st.session_state.archivo2_1p3p_df)} registros")
        if st.session_state.archivo2_4p5s_df is not None:
            st.success(f"**Hoja 4P-5S:** {len(st.session_state.archivo2_4p5s_df)} registros")
    
    with col2:
        st.markdown("### 🔄 Acciones")
        if st.button("🆕 Procesar Nuevo Colegio", type="primary", use_container_width=True):
            # Reiniciar todo
            st.session_state.paso_actual = 0
            st.session_state.nombre_colegio = ""
            st.session_state.archivo1_df = None
            st.session_state.archivo2_df = None
            st.rerun()
        
        if st.button("🔙 Volver al Paso 3", use_container_width=True):
            st.session_state.paso_actual = 2
            st.rerun()