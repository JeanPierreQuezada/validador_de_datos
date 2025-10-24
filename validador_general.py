# ===============================================
# PASO 1: SUBIDA Y VALIDACIÓN DEL ARCHIVO 1
# ===============================================

import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata

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
    "Nº", "Paterno", "Materno", "Nombres", "Nacimiento (DD/MM/YYYY)", "Sexo (M/F)",
    "Grado", "Sección", "Correo institucional", "Neurodiversidad (Sí/No)", "DNI"
]

COLUMNAS_ARCHIVO2 = [
    "Nro.", "Paterno", "Materno", "Nombres", "Curso", "Grado", "Sección", "Nota Vigesimal"
]

# Constantes de validación
SEXO_VALIDO = ["M", "F"]
SECCIONES_VALIDAS = ["A", "B", "C", "D", "E", "F", "G", "U", "UNICO", "UNICA"]
GRADOS_VALIDOS = ["1P", "2P", "3P", "4P", "5P", "6P", "1S", "2S", "3S", "4S", "5S"]
MAPEO_GRADOS = {
    "1": "1P", "2": "2P", "3": "3P", "4": "4P", "5": "5P", "6": "6P",
    "7": "1S", "8": "2S", "9": "3S", "10": "4S", "11": "5S"
}


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

def crear_identificador(df, col_paterno, col_materno, col_nombres):
    """Crea columna identificador"""
    return (
        df[col_nombres].astype(str).str.strip() + " " +
        df[col_paterno].astype(str).str.strip() + " " +
        df[col_materno].astype(str).str.strip()
    )

def normalizar_enie(texto):
    """
    Normaliza texto a mayúsculas y elimina acentos, preservando Ñ
    """
    if pd.isna(texto):
        return ""
    
    texto = str(texto).strip().upper()
    
    # Normalizar caracteres Unicode
    texto_nfd = unicodedata.normalize('NFD', texto)
    
    # Reconstruir el texto eliminando solo las marcas diacríticas (acentos) y con Ñ
    resultado = []
    for char in texto_nfd:
        # Si es la Ñ, mantenerla
        if char in ['Ñ', 'ñ']:
            resultado.append('Ñ')
        # Si no es una marca diacrítica (acento), incluirla
        elif unicodedata.category(char) != 'Mn':
            resultado.append(char)
    
    return ''.join(resultado)

def normalizar_columnas_texto(df, columnas):
    """
    Normaliza texto a mayúsculas y elimina acentos, preservando Ñ
    """
    for col in columnas:
        df[col] = df[col].apply(normalizar_enie)
        # Normalizar espacios múltiples
        df[col] = df[col].str.replace(r'\s+', ' ', regex=True).str.strip()
    return df

def limpiar_filas_vacias(df, columnas_clave=None):
    """
    Args:
        df: DataFrame a limpiar
        columnas_clave: Lista de nombres de columnas para verificar (default: primeras 4)
    
    Returns:
        DataFrame limpio sin filas completamente vacías (evita sólo los Nro o N°)
    """
    if columnas_clave is None:
        # Usar las columnas 2, 3, 4
        columnas_clave = df.columns[1:4].tolist()
    
    # Contar registros originales
    total_original = len(df)
    
    # Filtrar: mantener solo filas donde AL MENOS UNA columna clave tenga datos
    df_limpio = df.dropna(subset=columnas_clave, how='all').copy()
    
    # Eliminar filas donde todas las columnas clave sean strings vacíos
    mask = df_limpio[columnas_clave].apply(
        lambda x: x.astype(str).str.strip().ne('')
    ).any(axis=1)
    df_limpio = df_limpio[mask].reset_index(drop=True)
    
    # Mostrar info si se eliminaron filas
    filas_eliminadas = total_original - len(df_limpio)
    if filas_eliminadas > 0:
        st.info(f"🧹 Se eliminaron {filas_eliminadas} filas vacías (quedaron {len(df_limpio)} registros)")
    
    return df_limpio

def homologar_dataframe(df):
    """
    Homologa un DataFrame completo:
    - Todas las columnas: Convierte a mayúsculas y quita espacios
    - Columnas PATERNO, MATERNO, NOMBRES: Además quita acentos y mantiene la Ñ
    """
    # Columnas especiales que requieren normalización de acentos
    columnas_nombres = ["PATERNO", "MATERNO", "NOMBRES"]
    filas_vacias = df[df[columnas_nombres].isnull().any(axis=1)]

    if not filas_vacias.empty:
        st.error("❌ Se detectaron campos vacíos en nombres o apellidos (Archivo 1 - Nómina)")
        st.dataframe(filas_vacias, use_container_width=True)
        #st.stop()

    # Procesar todas las columnas
    for col in df.columns:
        if col.upper() in columnas_nombres:
            # Para columnas de nombres: usar función que preserva Ñ
            df[col] = df[col].apply(normalizar_enie)
            # Normalizar espacios múltiples
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True).str.strip()
        else:
            # Solo mayúsculas y quitar espacios
            df[col] = (
                df[col].astype(str)
                .str.strip()
                .str.upper()
                .str.replace(r'\s+', ' ', regex=True)
                .str.strip()
            )
    
    return df

def convertir_numericas_a_entero(df, columnas=None):
    """
    Convierte valores numéricos flotantes a enteros (1.0 → "1")
    Funciona incluso en columnas mixtas (1.0, "2P", 3.0)
    
    Args:
        df: DataFrame a procesar
        columnas: Lista de columnas a convertir
    
    Returns:
        DataFrame con columnas convertidas
    """
    if columnas is None:
        columnas = df.select_dtypes(include=['float64', 'float32']).columns.tolist()
    
    for col in columnas:
        if col not in df.columns:
            continue
        
        # Procesar cada valor individualmente
        def convertir_valor(val):
            """Convierte un valor individual"""
            if pd.isna(val):
                return val
            
            # Convertir a string para inspeccionar
            val_str = str(val).strip()
            
            # Si ya tiene letras, dejarlo como está
            if any(c.isalpha() for c in val_str):
                return val_str
            
            # Si es numérico puro, intentar convertir
            try:
                val_num = float(val)
                # Si es un entero disfrazado de float (1.0, 2.0)
                if val_num % 1 == 0:
                    return str(int(val_num))
                else:
                    # Si tiene decimales reales (1.5), mantener como string
                    return val_str
            except (ValueError, TypeError):
                return val_str
        
        # Aplicar la conversión a toda la columna
        df[col] = df[col].apply(convertir_valor)
    
    return df

def validar_y_mapear_grados(df, col_grado="GRADO"):
    """
    Valida y mapea los grados. Convierte números 1-11 a formato estándar (1P-6P, 1S-5S).
    Retorna DataFrame procesado y lista de errores.
    """
    errores = []
    df[col_grado] = df[col_grado].astype(str).str.strip().str.upper()
    
    # Mapear números a grados
    df[col_grado] = df[col_grado].replace(MAPEO_GRADOS)
    
    # Validar grados
    grados_invalidos = df[~df[col_grado].isin(GRADOS_VALIDOS)]
    if len(grados_invalidos) > 0:
        for idx, row in grados_invalidos.iterrows():
            errores.append(f"Fila {idx + 2}: Grado inválido '{row[col_grado]}'")
    
    return df, errores

def validar_sexo(df, col_sexo="SEXO (M/F)"):
    """
    Valida que el sexo sea M o F.
    Retorna lista de errores.
    """
    errores = []
    df[col_sexo] = df[col_sexo].astype(str).str.strip().str.upper()
    
    sexos_invalidos = df[~df[col_sexo].isin(SEXO_VALIDO)]
    if len(sexos_invalidos) > 0:
        for idx, row in sexos_invalidos.iterrows():
            errores.append(f"Fila {idx + 2}: Sexo inválido '{row[col_sexo]}' (debe ser M o F)")
    
    return errores

def validar_secciones(df, col_seccion="SECCIÓN"):
    """
    Valida que las secciones sean válidas (A-G, U).
    Retorna lista de errores.
    """
    errores = []
    df[col_seccion] = df[col_seccion].astype(str).str.strip().str.upper()
    
    secciones_invalidas = df[~df[col_seccion].isin(SECCIONES_VALIDAS)]
    if len(secciones_invalidas) > 0:
        for idx, row in secciones_invalidas.iterrows():
            errores.append(f"Fila {idx + 2}: Sección inválida '{row[col_seccion]}' (debe ser A-G o U)")
    
    return errores

def validar_neurodiversidad(df, col_neuro="NEURODIVERSIDAD (SÍ/NO)"):
    """
    Valida que neurodiversidad sea Sí o No.
    Retorna lista de errores.
    """
    errores = []
    df[col_neuro] = df[col_neuro].astype(str).str.strip().str.upper()
    
    # Mapear variaciones comunes
    mapeo_neuro = {
        "SI": "SÍ", "S": "SÍ", "YES": "SÍ", "Y": "SÍ",
        "N": "NO", "NOT": "NO"
    }
    df[col_neuro] = df[col_neuro].replace(mapeo_neuro)
    
    neuros_invalidas = df[~df[col_neuro].isin(["SÍ", "NO"])]
    if len(neuros_invalidas) > 0:
        for idx, row in neuros_invalidas.iterrows():
            identificador = crear_identificador(df.loc[[idx]], "PATERNO", "MATERNO", "NOMBRES").iloc[0]
            errores.append(f"Fila {idx + 2}: Neurodiversidad inválida '{row[col_neuro]}' - {identificador}")
    
    return errores

def validar_fecha_nacimiento(df, col_fecha="NACIMIENTO (DD/MM/YYYY)"):
    """
    Valida formato de fecha DD/MM/YYYY.
    Retorna lista de errores.
    """
    errores = []
    
    for idx, row in df.iterrows():
        fecha = str(row[col_fecha]).strip()
        identificador = crear_identificador(df.loc[[idx]], "PATERNO", "MATERNO", "NOMBRES").iloc[0]
        
        # Validar formato DD/MM/YYYY
        if not pd.to_datetime(fecha, format="%d/%m/%Y", errors="coerce") is pd.NaT:
            continue
        else:
            errores.append(f"Fila {idx + 2}: Fecha inválida '{fecha}' - {identificador}")
    
    return errores

def validar_dni(df, col_dni="DNI"):
    """
    Valida que el DNI tenga exactamente 8 dígitos.
    Retorna lista de errores.
    """
    errores = []
    df[col_dni] = df[col_dni].astype(str).str.strip()
    
    for idx, row in df.iterrows():
        dni = row[col_dni]
        identificador = crear_identificador(df.loc[[idx]], "PATERNO", "MATERNO", "NOMBRES").iloc[0]
        
        # Validar que sea número de 8 dígitos
        if not (dni.isdigit() and len(dni) == 8):
            errores.append(f"Fila {idx + 2}: DNI inválido '{dni}' (debe ser 8 dígitos) - {identificador}")
    
    return errores

def validar_correo(df, col_correo="CORREO INSTITUCIONAL"):
    """
    Valida formato básico de correo electrónico.
    Retorna lista de errores.
    """
    errores = []
    
    for idx, row in df.iterrows():
        correo = str(row[col_correo]).strip().lower()
        identificador = crear_identificador(df.loc[[idx]], "PATERNO", "MATERNO", "NOMBRES").iloc[0]
        
        # Validación básica: contiene @ y .
        if "@" not in correo or "." not in correo.split("@")[-1]:
            errores.append(f"Fila {idx + 2}: Correo inválido '{correo}' - {identificador}")
    
    return errores

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
# PASO 0: Nombre DEL COLEGIO
# ================================================
if st.session_state.paso_actual == 0:
    st.header("🏫 Paso 1: Información del Colegio")

    st.markdown("""
        <div style='background-color: #78808C; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
            <h4>Bienvenido al sistema de validación</h4>
            <p>Para comenzar, ingresa el Nombre del colegio. Este Nombre se usará para identificar los archivos descargables.</p>
        </div>
        """, unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    
    with col1:
        NOMBRES = st.text_input(
            "Nombre del Colegio",
            value=st.session_state.nombre_colegio,
            placeholder="Ejemplo: Colegio San Martín de Porres",
            help="Este Nombre aparecerá en los archivos descargados"
        )
        
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➡️ Continuar", type="primary", use_container_width=True):
            if NOMBRES.strip():
                st.session_state.nombre_colegio = NOMBRES.strip()
                st.session_state.paso_actual = 1
                st.rerun()
            else:
                st.error("Por favor, ingresa el Nombre del colegio")

# ================================================
# PASO 1: ARCHIVO 1 (NÓMINA)
# ================================================
elif st.session_state.paso_actual == 1:
    # Mostrar resumen del paso anterior
    with st.expander("✅ Paso 1 completado: Nombre del Colegio", expanded=False):
        st.info(f"**Colegio:** {st.session_state.nombre_colegio}")
        if st.button("🔄 Cambiar Nombre", key="cambiar_nombre"):
            st.session_state.paso_actual = 0
            st.rerun()
    
    st.header("📋 Paso 2: Archivo de Nómina de Alumnos")
    
    with st.container():
        st.markdown("""
        <div style='background-color: #78808C; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
            <h4>📄 Instrucciones</h4>
            <p>Sube el archivo Excel que contiene la nómina de alumnos.</p>
            <p><strong>Columnas requeridas:</strong></p>
            <code>Nº, Paterno, Materno, Nombres, Nacimiento (DD/MM/YYYY), Sexo (M/F), Grado, Sección, Correo institucional, Neurodiversidad (Sí/No), DNI</code>
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

                        # Eliminar filas con campos vacíos en PATERNO, MATERNO y NOMBRES
                        df = limpiar_filas_vacias(df, columnas_clave=["PATERNO", "MATERNO", "NOMBRES"])
                        
                        df = convertir_numericas_a_entero(df, columnas=["GRADO"])

                        # Convertir numéricas a enteros
                        df = homologar_dataframe(df)

                        # Validar campos vacíos en PATERNO, MATERNO o NOMBRES
                        columnas_obligatorias = ["PATERNO", "MATERNO", "NOMBRES"]
                        filas_vacias = df[df[columnas_obligatorias].isnull().any(axis=1)]

                        if not filas_vacias.empty:
                            st.error("❌ Se detectaron campos vacíos en nombres o apellidos (Archivo 1 - Nómina)")
                            st.dataframe(filas_vacias, use_container_width=True)
                            #st.stop()
                        
                        # Validaciones para Archivo 1 (nómina)
                        errores_fatales = []
                        alertas = []
                        
                        # 1. Validar y mapear grados
                        df, errores_grados = validar_y_mapear_grados(df, "GRADO")
                        errores_fatales.extend(errores_grados)
                        
                        # 2. Validar sexo
                        errores_sexo = validar_sexo(df, "SEXO (M/F)")
                        errores_fatales.extend(errores_sexo)
                        
                        # 3. Validar secciones
                        errores_secciones = validar_secciones(df, "SECCIÓN")
                        errores_fatales.extend(errores_secciones)

                        # 4. Validar neurodiversidad
                        errores_neuro = validar_neurodiversidad(df, "NEURODIVERSIDAD (SÍ/NO)")
                        alertas.extend(errores_neuro)
                        
                        # 5. Validar fecha
                        errores_fecha = validar_fecha_nacimiento(df, "NACIMIENTO (DD/MM/YYYY)")
                        alertas.extend(errores_fecha)
                        
                        # 6. Validar DNI
                        errores_dni = validar_dni(df, "DNI")
                        alertas.extend(errores_dni)
                        
                        # 7. Validar correo
                        errores_correo = validar_correo(df, "CORREO INSTITUCIONAL")
                        alertas.extend(errores_correo)
                        
                        # Mostrar errores si existen
                        if errores_fatales:
                            st.error("❌ Se encontraron errores de validación:")
                            # Convertir lista de alertas a DataFrame
                            df_errores_fatales = pd.DataFrame(alertas, columns=["Detalle de la Alerta"])
                                
                            # Mostrar tabla scrolleable
                            st.dataframe(
                                df_errores_fatales,
                                use_container_width=True,
                                height=220  # ajusta la altura visible (unas 5-6 filas aprox)
                            )
                                
                            st.caption(f"🔎 Total de errores: {len(errores_fatales)}")
                            st.info("Por favor, corrige estos errores en el archivo y vuelve a cargarlo")
                            #st.stop()

                        else:
                            df["IDENTIFICADOR"] = crear_identificador(df, "PATERNO", "MATERNO", "NOMBRES")
                            st.session_state.archivo1_df = df
                            
                            if alertas:
                                st.warning("⚠️ Se detectaron advertencias en los datos (no bloquean el proceso):")
                                with st.expander("Ver alertas detalladas", expanded=True):
                                    # Convertir lista de alertas a DataFrame
                                    df_alertas = pd.DataFrame(alertas, columns=["Detalle de la Alerta"])
                                    
                                    # Mostrar tabla scrolleable
                                    st.dataframe(
                                        df_alertas,
                                        use_container_width=True,
                                        height=220  # ajusta la altura visible (unas 5-6 filas aprox)
                                    )
                                    
                                    st.caption(f"🔎 Total de alertas: {len(alertas)}")
                            else:
                                st.success("✅ Todas las validaciones pasaron correctamente")

                        if not errores_fatales:
                            # Mostrar preview
                            st.markdown("### 📊 Vista Previa de Datos")
                            st.info(f"Total de registros: {len(df)}")
                            st.dataframe(df.head(10), use_container_width=True)
                        
                        # Botones de acción
                        col1, col2 = st.columns(2)
                        with col1:
                            df_descarga = df.drop(columns=["IDENTIFICADOR", "Nº"], errors="ignore")
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
                                #st.rerun()
                    
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
                                
                                # Homologar datos
                                df = homologar_dataframe(df)
                                
                                # Validar campos vacíos en PATERNO, MATERNO o NOMBRES
                                columnas_obligatorias = ["PATERNO", "MATERNO", "NOMBRES"]
                                filas_vacias = df[df[columnas_obligatorias].isnull().any(axis=1)]

                                if not filas_vacias.empty:
                                    st.error("❌ Se detectaron campos vacíos en nombres o apellidos (Archivo 1 - Nómina)")
                                    st.dataframe(filas_vacias, use_container_width=True)
                                    #st.stop()
                                
                                # Validaciones para Archivo 1 (nómina)
                                errores_validacion = []
                                
                                # 1. Validar y mapear grados
                                df, errores_grados = validar_y_mapear_grados(df, "GRADO")
                                errores_validacion.extend(errores_grados)
                                
                                # 2. Validar sexo
                                errores_sexo = validar_sexo(df, "SEXO (M/F)")
                                errores_validacion.extend(errores_sexo)
                                
                                # 3. Validar secciones
                                errores_secciones = validar_secciones(df, "SECCIÓN")
                                errores_validacion.extend(errores_secciones)

                                # 4. Validar neurodiversidad
                                errores_neuro = validar_neurodiversidad(df, "NEURODIVERSIDAD (SÍ/NO)")
                                errores_validacion.extend(errores_neuro)
                                
                                # 5. Validar fecha
                                errores_fecha = validar_fecha_nacimiento(df, "NACIMIENTO (DD/MM/YYYY)")
                                errores_validacion.extend(errores_fecha)
                                
                                # 6. Validar DNI
                                errores_dni = validar_dni(df, "DNI")
                                errores_validacion.extend(errores_dni)
                                
                                # 7. Validar correo
                                errores_correo = validar_correo(df, "CORREO INSTITUCIONAL")
                                errores_validacion.extend(errores_correo)
                                
                                # Mostrar errores o continuar
                                if errores_validacion:
                                    st.error("❌ Se encontraron errores de validación:")
                                    with st.expander("Ver errores detallados", expanded=True):
                                        for error in errores_validacion[:50]:
                                            st.warning(error)
                                        if len(errores_validacion) > 50:
                                            st.info(f"... y {len(errores_validacion) - 50} errores más")
                                else:
                                    df["IDENTIFICADOR"] = crear_identificador(df, "PATERNO", "MATERNO", "NOMBRES")
                                    st.session_state.archivo1_df = df
                                    st.success("✅ Validaciones pasadas correctamente")
                                    #st.rerun()
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
        <code>Nro., Paterno, Materno, Nombres, Curso, Grado, Sección,  Nota Vigesimal</code>
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
                    #st.stop()
                
                # Mostrar información de hojas detectadas
                st.success(f"✅ Hojas detectadas en el archivo, Únicas Opciones ('1P-3P' o '4P-5S'):")
                cols_info = st.columns(2)
                with cols_info[0]:
                    if tiene_1p3p:
                        st.info("📘 **1P-3P** encontrada")
                with cols_info[1]:
                    if tiene_4p5s:
                        st.info("📗 **4P-5S** encontrada")
                
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
                        
                        # Eliminar filas con campos vacíos en PATERNO, MATERNO y NOMBRES
                        df_1p3p = limpiar_filas_vacias(df_1p3p, columnas_clave=["PATERNO", "MATERNO", "NOMBRES"])

                        # Convertir numéricas a enteros
                        df_1p3p = convertir_numericas_a_entero(df_1p3p, columnas=["GRADO", "NOTA VIGESIMAL"])

                        # Homologar datos
                        df_1p3p = homologar_dataframe(df_1p3p)

                        # 🔍 Validar campos vacíos en PATERNO, MATERNO o NOMBRES
                        columnas_obligatorias = ["PATERNO", "MATERNO", "NOMBRES"]
                        filas_vacias = df_1p3p[df_1p3p[columnas_obligatorias].isnull().any(axis=1)]

                        if not filas_vacias.empty:
                            st.error("❌ Se detectaron campos vacíos en nombres o apellidos (Hoja 1P-3P)")
                            st.dataframe(filas_vacias, use_container_width=True)
                            #st.stop()
                        
                        # Validaciones para Archivo 2 - Hoja 1P-3P
                        errores_validacion_1p3p = []

                        # Completar valores vacíos en NOTA VIGESIMAL con "NP"
                        if "NOTA VIGESIMAL" in df_1p3p.columns:
                            df_1p3p["NOTA VIGESIMAL"] = df_1p3p["NOTA VIGESIMAL"].fillna("NP").replace("", "NP")

                        # 1. Validar y mapear grados
                        df_1p3p, errores_grados = validar_y_mapear_grados(df_1p3p, "GRADO")
                        errores_validacion_1p3p.extend(errores_grados)
                        
                        # 2. Validar secciones
                        errores_secciones = validar_secciones(df_1p3p, "SECCIÓN")
                        errores_validacion_1p3p.extend(errores_secciones)
                        
                        # Mostrar errores de validación si existen
                        if errores_validacion_1p3p:
                            st.error("❌ Errores de validación en 1P-3P:")
                            with st.expander("Ver errores detallados", expanded=True):
                                for error in errores_validacion_1p3p[:30]:
                                    st.warning(error)
                                if len(errores_validacion_1p3p) > 30:
                                    st.info(f"... y {len(errores_validacion_1p3p) - 30} errores más")
                            st.info("Por favor, corrige estos errores y vuelve a cargar el archivo")
                        else:
                            st.success("✅ Validaciones de grados y secciones pasadas (1P-3P)")
                        
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
                                        
                                        # Agregar solo IDENTIFICADOR
                                        df_1p3p["IDENTIFICADOR"] = crear_identificador(df_1p3p, "PATERNO", "MATERNO", "NOMBRES")
                                        
                                        # Reordenar
                                        cols_orden = [c for c in df_1p3p.columns if c != "IDENTIFICADOR"]
                                        cols_orden.append("IDENTIFICADOR")
                                        df_1p3p = df_1p3p[cols_orden]
                                        
                                        # Guardar en session_state
                                        st.session_state.archivo2_1p3p_df = df_1p3p
                                        
                                        st.success("✅ Cursos homologados correctamente en 1P-3P")
                                        #st.rerun()
                        
                        # Si no hay cursos inválidos
                        if len(cursos_invalidos_1p3p) == 0 or st.session_state.archivo2_1p3p_df is not None:
                            # Usar el DataFrame guardado si existe, sino usar el actual
                            if st.session_state.archivo2_1p3p_df is not None:
                                df_1p3p = st.session_state.archivo2_1p3p_df
                            else:
                                # Agregar solo IDENTIFICADOR
                                df_1p3p["IDENTIFICADOR"] = crear_identificador(df_1p3p, "PATERNO", "MATERNO", "NOMBRES")
                                
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

                    # Eliminar filas con campos vacíos en PATERNO, MATERNO y NOMBRES
                    df2 = limpiar_filas_vacias(df2, columnas_clave=["PATERNO", "MATERNO", "NOMBRES"])
                    
                    # Convertir numéricas a enteros
                    df2 = convertir_numericas_a_entero(df2, columnas=["GRADO", "NOTA VIGESIMAL"])

                    # Homologar datos
                    df2 = homologar_dataframe(df2)

                    # Completar valores vacíos en NOTA VIGESIMAL con "NP"
                    if "NOTA VIGESIMAL" in df2.columns:
                        df2["NOTA VIGESIMAL"] = df2["NOTA VIGESIMAL"].fillna("NP").replace("", "NP")
                    
                    # Validar campos vacíos
                    columnas_oblig = ["PATERNO", "MATERNO", "NOMBRES", "CURSO", "GRADO", "SECCIÓN", "NOTA VIGESIMAL"]
                    filas_vacias = df2[df2[columnas_oblig].isnull().any(axis=1)]
                    
                    if not filas_vacias.empty:
                        st.error("❌ Se detectaron campos vacíos")
                        st.dataframe(filas_vacias, use_container_width=True)
                        #st.stop()
                    
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
                                    
                                    # Guardar en session_state
                                    df2["IDENTIFICADOR"] = crear_identificador(df2, "PATERNO", "MATERNO", "NOMBRES")
                                    df2["NOTAS VIGESIMALES 75%"] = ""
                                    df2["PROMEDIO"] = ""
                                    
                                    # Reordenar columnas
                                    cols_orden = [c for c in df2.columns if c != "IDENTIFICADOR"]
                                    cols_orden.append("IDENTIFICADOR")
                                    df2 = df2[cols_orden]
                                    
                                    # Guardar en session_state
                                    st.session_state.archivo2_4p5s_df = df2
                                    
                                    st.success("✅ Cursos homologados correctamente")
                                    #st.rerun()
                    
                    # Si no hay cursos inválidos
                    if len(cursos_invalidos) == 0 or st.session_state.archivo2_4p5s_df is not None:
                        df2["IDENTIFICADOR"] = crear_identificador(df2, "PATERNO", "MATERNO", "NOMBRES")
                        df2["NOTAS VIGESIMALES 75%"] = ""
                        df2["PROMEDIO"] = ""
                        
                        # Reordenar columnas
                        cols_orden = [c for c in df2.columns if c != "IDENTIFICADOR"]
                        cols_orden.append("IDENTIFICADOR")
                        df2 = df2[cols_orden]
                        
                        st.session_state.archivo2_4p5s_df = df2
                        
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
                    
                    # --- Si solo hay un botón ---
                    if num_descargas == 1:
                        with cols_descarga[1]:
                            if st.button("✅ Finalizar Proceso", type="primary", use_container_width=True):
                                st.session_state.paso_actual = 3
                                st.rerun()

                    # --- Si hay varios botones ---
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