
from __future__ import annotations
import io
import re
import unicodedata
from datetime import datetime
from typing import List, Tuple, Optional

import pandas as pd
import streamlit as st

# ===========================
# Configuración general
# ===========================
st.set_page_config(
    page_title="Validador & Homologador",
    page_icon="🗂️",
    layout="wide"
)

# ===========================
# Utilidades genéricas
# ===========================
def normalize_text(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s.upper()

def to_ddmmyyyy(value) -> Tuple[str, Optional[str]]:
    if pd.isna(value) or str(value).strip() == "":
        return "", "Fecha vacía"
    candidates = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%y", "%d-%m-%y"]
    try:
        if isinstance(value, (int, float)) and not pd.isna(value):
            base = datetime(1899, 12, 30)
            dt = base + pd.to_timedelta(int(value), unit="D")
            return dt.strftime("%d/%m/%Y"), None
    except Exception:
        pass
    s = str(value).strip()
    for fmt in candidates:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%d/%m/%Y"), None
        except Exception:
            continue
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="raise")
        return dt.strftime("%d/%m/%Y"), None
    except Exception:
        return s, f"Formato de fecha no reconocido: {s}"

def is_valid_email(s: str) -> bool:
    if not s:
        return False
    return bool(re.match(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", s))

def is_valid_dni(s: str) -> bool:
    return bool(re.fullmatch(r"\d{8}", s or ""))

def infer_sex_from_name(name: str) -> Optional[str]:
    if not name:
        return None
    name = normalize_text(name)
    if name.endswith("A"):
        return "F"
    if name.endswith("O"):
        return "M"
    return None

def full_name(nombre: str, paterno: str, materno: str) -> str:
    return normalize_text(f"{nombre} {paterno} {materno}").strip()

def clean_section(s: str) -> str:
    s = normalize_text(s)
    s = s.replace(" ", "")
    return s

def clean_grade(s: str) -> str:
    s = normalize_text(s)
    return s.replace(" ", "")

# ===========================
# Equivalencias y grados válidos
# ===========================
INTERNAL_COURSES_TXT = """ADOBE ILLUSTRATOR
ADOBE INDESIGN
ADOBE PHOTOSHOP PROFICIENT
COACHING PERSONAL Y VOCACIONAL
DE LA IDEA AL EMPRENDIMIENTO
DESARROLLO DE APLICACIONES MÓVILES
DESARROLLO WEB
DISEÑO WEB
EDICIÓN DE AUDIO
EDICIÓN DE VIDEO
EXCEL EXPERT SPECIALIST
EXCEL INTERMEDIATE SPECIALIST
EXCEL PROFICIENT SPECIALIST
FINANZAS PERSONALES
GESTIÓN DE DATA CON MS EXCEL Y POWER BI
GESTIÓN EMPRESARIAL
HABILIDADES BLANDAS
MARKETING DIGITAL
MARKETING PERSONAL
PRESENTACIONES DE IMPACTO
PROGRAMACIÓN VISUAL KODU PLANET I
PROGRAMACIÓN VISUAL KODU PLANET II
PROGRAMACIÓN VISUAL KODU PLANET III
ROBLOX FOR TEENS
ROBÓTICA
SCRATCH
WORD EXPERT SPECIALIST
WORD PROFICIENT SPECIALIST
LEARNING FOR BEGINNERS 1
LEARNING FOR BEGINNERS 2
CODE FOR KIDS
"""

VALID_GRADES = ["1P", "2P", "3P", "4P", "5P", "1S", "2S", "3S", "4S", "5S"]

@st.cache_data
def load_internal_courses() -> pd.DataFrame:
    lines = [normalize_text(x) for x in INTERNAL_COURSES_TXT.strip().splitlines() if x.strip()]
    return pd.DataFrame({"curso_oficial": sorted(set(lines))})

def read_courses_from_txt(file_bytes: bytes) -> pd.DataFrame:
    text = file_bytes.decode("utf-8", errors="ignore")
    rows = [normalize_text(line.strip()) for line in text.splitlines() if line.strip()]
    return pd.DataFrame({"curso_oficial": sorted(set(rows))})

def merge_new_equivalences(base: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    merged = pd.concat([base, new_df], ignore_index=True).drop_duplicates(subset=["curso_oficial"], keep="first")
    return merged.sort_values("curso_oficial").reset_index(drop=True)

# ===========================
# Validadores
# ===========================
def validate_nomina(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    errors_blocking: List[str] = []
    expected = ["Nombres", "Paterno", "Materno", "Fecha", "Sexo", "Grado", "Sección", "Correo", "DNI"]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas en Nómina: {missing}")
    out = df.copy()
    out["Nombre"] = out["Nombres"].map(normalize_text)
    out["Paterno"] = out["Paterno"].map(normalize_text)
    out["Materno"] = out["Materno"].map(normalize_text)
    out["alumno_clave"] = out.apply(lambda r: full_name(r["Nombre"], r["Paterno"], r["Materno"]), axis=1)
    report = pd.DataFrame({"fila": out.index, "alumno_clave": out["alumno_clave"]})
    return out, report, errors_blocking

def validate_notas(df: pd.DataFrame, courses_catalog: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, List[str], List[str]]:
    warnings: List[str] = []
    errors_blocking: List[str] = []
    essential = ["Nombres", "Paterno", "Materno", "Curso", "Nota Vigesimal"]
    missing_essential = [c for c in essential if c not in df.columns]
    if missing_essential:
        raise ValueError(f"Faltan columnas esenciales en Notas: {missing_essential}")
    df["Nombre"] = df["Nombres"].map(normalize_text)
    df["Paterno"] = df["Paterno"].map(normalize_text)
    df["Materno"] = df["Materno"].map(normalize_text)
    df["Curso"] = df["Curso"].map(normalize_text)
    df["Nota Vigesimal"] = df["Nota Vigesimal"].map(lambda v: v if str(v).strip() != "" else "NP")
    df["alumno_clave"] = df.apply(lambda r: full_name(r["Nombre"], r["Paterno"], r["Materno"]), axis=1)
    catalog = set(courses_catalog["curso_oficial"].tolist())
    df["curso_homologado"] = df["Curso"].apply(lambda c: c if c in catalog else None)
    report = df[["alumno_clave", "Curso", "curso_homologado", "Nota Vigesimal"]].copy()
    return df, report, errors_blocking, warnings

# ===========================
# Interfaz principal
# ===========================
def main():
    st.title("Validador y Homologador de Datos")
    if "courses_catalog" not in st.session_state:
        st.session_state.courses_catalog = load_internal_courses()

    st.sidebar.header("Carga de Archivos")
    file_nomina = st.sidebar.file_uploader("Archivo 1 - Nómina (Excel)", type=["xlsx", "xls"])
    file_notas = st.sidebar.file_uploader("Archivo 2 - Notas (Excel)", type=["xlsx", "xls"])
    file_eq = st.sidebar.file_uploader("equivalencias.txt (opcional)", type=["txt"])

    if file_eq is not None:
        try:
            new_eq = read_courses_from_txt(file_eq.getvalue())
            st.session_state.courses_catalog = merge_new_equivalences(st.session_state.courses_catalog, new_eq)
            st.success(f"Equivalencias actualizadas. Total cursos: {len(st.session_state.courses_catalog)}")
        except Exception as e:
            st.warning(f"No se pudo leer el archivo de equivalencias: {e}")

    st.divider()
    st.subheader("🧭 Configuración de cabeceras")
    col_a, col_b = st.columns(2)
    with col_a:
        header_nomina_row = st.number_input("Fila de cabecera para Nómina (0 = primera fila)", min_value=0, max_value=30, value=0, step=1)
    with col_b:
        header_notas_row = st.number_input("Fila de cabecera para Notas (0 = primera fila)", min_value=0, max_value=30, value=0, step=1)

    st.divider()
    st.subheader("1) Validación individual")
    cols = st.columns(2)

    valid_nomina, valid_notas = None, None

    with cols[0]:
        st.markdown("#### Validar Nómina")
        if file_nomina is not None:
            try:
                file_nomina.seek(0)
                df_nom = pd.read_excel(file_nomina, header=int(header_nomina_row))
                st.caption(f"Cabecera leída desde fila {header_nomina_row}")
                valid_nomina, report_nom, errors_nom = validate_nomina(df_nom)
                st.success("Nómina validada correctamente.")
                st.dataframe(valid_nomina, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"Error al validar Nómina: {e}")
        else:
            st.info("Sube el Archivo 1 (Nómina) para validar.")

    with cols[1]:
        st.markdown("#### Validar Notas")
        if file_notas is not None:
            try:
                file_notas.seek(0)
                df_not = pd.read_excel(file_notas, header=int(header_notas_row))
                st.caption(f"Cabecera leída desde fila {header_notas_row}")
                valid_notas, report_not, errors_not, warns_not = validate_notas(df_not, st.session_state.courses_catalog)
                st.success("Notas validadas correctamente.")
                st.dataframe(valid_notas, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"Error al validar Notas: {e}")
        else:
            st.info("Sube el Archivo 2 (Notas) para validar.")

    # Sección 2: Homologación y comparación/exportación
    st.divider()
    st.subheader("2) Homologación de cursos no reconocidos")
    if valid_notas is not None:
        catalog = st.session_state.courses_catalog
        unknown_courses = sorted(set(valid_notas.loc[valid_notas["curso_homologado"].isna(), "Curso"]))
        if len(unknown_courses) > 0:
            st.warning(f"Se detectaron {len(unknown_courses)} cursos no homologados.")
        else:
            st.success("No hay cursos pendientes por homologar.")

    st.divider()
    st.subheader("3) Comparación de alumnos entre Nómina y Notas")
    if valid_nomina is not None and valid_notas is not None:
        set_nom = set(valid_nomina["alumno_clave"])
        set_not = set(valid_notas["alumno_clave"])
        missing = sorted(list(set_nom - set_not))
        extra = sorted(list(set_not - set_nom))
        st.write("Alumnos faltantes en Notas:", missing)
        st.write("Alumnos extra en Notas:", extra)

    st.divider()
    st.subheader("4) Exportación de resultados")
    if valid_nomina is not None and valid_notas is not None:
        bytes_nom = io.BytesIO()
        with pd.ExcelWriter(bytes_nom, engine="xlsxwriter") as writer:
            valid_nomina.to_excel(writer, index=False, sheet_name="Nomina")
        bytes_nom.seek(0)
        bytes_not = io.BytesIO()
        with pd.ExcelWriter(bytes_not, engine="xlsxwriter") as writer:
            valid_notas.to_excel(writer, index=False, sheet_name="Notas")
        bytes_not.seek(0)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Descargar Nómina limpia", data=bytes_nom, file_name="Nomina_Validada.xlsx")
        with c2:
            st.download_button("📥 Descargar Notas limpias", data=bytes_not, file_name="Notas_Validadas.xlsx")

if __name__ == "__main__":
    main()