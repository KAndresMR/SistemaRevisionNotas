"""
config.py
=========

Configuración central del sistema de auditoría de calificaciones.

Este es el ÚNICO lugar donde deben definirse:
    - qué materias participan en la comparación (RELEVANT_SUBJECTS)
    - la tolerancia permitida entre nota de docentes y nota del sistema
    - nombres de columnas "especiales" (nómina, promedio) usados por el lector

Si en el futuro cambia una materia, se agrega/edita/elimina una sola vez
aquí y el resto del sistema no necesita modificarse.
"""

from typing import NamedTuple


class SubjectMapping(NamedTuple):
    """
    Relaciona una materia con el nombre de su columna en cada archivo.

    teacher_column: nombre exacto de la columna en el Excel de docentes.
    official_column: nombre exacto de la columna "Nota Final" dentro del
                      bloque de esa materia en el Excel oficial del sistema.
    """
    teacher_column: str
    official_column: str


# --------------------------------------------------------------------------
# MATERIAS RELEVANTES
# --------------------------------------------------------------------------
# Clave interna (usada en resultados y reportes) -> SubjectMapping.
#
# IMPORTANTE:
#   - El orden de este diccionario determina el orden en que las materias
#     se procesan y aparecen en el reporte.
#   - "promedio_general" es un caso especial: compara PROMEDIO FINAL
#     (docentes) contra P.GENERAL (sistema), no es una materia normal,
#     pero se maneja igual que las demás gracias a este mapeo uniforme.
#   - "ANIMACIÓN A LA LECTURA" existe en el Excel oficial pero NO está
#     aquí a propósito: por eso el sistema nunca la compara.
# --------------------------------------------------------------------------
RELEVANT_SUBJECTS: dict[str, SubjectMapping] = {
    "lengua_y_literatura": SubjectMapping("LENGUA Y LITERATURA", "LENGUA Y LITERATURA"),
    "matematica": SubjectMapping("MATEMÁTICA", "MATEMÁTICA"),
    "ciencias_naturales": SubjectMapping("CIENCIAS NATURALES", "CIENCIAS NATURALES"),
    "estudios_sociales": SubjectMapping("ESTUDIOS SOCIALES", "ESTUDIOS SOCIALES"),
    "ingles": SubjectMapping("INGLÉS", "INGLÉS"),
    "educacion_cultural_artistica": SubjectMapping(
        "EDUCACIÓN CULTURAL Y ARTÍSTICA", "EDUCACIÓN CULTURAL Y ARTÍSTICA"
    ),
    "educacion_fisica": SubjectMapping("EDUCACIÓN FÍSICA", "EDUCACIÓN FÍSICA"),
    "promedio_general": SubjectMapping("PROMEDIO FINAL", "P.GENERAL"),
}

# Alias posibles en el archivo de docentes para tolerar variantes de nombres de columnas
# entre grados superiores (8vo-10mo) y elementales (2do-7mo, ej: ECA vs EDUCACIÓN CULTURAL Y ARTÍSTICA,
# PROMEDIO vs PROMEDIO FINAL).
TEACHER_SUBJECT_ALIASES: dict[str, tuple[str, ...]] = {
    "lengua_y_literatura": ("LENGUA Y LITERATURA", "LENGUA"),
    "matematica": ("MATEMÁTICA", "MATEMATICA"),
    "ciencias_naturales": ("CIENCIAS NATURALES",),
    "estudios_sociales": ("ESTUDIOS SOCIALES",),
    "ingles": ("INGLÉS", "INGLES"),
    "educacion_cultural_artistica": (
        "EDUCACIÓN CULTURAL Y ARTÍSTICA",
        "EDUCACION CULTURAL Y ARTISTICA",
        "ECA",
    ),
    "educacion_fisica": ("EDUCACIÓN FÍSICA", "EDUCACION FISICA"),
    "promedio_general": ("PROMEDIO FINAL", "PROMEDIO"),
}

# Nombre "amigable" para mostrar en el reporte (clave interna -> texto legible).
# Si no se especifica una materia aquí, se usa la clave interna tal cual.
SUBJECT_DISPLAY_NAMES: dict[str, str] = {
    "lengua_y_literatura": "Lengua y Literatura",
    "matematica": "Matemática",
    "ciencias_naturales": "Ciencias Naturales",
    "estudios_sociales": "Estudios Sociales",
    "ingles": "Inglés",
    "educacion_cultural_artistica": "Educación Cultural y Artística",
    "educacion_fisica": "Educación Física",
    "promedio_general": "Promedio Final / P. General",
}


# --------------------------------------------------------------------------
# TOLERANCIA
# --------------------------------------------------------------------------
# diferencia = abs(nota_docente - nota_sistema)
# diferencia <= TOLERANCE  -> equivalente, NO se reporta
# diferencia >  TOLERANCE  -> discrepancia, SÍ se reporta
TOLERANCE: float = 0.02


# --------------------------------------------------------------------------
# COLUMNAS ESPECIALES (no son "materias", pero el lector necesita saber
# cómo se llaman para identificar al estudiante y las notas "Nota Final")
# --------------------------------------------------------------------------

# Posibles nombres de columna para la nómina/nombre del estudiante en el
# Excel de docentes. Se busca la primera que exista (case-insensitive).
TEACHER_STUDENT_NAME_COLUMNS: tuple[str, ...] = (
    "NOMINA",
    "NÓMINA",
    "ESTUDIANTE",
    "LISTADO",
    "APELLIDOS/NOMBRES",
    "APELLIDOS / NOMBRES",
    "NOMBRES Y APELLIDOS",
)

# Posibles nombres de columna para la nómina/nombre del estudiante en el
# Excel oficial del sistema.
OFFICIAL_STUDENT_NAME_COLUMNS: tuple[str, ...] = (
    "NOMINA",
    "NÓMINA",
    "ESTUDIANTE",
    "LISTADO",
    "APELLIDOS/NOMBRES",
    "APELLIDOS / NOMBRES",
    "NOMBRES Y APELLIDOS",
)

# En el Excel oficial, dentro del bloque de cada materia, este es el texto
# que identifica la sub-columna que sí nos interesa (las demás: I, II, III
# se ignoran).
OFFICIAL_FINAL_GRADE_LABEL: str = "NOTA FINAL"

# Palabras clave en el nombre que identifican filas de resumen/pie de página
# (NO son estudiantes) y deben ser descartadas automáticamente.
KEYWORDS_TO_IGNORE_IN_STUDENTS: tuple[str, ...] = (
    "PROMEDIO",
    "RECTOR",
    "VICERRECTOR",
    "DOCENTE",
    "FIRMA",
    "TOTAL",
)

# Orden deseado de las hojas en el archivo consolidado marcado (de 1ro a 10mo).
SHEET_ORDER: list[str] = [
    "primero A", "primero B",
    "segundo A", "segundo B",
    "tercero A", "tercero B",
    "cuarto A", "cuarto B",
    "quinto A", "quinto B",
    "sexto A", "sexto B",
    "séptimo A", "séptimo B",
    "octavo A", "octavo B",
    "noveno A", "noveno B",
    "décimo A", "décimo B",
]

# ---------------------------------------------------------------------------
# MAPEO DE CURSOS PARA LA APP WEB
# ---------------------------------------------------------------------------
MATUTINA_MAP: dict[str, tuple[str, str]] = {
    "1AM":  ("primero A", "1AM Matutino"),
    "1BM":  ("primero B", "1BM Matutino"),
    "2AM":  ("segundo A", "2AM Matutino"),
    "2BM":  ("segundo B", "2BM Matutino"),
    "3AM":  ("tercero A", "3AM Matutino"),
    "3BM":  ("tercero B", "3BM Matutino"),
    "4AM":  ("cuarto A",  "4AM Matutino"),
    "4BM":  ("cuarto B",  "4BM Matutino"),
    "5AM":  ("quinto A",  "5AM Matutino"),
    "5BM":  ("quinto B",  "5BM Matutino"),
    "6AM":  ("sexto A",   "6AM Matutino"),
    "6BM":  ("sexto B",   "6BM Matutino"),
    "7AM":  ("séptimo A", "7AM Matutino"),
    "7BM":  ("séptimo B", "7BM Matutino"),
    "8AM":  ("octavo A",  "8AM Matutino"),
    "8BM":  ("octavo B",  "8BM Matutino"),
    "9AM":  ("noveno A",  "9AM Matutino"),
    "9BM":  ("noveno B",  "9BM Matutino"),
    "10AM": ("décimo A",  "10AM Matutino"),
    "10BM": ("décimo B",  "10BM Matutino"),
}

VESPERTINA_MAP: dict[str, tuple[str, str]] = {
    "1AV":  ("primero A",  "1AV Vespertina"),
    "1BV":  ("primero B",  "1BV Vespertina"),
    "2AV":  ("segundo A",  "2AV Vespertina"),
    "2BV":  ("segundo B",  "2BV Vespertina"),
    "3AV":  ("tercero A",  "3AV Vespertina"),
    "3BV":  ("tercero B",  "3BV Vespertina"),
    "4AV":  ("cuarto A",   "4AV Vespertina"),
    "4BV":  ("cuarto B",   "4BV Vespertina"),
    "5AV":  ("quinto A",   "5AV Vespertina"),
    "5BV":  ("quinto B",   "5BV Vespertina"),
    "6AV":  ("sexto A",    "6AV Vespertina"),
    "6BV":  ("sexto B",    "6BV Vespertina"),
    "7AV":  ("séptimo A",  "7AV Vespertina"),
    "7BV":  ("séptimo B",  "7BV Vespertina"),
    "8AV":  ("octavo A",   "8AV Vespertina"),
    "8BV":  ("octavo B",   "8BV Vespertina"),
    "9AV":  ("noveno A",   "9AV Vespertina"),
    "9BV":  ("noveno B",   "9BV Vespertina"),
    "10AV": ("décimo A",   "10AV Vespertina"),
    "10BV": ("décimo B",   "10BV Vespertina"),
}

FULL_MAP = {**MATUTINA_MAP, **VESPERTINA_MAP}
