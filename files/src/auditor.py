

# ========================================
# models.py
# ========================================

"""
models.py
=========

Estructuras de datos usadas a través de todo el sistema.

Se usan `dataclasses` (no dicts sueltos) para que el resto del código
tenga autocompletado, tipado, y sea explícito sobre qué información
viaja entre módulos.
"""

from dataclasses import dataclass, field
from enum import Enum


class GradeStatus(str, Enum):
    """Estado de una comparación de nota individual (una materia)."""
    OK = "OK"
    ERROR = "ERROR"
    SIN_DATO = "SIN DATO"  # nota vacía o no convertible a número en alguno de los dos archivos


class StudentMatchStatus(str, Enum):
    """Estado del emparejamiento de un estudiante entre ambos archivos."""
    MATCHED_BY_POSITION = "MATCHED_BY_POSITION"      # misma posición, mismo nombre normalizado
    MATCHED_BY_NAME = "MATCHED_BY_NAME"              # desfase de posición, encontrado por nombre
    ONLY_IN_TEACHERS = "ONLY_IN_TEACHERS"            # existe solo en el Excel de docentes
    ONLY_IN_OFFICIAL = "ONLY_IN_OFFICIAL"            # existe solo en el Excel oficial


@dataclass
class Student:
    """
    Representa a un estudiante leído de UNO de los dos archivos.

    raw_name:        nombre tal como aparece en el Excel (para mostrar en reportes).
    normalized_name: nombre normalizado (mayúsculas, sin tildes, sin espacios
                      extra) usado únicamente para comparar/matchear.
    row_index:       posición (0-based) en la que apareció este estudiante
                      dentro de su archivo de origen. Es la base de la
                      "ruta rápida" de comparación por índice.
    grades:          diccionario {clave_interna_materia: valor_crudo_de_celda}.
                      Se guarda el valor crudo (puede ser str, float, None);
                      la conversión a float ocurre en el comparador, no aquí.
    """
    raw_name: str
    normalized_name: str
    row_index: int
    excel_row: int | None = None
    grades: dict[str, object] = field(default_factory=dict)


@dataclass
class MatchResult:
    """
    Resultado de emparejar un estudiante entre ambos archivos.

    teacher:  Student del archivo de docentes, o None si no existe ahí
              (caso ONLY_IN_OFFICIAL).
    official: Student del archivo oficial, o None si no existe ahí
              (caso ONLY_IN_TEACHERS).
    status:   cómo se llegó a este resultado (ver StudentMatchStatus).
    """
    teacher: "Student | None"
    official: "Student | None"
    status: StudentMatchStatus


@dataclass
class ComparisonResult:
    """Resultado de comparar UNA materia entre docentes y sistema para un estudiante."""
    subject_key: str
    subject_display: str
    teacher_grade: float | None
    official_grade: float | None
    difference: float | None
    status: GradeStatus


@dataclass
class StudentAuditRow:
    """
    Agrupa todo lo relevante de UN estudiante para el reporte final.

    student_name:      nombre a mostrar (se prioriza el de docentes si existe).
    match_status:       cómo fue emparejado (o si no se encontró en un archivo).
    comparisons:        lista de ComparisonResult, una por materia relevante.
                         Vacía si el estudiante no se encontró en ambos archivos.
    has_error:           True si al menos una comparación tiene status ERROR.
    has_desfase:         True si el estudiante fue encontrado por nombre en vez
                          de por posición directa (match_status == MATCHED_BY_NAME),
                          o si no se encontró en uno de los dos archivos.
    """
    student_name: str
    match_status: StudentMatchStatus
    excel_row: int | None = None
    comparisons: list[ComparisonResult] = field(default_factory=list)

    @property
    def has_error(self) -> bool:
        return any(c.status == GradeStatus.ERROR for c in self.comparisons)

    @property
    def has_desfase(self) -> bool:
        return self.match_status in (
            StudentMatchStatus.MATCHED_BY_NAME,
            StudentMatchStatus.ONLY_IN_TEACHERS,
            StudentMatchStatus.ONLY_IN_OFFICIAL,
        )


# ========================================
# name_normalizer.py
# ========================================

"""
name_normalizer.py
===================

Normalización de nombres de estudiantes.

Objetivo: que "JUAN PÉREZ", "Juan Pérez" y "juan   perez " se consideren
el mismo estudiante al hacer matching, sin afectar el nombre que se
muestra en el reporte final (para eso se conserva raw_name aparte).

Esta función se llama UNA SOLA VEZ por estudiante al momento de leer
cada archivo (ver excel_reader.py), nunca dentro de un bucle de
comparación repetida.
"""

import unicodedata
import re

def normalize_name(raw_name: str) -> str:
    """
    Normaliza un nombre (de estudiante o encabezado) para comparación.

    Pasos:
        1. Convierte a mayúsculas.
        2. Elimina tildes/diacríticos (á -> A, ñ -> N, etc.)
        3. Elimina TODOS los espacios, guiones, puntos y caracteres especiales.
        
    El resultado es una cadena puramente alfanumérica (ej. JUANPEREZ).
    No afecta al raw_name original que se muestra en el reporte.
    """
    if raw_name is None:
        return ""

    text = str(raw_name).upper()
    
    # Eliminar diacríticos
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    
    # Conservar solo letras (A-Z) y números (0-9)
    text = re.sub(r'[^A-Z0-9]', '', text)

    return text


# ========================================
# matcher.py
# ========================================

"""
matcher.py
==========

Empareja estudiantes entre el archivo de docentes y el archivo oficial.

Estrategia (en este orden):

    1. RUTA RÁPIDA — comparación por posición/índice (O(n)):
       teachers[i] vs official[i]. Si el nombre normalizado coincide,
       es un match directo y NO se hace ninguna búsqueda adicional.

    2. VALIDACIÓN DE SEGURIDAD — si los nombres en la misma posición NO
       coinciden, nunca se asume que son "la misma persona en distinto
       orden de columnas": se marca como posible desfase y se pasa a
       la ruta secundaria.

    3. RUTA SECUNDARIA (solo cuando hace falta) — se construye, la
       PRIMERA vez que se necesita (no antes), un diccionario
       {nombre_normalizado: Student} del archivo oficial. Esto cuesta
       O(n) una sola vez y cada búsqueda posterior es O(1), en vez de
       recorrer la lista completa por cada estudiante desfasado.

Al final, cualquier estudiante que:
    - solo existe en docentes           -> ONLY_IN_TEACHERS
    - solo existe en el sistema         -> ONLY_IN_OFFICIAL
    - fue encontrado por nombre (no por posición) -> MATCHED_BY_NAME
queda marcado explícitamente. Nunca se "arrastra" una comparación de
notas entre dos estudiantes que en realidad no son la misma persona.
"""

import logging


logger = logging.getLogger(__name__)


def match_students(teachers: list[Student], official: list[Student]) -> list[MatchResult]:
    results: list[MatchResult] = []
    matched_official_row_indices: set[int] = set()

    # Se construye SOLO si en algún momento hace falta (primer desfase).
    official_by_name: dict[str, Student] | None = None

    def get_official_by_name() -> dict[str, Student]:
        nonlocal official_by_name
        if official_by_name is None:
            official_by_name = {s.normalized_name: s for s in official}
            logger.info(
                "Desfase detectado: se construyó índice por nombre (%d estudiantes) "
                "para la ruta secundaria de matching.",
                len(official_by_name),
            )
        return official_by_name

    def try_match_by_name(teacher: Student) -> Student | None:
        candidate = get_official_by_name().get(teacher.normalized_name)
        if candidate is not None and candidate.row_index not in matched_official_row_indices:
            return candidate
        return None

    n = min(len(teachers), len(official))

    # --- Ruta rápida: comparación por posición para el rango común ---
    for i in range(n):
        teacher = teachers[i]
        official_candidate = official[i]

        if teacher.normalized_name == official_candidate.normalized_name:
            results.append(MatchResult(teacher, official_candidate, StudentMatchStatus.MATCHED_BY_POSITION))
            matched_official_row_indices.add(official_candidate.row_index)
            continue

        # Nombres distintos en la misma posición: NO se asume coincidencia.
        match = try_match_by_name(teacher)
        if match is not None:
            results.append(MatchResult(teacher, match, StudentMatchStatus.MATCHED_BY_NAME))
            matched_official_row_indices.add(match.row_index)
        else:
            results.append(MatchResult(teacher, None, StudentMatchStatus.ONLY_IN_TEACHERS))

    # --- Docentes con más filas que el archivo oficial (cola sobrante) ---
    for teacher in teachers[n:]:
        match = try_match_by_name(teacher)
        if match is not None:
            results.append(MatchResult(teacher, match, StudentMatchStatus.MATCHED_BY_NAME))
            matched_official_row_indices.add(match.row_index)
        else:
            results.append(MatchResult(teacher, None, StudentMatchStatus.ONLY_IN_TEACHERS))

    # --- Estudiantes del sistema que nunca fueron reclamados por nadie ---
    for official_student in official:
        if official_student.row_index not in matched_official_row_indices:
            results.append(MatchResult(None, official_student, StudentMatchStatus.ONLY_IN_OFFICIAL))

    n_desfase = sum(1 for r in results if r.status == StudentMatchStatus.MATCHED_BY_NAME)
    n_only_teachers = sum(1 for r in results if r.status == StudentMatchStatus.ONLY_IN_TEACHERS)
    n_only_official = sum(1 for r in results if r.status == StudentMatchStatus.ONLY_IN_OFFICIAL)
    logger.info(
        "Matching completo: %d por posición, %d por nombre (desfase), "
        "%d solo en docentes, %d solo en sistema.",
        sum(1 for r in results if r.status == StudentMatchStatus.MATCHED_BY_POSITION),
        n_desfase, n_only_teachers, n_only_official,
    )

    return results


# ========================================
# comparator.py
# ========================================

"""
comparator.py
=============

Aplica la regla de tolerancia (config.TOLERANCE) a los pares de
estudiantes ya emparejados por matcher.py, y produce un StudentAuditRow
por estudiante con el detalle de cada materia.

Este módulo NO sabe nada de Excel, colores ni resaltado — solo hace la
matemática de la comparación. Eso permite probarlo de forma aislada
(ver test_comparator.py) sin tocar archivos.
"""

import logging

from config import RELEVANT_SUBJECTS, SUBJECT_DISPLAY_NAMES, TOLERANCE

logger = logging.getLogger(__name__)


def _to_float(value: object) -> float | None:
    """
    Convierte el valor crudo de una celda a float, de forma defensiva.

    Casos que maneja sin lanzar excepción (devuelve None):
        - None / celda vacía
        - cadena vacía o solo espacios
        - texto no numérico

    Casos que sí convierte correctamente:
        - int / float ya nativos de openpyxl
        - texto con punto decimal ("8.5")
        - texto con coma decimal ("8,5") — común en configuración regional
        - texto con espacios alrededor (" 8.5 ")
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text == "":
        return None

    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _compare_subject(
    subject_key: str, teacher_raw: object, official_raw: object, tolerance: float
) -> ComparisonResult:
    teacher_grade = _to_float(teacher_raw)
    official_grade = _to_float(official_raw)
    display_name = SUBJECT_DISPLAY_NAMES.get(subject_key, subject_key)

    if teacher_grade is None or official_grade is None:
        return ComparisonResult(
            subject_key=subject_key,
            subject_display=display_name,
            teacher_grade=teacher_grade,
            official_grade=official_grade,
            difference=None,
            status=GradeStatus.SIN_DATO,
        )

    # round() evita ruido de punto flotante (p. ej. 0.020000000000003)
    # antes de comparar contra la tolerancia.
    difference = round(abs(teacher_grade - official_grade), 4)
    status = GradeStatus.OK if difference <= tolerance else GradeStatus.ERROR

    return ComparisonResult(
        subject_key=subject_key,
        subject_display=display_name,
        teacher_grade=teacher_grade,
        official_grade=official_grade,
        difference=difference,
        status=status,
    )


def compare_students(
    match_results: list[MatchResult], tolerance: float = TOLERANCE
) -> list[StudentAuditRow]:
    """
    Convierte cada MatchResult en un StudentAuditRow con todas las
    comparaciones de materias relevantes (O(n × m): n estudiantes,
    m materias configuradas en RELEVANT_SUBJECTS).
    """
    audit_rows: list[StudentAuditRow] = []

    for match in match_results:
        if match.status in (StudentMatchStatus.ONLY_IN_TEACHERS, StudentMatchStatus.ONLY_IN_OFFICIAL):
            # No hay con qué comparar: el estudiante falta en uno de los dos archivos.
            student = match.teacher if match.teacher is not None else match.official
            teacher_excel_row = match.teacher.excel_row if match.teacher is not None else None
            audit_rows.append(
                StudentAuditRow(
                    student_name=student.raw_name,
                    match_status=match.status,
                    excel_row=teacher_excel_row,
                    comparisons=[],
                )
            )
            continue

        comparisons = [
            _compare_subject(
                subject_key,
                match.teacher.grades.get(subject_key),
                match.official.grades.get(subject_key),
                tolerance,
            )
            for subject_key in RELEVANT_SUBJECTS
        ]
        audit_rows.append(
            StudentAuditRow(
                student_name=match.teacher.raw_name,
                match_status=match.status,
                excel_row=match.teacher.excel_row,
                comparisons=comparisons,
            )
        )

    n_errors = sum(1 for row in audit_rows if row.has_error)
    logger.info("Comparación completa: %d estudiantes procesados, %d con al menos un error.", len(audit_rows), n_errors)

    return audit_rows
