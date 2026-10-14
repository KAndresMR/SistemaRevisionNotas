

"""
auditor.py
==========
Módulo central de auditoría de calificaciones:
- Estructuras de datos (Student, MatchResult, StudentAuditRow)
- Normalización de nombres de estudiantes
- Emparejamiento inteligente de nómina (matching por posición + fuzzy matching para desfases)
- Comparación matemática de calificaciones aplicando tolerancia
"""

import difflib
import logging
import re
import unicodedata
from dataclasses import dataclass, field
from enum import Enum

from config import RELEVANT_SUBJECTS, SUBJECT_DISPLAY_NAMES, TOLERANCE

logger = logging.getLogger(__name__)


class GradeStatus(str, Enum):
    """Estado de una comparación de nota individual (una materia)."""
    OK = "OK"
    ERROR = "ERROR"
    SIN_DATO = "SIN DATO"  # nota vacía o no convertible a número en alguno de los dos archivos


class StudentMatchStatus(str, Enum):
    """Estado del emparejamiento de un estudiante entre ambos archivos."""
    MATCHED_BY_POSITION = "MATCHED_BY_POSITION"      # misma posición, mismo nombre normalizado
    MATCHED_BY_NAME = "MATCHED_BY_NAME"              # desfase de posición, encontrado por nombre
    MATCHED_WITH_TYPO = "MATCHED_WITH_TYPO"          # encontrado con similitud difusa (typo)
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
    difference: float | str | None
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
            StudentMatchStatus.MATCHED_WITH_TYPO,
            StudentMatchStatus.ONLY_IN_TEACHERS,
            StudentMatchStatus.ONLY_IN_OFFICIAL,
        )


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

    def try_match_by_name(teacher: Student) -> tuple[Student | None, StudentMatchStatus | None]:
        candidate = get_official_by_name().get(teacher.normalized_name)
        if candidate is not None and candidate.row_index not in matched_official_row_indices:
            return candidate, StudentMatchStatus.MATCHED_BY_NAME
            
        # Fallback: Fuzzy matching optimizado
        best_match = None
        best_score = 0.0
        t_norm = teacher.normalized_name
        
        for name, off_student in get_official_by_name().items():
            if off_student.row_index in matched_official_row_indices:
                continue
            
            # Filtro rápido O(len) antes de calcular matriz completa
            matcher = difflib.SequenceMatcher(None, t_norm, name)
            if matcher.quick_ratio() < 0.80:
                continue

            score = matcher.ratio()
            if score > best_score:
                best_score = score
                best_match = off_student
                
        # Tolerancia de 85% para typos
        if best_match is not None and best_score >= 0.85:
            return best_match, StudentMatchStatus.MATCHED_WITH_TYPO
            
        return None, None

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
        match, status = try_match_by_name(teacher)
        if match is not None:
            results.append(MatchResult(teacher, match, status))
            matched_official_row_indices.add(match.row_index)
        else:
            results.append(MatchResult(teacher, None, StudentMatchStatus.ONLY_IN_TEACHERS))

    # --- Docentes con más filas que el archivo oficial (cola sobrante) ---
    for teacher in teachers[n:]:
        match, status = try_match_by_name(teacher)
        if match is not None:
            results.append(MatchResult(teacher, match, status))
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

    # Caso 1: Ambos vacíos -> Sin dato legítimo (no hay nota registrada en ninguna fuente)
    if teacher_grade is None and official_grade is None:
        return ComparisonResult(
            subject_key=subject_key,
            subject_display=display_name,
            teacher_grade=None,
            official_grade=None,
            difference=None,
            status=GradeStatus.SIN_DATO,
        )

    # Caso 2: Falta nota en docente pero sí existe en sistema oficial
    if teacher_grade is None and official_grade is not None:
        return ComparisonResult(
            subject_key=subject_key,
            subject_display=display_name,
            teacher_grade=None,
            official_grade=official_grade,
            difference="Falta nota en Docente",
            status=GradeStatus.ERROR,
        )

    # Caso 3: Docente tiene nota pero falta en sistema oficial
    if teacher_grade is not None and official_grade is None:
        return ComparisonResult(
            subject_key=subject_key,
            subject_display=display_name,
            teacher_grade=teacher_grade,
            official_grade=None,
            difference="Falta nota en Sistema",
            status=GradeStatus.ERROR,
        )

    # Caso 4: Ambos tienen nota numérica -> Comparación matemática con tolerancia
    # round() evita ruido de punto flotante (p. ej. 0.020000000000003)
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
