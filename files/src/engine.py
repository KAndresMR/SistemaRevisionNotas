"""
engine.py
=========

Orquesta el flujo completo de auditoría para UN curso:

    leer docentes -> leer sistema -> emparejar -> comparar -> generar reporte

Este es el único módulo que conoce a todos los demás. Para procesar
varios cursos (8A, 8B, 8C...) simplemente se llama a `run_audit` una
vez por curso, con rutas de archivo distintas — no hay que duplicar
ni modificar lógica.
"""

import logging
import os

from auditor import compare_students
from excel_reader import read_official_file, read_teacher_file
from auditor import match_students

from teacher_highlighter import highlight_teacher_file

logger = logging.getLogger(__name__)


def run_audit(
    teacher_file: str,
    official_file: str,
    course_name: str | None = None,
    teacher_sheet: str | None = None,
    official_sheet: str | None = None,
    teacher_marked_output: str | None = None,
    skip_mark_teacher: bool = False,
    reset_marked_teacher: bool = False,
    template_file: str | None = None,
    fix_errors: bool = False,
    output_file: str | None = None,
) -> tuple[str | None, list, str]:
    """
    Ejecuta la auditoría completa para un curso:
    1. Lee y empareja estudiantes entre el archivo oficial y el del docente.
    2. Compara notas aplicando la tolerancia configurada.
    3. Marca visualmente los errores en el archivo de docentes (o corrige notas si fix_errors=True).

    Devuelve (ruta_marcada, audit_rows, actual_sheet).
    """
    logger.info("=== Iniciando auditoría%s ===", f" — {course_name}" if course_name else "")

    official = read_official_file(official_file, sheet_name=official_sheet)
    official_names = [s.raw_name for s in official]
    
    teachers, actual_sheet = read_teacher_file(teacher_file, sheet_name=teacher_sheet, expected_students=official_names)

    matches = match_students(teachers, official)
    audit_rows = compare_students(matches)

    marked_path: str | None = None
    if not skip_mark_teacher:
        if teacher_marked_output is None:
            base, ext = os.path.splitext(teacher_file)
            teacher_marked_output = f"{base}_marcado{ext}"
        highlight_teacher_file(
            teacher_file=teacher_file,
            output_file=teacher_marked_output,
            audit_rows=audit_rows,
            sheet_name=actual_sheet,
            reset_marked=reset_marked_teacher,
            template_file=template_file,
            fix_errors=fix_errors,
        )
        marked_path = teacher_marked_output

    logger.info("=== Auditoría finalizada%s ===", f" — {course_name}" if course_name else "")
    return marked_path, audit_rows, actual_sheet
