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
    output_file: str,
    course_name: str | None = None,
    teacher_sheet: str | None = None,
    official_sheet: str | None = None,
    teacher_marked_output: str | None = None,
    skip_mark_teacher: bool = False,
    reset_marked_teacher: bool = False,
    template_file: str | None = None,
) -> str | None:
    """
    Ejecuta la auditoría completa para un curso:
    1. Genera el Excel de reporte en `output_file`.
    2. Colorea los estudiantes y notas con error en el archivo de docentes,
       guardándolo por defecto como copia marcada (`<nombre>_marcado.xlsx`).

    course_name: nombre opcional del curso (p. ej. "8A Matutino"), usado
                 como nombre de la hoja principal del reporte.
    teacher_sheet / official_sheet: nombre de hoja a leer dentro de cada
                 archivo de entrada, si no es la hoja activa por defecto.
    teacher_marked_output: ruta personalizada para el archivo de docentes marcado.
                           Si no se indica, se usa `<teacher_file>_marcado.xlsx`.
    skip_mark_teacher: si es True, omite el marcado del archivo de docentes.
    reset_marked_teacher: si es True, inicia el archivo marcado desde cero en lugar de acumular.
    template_file: ruta al libro plantilla consolidado (ej. BS consolidado) para inicializar
                   el archivo marcado cuando el curso proviene de un archivo individual.

    Devuelve la ruta del archivo de docentes marcado (o None si se omitió).
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
            sheet_name=actual_sheet, # Usamos la pestaña que el algoritmo descubrió realmente
            reset_marked=reset_marked_teacher,
            template_file=template_file,
        )
        marked_path = teacher_marked_output

    logger.info("=== Auditoría finalizada%s ===", f" — {course_name}" if course_name else "")
    return marked_path, audit_rows
