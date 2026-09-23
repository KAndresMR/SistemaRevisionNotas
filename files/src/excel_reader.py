"""
excel_reader.py
================

Responsable ÚNICAMENTE de leer los archivos .xlsx y convertirlos en
estructuras de datos en memoria (`list[Student]`). No compara nada,
no normaliza tolerancias, no genera reportes.

Detecta las columnas por NOMBRE de encabezado (no por posición fija),
para tolerar que ambos archivos tengan encabezados en distinto orden.

--------------------------------------------------------------------
SUPUESTOS SOBRE LA ESTRUCTURA DE CADA ARCHIVO
--------------------------------------------------------------------

ARCHIVO DE DOCENTES (estructura "plana"):
    Una sola fila de encabezados. Cada materia es UNA columna con el
    nombre exacto definido en config.RELEVANT_SUBJECTS (teacher_column).

ARCHIVO OFICIAL DEL SISTEMA (estructura "de dos niveles"):
    - Una fila con el NOMBRE DE LA MATERIA, típicamente en una celda
      combinada (merged cell) que cubre las columnas I, II, III, Nota Final.
    - La fila inmediatamente debajo con las sub-columnas: "I", "II",
      "III", "Nota Final".
    - P.GENERAL normalmente es una columna plana (sin sub-columnas).

Estos supuestos son razonables para el formato descrito, pero como no
se ha probado aún contra el archivo real, cualquier desajuste debe
reportarse como un error CLARO (no fallar silenciosamente) para poder
ajustar la detección en la siguiente iteración.
"""

import logging
import os
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import (
    RELEVANT_SUBJECTS,
    TEACHER_SUBJECT_ALIASES,
    TEACHER_STUDENT_NAME_COLUMNS,
    OFFICIAL_STUDENT_NAME_COLUMNS,
    OFFICIAL_FINAL_GRADE_LABEL,
    KEYWORDS_TO_IGNORE_IN_STUDENTS,
)
from auditor import Student
from auditor import normalize_name

logger = logging.getLogger(__name__)

# Cuántas filas iniciales se inspeccionan para localizar encabezados.
# Los datos de estudiantes empiezan después de la última fila de encabezado
# detectada, no en una fila fija asumida de antemano.
MAX_HEADER_SEARCH_ROWS = 15


class ExcelStructureError(Exception):
    """Se lanza cuando no se puede localizar una columna/encabezado esperado.

    Se usa deliberadamente en vez de continuar con datos incompletos:
    la prioridad #1 del sistema es correctitud, así que un archivo con
    estructura inesperada debe detener el proceso con un mensaje claro,
    no producir un reporte silenciosamente incompleto.
    """
    pass


def _cell_text(value) -> str:
    """Convierte el valor de una celda a texto normalizado para comparar
    contra nombres de columna esperados (case/accent-insensitive)."""
    if value is None:
        return ""
    return normalize_name(str(value))


def _find_flat_header_row(ws: Worksheet, required_columns: list[str]) -> tuple[int, dict[str, int]]:
    """
    Busca, dentro de las primeras MAX_HEADER_SEARCH_ROWS filas, la fila que
    contiene TODAS las columnas de `required_columns` (comparación
    normalizada). Devuelve (número_de_fila, {nombre_columna: índice_columna}).

    Usado para el archivo de docentes, que tiene un único nivel de encabezado.
    """
    required_normalized = [normalize_name(c) for c in required_columns]

    for row_idx in range(1, MAX_HEADER_SEARCH_ROWS + 1):
        row_values = {
            _cell_text(cell.value): cell.column
            for cell in ws[row_idx]
            if cell.value is not None
        }
        if all(req in row_values for req in required_normalized):
            column_map = {
                original: row_values[normalize_name(original)]
                for original in required_columns
            }
            return row_idx, column_map

    raise ExcelStructureError(
        f"No se encontró una fila de encabezado (dentro de las primeras "
        f"{MAX_HEADER_SEARCH_ROWS} filas) que contenga todas las columnas: "
        f"{required_columns}"
    )


def _find_first_matching_column(
    ws: Worksheet, header_row: int, candidates: tuple[str, ...]
) -> int:
    """Devuelve el índice de columna de la primera cabecera en `header_row`
    cuyo texto normalizado coincide con alguno de los `candidates`."""
    candidates_normalized = {normalize_name(c) for c in candidates}
    for cell in ws[header_row]:
        if _cell_text(cell.value) in candidates_normalized:
            return cell.column
    raise ExcelStructureError(
        f"No se encontró ninguna columna llamada {candidates} en la fila {header_row}."
    )


import difflib

def _find_teacher_columns(ws: Worksheet) -> tuple[int, int, dict[str, int]]:
    """
    Localiza las columnas de materias y la columna del estudiante en una hoja de docentes
    usando un enfoque heurístico (Smart Parsing).
    """
    student_col = None
    header_row = None
    
    # 1. Buscar la columna del estudiante dinámicamente
    for r in range(1, min(ws.max_row + 1, 150)):
        for cell in ws[r]:
            if cell.value is None:
                continue
            text = _cell_text(cell.value)
            if not text or len(text) < 4:
                continue
            
            for cand in TEACHER_STUDENT_NAME_COLUMNS:
                cand_norm = normalize_name(cand)
                if len(cand_norm) >= 4 and (cand_norm in text or text in cand_norm):
                    # Verificar heurísticamente: mirar 2-3 celdas abajo para ver si hay contenido real
                    valid_content = False
                    for offset in range(1, 4):
                        test_cell = ws.cell(row=r + offset, column=cell.column).value
                        test_text = _cell_text(test_cell)
                        if test_text and len(test_text) > 5 and not test_text.isdigit():
                            valid_content = True
                            break
                    
                    if valid_content:
                        student_col = cell.column
                        header_row = r
                        break
            if student_col:
                break
        if student_col:
            break

    if student_col is None or header_row is None:
        raise ExcelStructureError(
            f"No se encontró la columna de estudiante (Nómina/Nombres) en las primeras 150 filas. "
            f"Asegúrate de que exista una columna con nombres válidos."
        )

    # 2. Buscar materias en la fila del encabezado encontrado usando Fuzzy Matching
    subject_col_map: dict[str, int] = {}
    
    for cell in ws[header_row]:
        if cell.value is None:
            continue
        text = _cell_text(cell.value)
        if not text:
            continue
            
        norm_text = normalize_name(text)
        best_match = None
        best_score = 0.0
        
        for subject_key, candidates in TEACHER_SUBJECT_ALIASES.items():
            if subject_key in subject_col_map:
                continue
                
            for cand in candidates:
                norm_cand = normalize_name(cand)
                # Exact match o substring exacto
                if norm_cand == norm_text or (len(norm_cand) >= 4 and norm_cand in norm_text):
                    score = 1.0
                else:
                    # Fuzzy match
                    score = difflib.SequenceMatcher(None, norm_text, norm_cand).ratio()
                
                if score > best_score:
                    best_score = score
                    best_match = subject_key

        # Umbral de tolerancia de 0.65 para considerar que es la misma materia
        if best_match and best_score >= 0.65:
            subject_col_map[best_match] = cell.column

    missing = [k for k in RELEVANT_SUBJECTS if k not in subject_col_map]
    if missing:
        missing_names = [RELEVANT_SUBJECTS[k].teacher_column for k in missing]
        logger.warning(
            "Materias faltantes (incluso con fuzzy matching). Se omitirán: %s", missing_names
        )

    return header_row, student_col, subject_col_map


def _get_sheet(wb, sheet_name: str | None, file_path: str = "", expected_students: list[str] | None = None) -> Worksheet:
    """
    Obtiene la hoja requerida de forma segura y tolerante:
    1. Si no se especifica hoja, usa wb.active.
    2. Coincidencia exacta de nombre.
    3. Coincidencia insensible a mayúsculas y espacios en blanco.
    4. Si el archivo tiene una única hoja, se usa automáticamente.
    5. DATA-DRIVEN DISCOVERY: Si se proporcionan expected_students, busca la hoja que contenga más nombres.
    """
    if not sheet_name:
        return wb.active
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    norm = sheet_name.strip().lower()
    for s in wb.sheetnames:
        if s.strip().lower() == norm:
            return wb[s]

    if len(wb.sheetnames) == 1:
        logger.info(
            "El archivo '%s' tiene una única hoja ('%s'); usándola para '%s'.",
            os.path.basename(file_path) if file_path else "",
            wb.sheetnames[0],
            sheet_name,
        )
        return wb[wb.sheetnames[0]]

    # 5. DATA-DRIVEN DISCOVERY
    if expected_students:
        logger.info("Activando Data-Driven Discovery para encontrar la hoja '%s'.", sheet_name)
        best_sheet = None
        max_matches = 0
        expected_norm = {normalize_name(name) for name in expected_students}
        
        for s in wb.sheetnames:
            ws_test = wb[s]
            matches = 0
            # Escanear las primeras 150 filas para encontrar alumnos
            for row in ws_test.iter_rows(min_row=1, max_row=150, values_only=True):
                for cell_val in row:
                    if isinstance(cell_val, str):
                        n = normalize_name(cell_val)
                        if n and n in expected_norm:
                            matches += 1
                            
            if matches > max_matches:
                max_matches = matches
                best_sheet = s
                
        if best_sheet and max_matches >= 3: # Umbral mínimo de confianza
            logger.info("Escaneo Inteligente exitoso: La hoja '%s' coincide con %d estudiantes.", best_sheet, max_matches)
            return wb[best_sheet]

    available = ", ".join(repr(s) for s in wb.sheetnames)
    filename = os.path.basename(file_path) if file_path else "el archivo"
    raise ExcelStructureError(
        f"No se encontró la hoja '{sheet_name}' ni se pudo adivinar de forma inteligente en {filename}. Hojas disponibles: [{available}]"
    )


def read_teacher_file(path: str, sheet_name: str | None = None, expected_students: list[str] | None = None) -> tuple[list[Student], str]:
    """
    Lee el Excel de docentes y devuelve (estudiantes, nombre_de_hoja_descubierta).
    """
    logger.info("Leyendo archivo de docentes: %s (hoja esperada: %s)", path, sheet_name or "activa")
    wb = load_workbook(path, data_only=True, read_only=False)
    try:
        ws = _get_sheet(wb, sheet_name, file_path=path, expected_students=expected_students)
        header_row, student_col, subject_col_map = _find_teacher_columns(ws)

        students: list[Student] = []
        row_index = 0
        for row in ws.iter_rows(min_row=header_row + 1):
            if len(row) < student_col:
                continue

            # Descartar filas de sub-encabezados como NOTA / CUALIT.
            row_texts = [_cell_text(c.value) for c in row[:min(len(row), 25)]]
            if "NOTA" in row_texts or "CUALIT" in row_texts or "CUALIT." in row_texts:
                continue

            name_cell = row[student_col - 1]
            if name_cell.value is None or str(name_cell.value).strip() == "":
                continue  # fila vacía, se ignora

            raw_name = str(name_cell.value).strip()
            norm_name = normalize_name(raw_name)

            # Descartar filas de resumen o firmas (Promedios, Rector, etc.)
            if any(kw in norm_name for kw in KEYWORDS_TO_IGNORE_IN_STUDENTS):
                logger.debug("Fila ignorada por palabra clave de resumen/pie: '%s'", raw_name)
                continue

            # Si hay una columna antes de la de nombre (típicamente "No", con un número secuencial),
            # verificar que sea numérica para descartar filas que no correspondan a estudiantes.
            if student_col > 1:
                first_cell_value = row[0].value if len(row) > 0 else None
                is_seq_num = (
                    isinstance(first_cell_value, (int, float))
                    or (isinstance(first_cell_value, str) and first_cell_value.strip().isdigit())
                )
                if not is_seq_num:
                    logger.debug("Fila ignorada (sin número secuencial en col 1): '%s'", raw_name)
                    continue

            grades: dict[str, object] = {}
            for subject_key in RELEVANT_SUBJECTS:
                col_idx = subject_col_map.get(subject_key)
                if col_idx is not None:
                    grades[subject_key] = row[col_idx - 1].value if len(row) >= col_idx else None
                else:
                    grades[subject_key] = None  # Materia no encontrada en el archivo

            students.append(
                Student(
                    raw_name=raw_name,
                    normalized_name=norm_name,
                    row_index=row_index,
                    excel_row=name_cell.row,
                    grades=grades,
                )
            )
            row_index += 1

        logger.info("Archivo de docentes: %d estudiantes leídos desde la hoja '%s'.", len(students), ws.title)
        return students, ws.title
    finally:
        wb.close()


def _build_subject_header_map(ws: Worksheet, subject_name_row: int) -> dict[int, str]:
    """
    Construye {índice_columna: texto_normalizado_de_materia} para la fila
    donde están los nombres de materia, expandiendo celdas combinadas
    (merged cells) para que TODAS las columnas cubiertas por el merge
    apunten al mismo texto de materia.
    """
    column_to_subject: dict[int, str] = {}

    # Primero, valores directos (celdas no combinadas o esquina superior-izq.
    # de una celda combinada, que es donde openpyxl guarda el valor real).
    for cell in ws[subject_name_row]:
        if cell.value is not None:
            column_to_subject[cell.column] = _cell_text(cell.value)

    # Expandir rangos combinados que caigan sobre esta fila.
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= subject_name_row <= merged_range.max_row:
            top_left_value = ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
            if top_left_value is None:
                continue
            text = _cell_text(top_left_value)
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                column_to_subject[col] = text

    return column_to_subject


def read_official_file(path: str, sheet_name: str | None = None) -> list[Student]:
    """
    Lee el Excel oficial del sistema (estructura de dos niveles: materia +
    sub-columnas I/II/III/Nota Final) y devuelve una lista de Student.

    Solo se extrae "Nota Final" de cada materia relevante, más P.GENERAL.
    """
    logger.info("Leyendo archivo oficial del sistema: %s", path)
    wb = load_workbook(path, data_only=True, read_only=False)  # read_only=False: necesitamos merged_cells
    try:
        ws = _get_sheet(wb, sheet_name, file_path=path)

        # 1. Localizar la fila de sub-encabezados que contiene "NOTA FINAL".
        final_label_normalized = normalize_name(OFFICIAL_FINAL_GRADE_LABEL)
        grade_header_row = None
        for row_idx in range(1, MAX_HEADER_SEARCH_ROWS + 1):
            texts = [_cell_text(cell.value) for cell in ws[row_idx]]
            if final_label_normalized in texts:
                grade_header_row = row_idx
                break
        if grade_header_row is None:
            raise ExcelStructureError(
                f"No se encontró ninguna fila con la etiqueta '{OFFICIAL_FINAL_GRADE_LABEL}' "
                f"dentro de las primeras {MAX_HEADER_SEARCH_ROWS} filas del archivo oficial."
            )

        # 2. La fila de nombres de materia se asume inmediatamente arriba.
        subject_name_row = grade_header_row - 1
        column_to_subject = _build_subject_header_map(ws, subject_name_row)

        # 3. Para cada materia relevante, encontrar su columna "Nota Final":
        #    buscar, en grade_header_row, la columna cuyo texto == "NOTA FINAL"
        #    Y cuya materia asociada (column_to_subject) coincide con el
        #    official_column configurado.
        subject_final_col: dict[str, int] = {}
        for cell in ws[grade_header_row]:
            if _cell_text(cell.value) != final_label_normalized:
                continue
            subject_text = column_to_subject.get(cell.column)
            if subject_text is None:
                continue
            for subject_key, mapping in RELEVANT_SUBJECTS.items():
                if subject_key == "promedio_general":
                    continue  # P.GENERAL se resuelve aparte, es columna plana
                if normalize_name(mapping.official_column) == subject_text:
                    subject_final_col[subject_key] = cell.column

        missing = [
            key for key in RELEVANT_SUBJECTS
            if key != "promedio_general" and key not in subject_final_col
        ]
        if missing:
            display = [RELEVANT_SUBJECTS[k].official_column for k in missing]
            logger.warning(
                "Materias faltantes en el archivo oficial (se omitirán): %s", display
            )

        # 4. P.GENERAL: columna plana, se busca por nombre directo en cualquiera
        #    de las filas de encabezado inspeccionadas.
        p_general_col = None
        for row_idx in range(1, MAX_HEADER_SEARCH_ROWS + 1):
            try:
                p_general_col = _find_first_matching_column(
                    ws, row_idx, (RELEVANT_SUBJECTS["promedio_general"].official_column,)
                )
                break
            except ExcelStructureError:
                continue
        if p_general_col is None:
            logger.warning(
                "No se encontró la columna '%s' en el archivo oficial. Se omitirá.",
                RELEVANT_SUBJECTS['promedio_general'].official_column,
            )
        else:
            subject_final_col["promedio_general"] = p_general_col

        # 5. Columna de nombre del estudiante.
        student_col = None
        for row_idx in range(1, MAX_HEADER_SEARCH_ROWS + 1):
            try:
                student_col = _find_first_matching_column(ws, row_idx, OFFICIAL_STUDENT_NAME_COLUMNS)
                break
            except ExcelStructureError:
                continue
        if student_col is None:
            raise ExcelStructureError(
                f"No se encontró columna de nombre de estudiante {OFFICIAL_STUDENT_NAME_COLUMNS} "
                f"en el archivo oficial."
            )

        # 6. Leer filas de datos, empezando justo después de la fila de sub-encabezados.
        students: list[Student] = []
        row_index = 0
        for row in ws.iter_rows(min_row=grade_header_row + 1):
            name_cell = row[student_col - 1]
            if name_cell.value is None or str(name_cell.value).strip() == "":
                continue

            raw_name = str(name_cell.value)
            grades: dict[str, object] = {}
            for subject_key in RELEVANT_SUBJECTS:
                col_idx = subject_final_col.get(subject_key)
                if col_idx is not None:
                    grades[subject_key] = row[col_idx - 1].value if len(row) >= col_idx else None
                else:
                    grades[subject_key] = None  # Materia no encontrada

            students.append(
                Student(
                    raw_name=raw_name,
                    normalized_name=normalize_name(raw_name),
                    row_index=row_index,
                    grades=grades,
                )
            )
            row_index += 1

        logger.info("Archivo oficial: %d estudiantes leídos.", len(students))
        return students
    finally:
        wb.close()
