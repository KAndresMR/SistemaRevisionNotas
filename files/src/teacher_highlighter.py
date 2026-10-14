"""
teacher_highlighter.py
======================

Se encarga de marcar visualmente (colorear) en el archivo de docentes
(BS MAT ANUAL MATRIZ TUTOR.xlsx o una copia de él):
  1. El estudiante con discrepancias (celda de NÓMINA con color ámbar/amarillo).
  2. La celda exacta de la nota con error (materia o promedio con color rojo/coral).

PRESERVACIÓN DE FORMATO:
  - Abre el libro con `data_only=False` para que las fórmulas (ej. =TRUNC(...))
    NO se pierdan ni se conviertan en texto plano.
  - No altera los datos, números ni fórmulas, únicamente la propiedad `.fill`
    de las celdas afectadas.
  - Mantiene todas las demás hojas ('octavo B', 'noveno A', etc.) intactas.
"""

from collections import defaultdict
from copy import copy
import glob
import logging
import os
import re
import unicodedata
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from config import SHEET_ORDER, get_school_logo_path
from excel_reader import _find_teacher_columns
from auditor import GradeStatus, StudentAuditRow

logger = logging.getLogger(__name__)

# Paleta de colores para el archivo de docentes (ARGB con canal Alpha=FF para opacidad 100% en Excel):
# - STUDENT_ERROR_FILL: Amarillo / ámbar suave para el nombre del estudiante con error (#FFF2CC -> FFFFF2CC)
STUDENT_ERROR_FILL = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")

# - GRADE_ERROR_FILL: Rojo suave / coral para la nota específica donde está la discrepancia (#F1948A -> FFF1948A)
GRADE_ERROR_FILL = PatternFill(start_color="FFF1948A", end_color="FFF1948A", fill_type="solid")


def _copy_worksheet(src_ws: Worksheet, dest_ws: Worksheet) -> None:
    """
    Copia íntegramente una hoja de cálculo a otra en un libro distinto:
    dimensiones de columnas/filas, celdas combinadas, fórmulas, valores y estilos visuales.
    """
    # Dimensiones de columnas y filas
    for col_letter, dim in src_ws.column_dimensions.items():
        dest_ws.column_dimensions[col_letter].width = dim.width
    for row_idx, dim in src_ws.row_dimensions.items():
        dest_ws.row_dimensions[row_idx].height = dim.height

    # Celdas combinadas
    for merged_range in src_ws.merged_cells.ranges:
        dest_ws.merge_cells(str(merged_range))

    # Celdas: valores, fórmulas y formatos
    for row in src_ws.iter_rows():
        for cell in row:
            dest_cell = dest_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                dest_cell.font = copy(cell.font)
                dest_cell.border = copy(cell.border)
                dest_cell.fill = copy(cell.fill)
                dest_cell.number_format = copy(cell.number_format)
                dest_cell.protection = copy(cell.protection)
                dest_cell.alignment = copy(cell.alignment)


def _ensure_school_logo_in_sheet(ws: Worksheet) -> None:
    """
    Adjunta el logotipo oficial de la escuela (logoEscuela.png) en la esquina superior izquierda
    (anclado en la celda A1) de la hoja de cálculo generada.
    """
    logo_path = get_school_logo_path()
    if not logo_path or not os.path.exists(logo_path):
        logger.debug("No se encontró el archivo 'logoEscuela.png' para adjuntar en la hoja '%s'.", ws.title)
        return

    try:
        from openpyxl.drawing.image import Image as OpenpyxlImage

        # Limpiar cualquier dibujo o imagen previa para evitar duplicados
        if hasattr(ws, "_images") and ws._images:
            ws._images.clear()

        img = OpenpyxlImage(logo_path)
        # Proporción original (388x449): tamaño institucional adecuado para la cabecera
        img.width = 80
        img.height = 93
        ws.add_image(img, "A1")
        logger.debug("Logotipo de la escuela adjuntado en la hoja '%s' (celda A1).", ws.title)
    except Exception as e:
        logger.warning("No se pudo adjuntar el logotipo en la hoja '%s': %s", ws.title, str(e))


def _detect_jornada(filename: str) -> str:
    """Detecta si el archivo es Matutina o Vespertina basándose en el nombre."""
    upper = os.path.basename(filename).upper()
    # Revisar vespertina PRIMERO para evitar falsos positivos con "MATRIZ" que contiene "MAT"
    if "VESP" in upper:
        return "VESPERTINA"
    # Patrones vespertinos en nombres de cursos individuales (ej: 2AV, 8BV)
    if re.search(r'\d+[A-Z]?V\.', upper):
        return "VESPERTINA"
    # Default: matutina
    return "MATUTINA"


def _clean_tutor_name(raw_name: str) -> str:
    """Aplica heurística avanzada para extraer y limpiar nombres de docentes."""
    if not raw_name:
        return ""

    cleaned = raw_name
    
    # 1. Separar títulos pegados al nombre (ej: "Mgs.mónica" -> "Mgs. mónica")
    # Busca 2 a 4 letras seguidas de un punto, y luego inmediatamente otra letra
    cleaned = re.sub(r'([a-zA-Z]{2,4}\.)([a-zA-Z]+)', r'\1 \2', cleaned)
    
    # 2. Eliminar palabras basura severas institucionales/estructurales (case insensitive)
    junk_words = [
        r'\bmatriz\b', r'\bfinal\b', r'\bdel\b', 
        r'\b(de\s+)?jornada(\s+matutina|\s+vespertina)?\b', 
        r'\bmatutina\b', r'\bvespertina\b',
        r'\bprimero\b', r'\bsegundo\b', r'\btercero\b', r'\bcuarto\b', 
        r'\bquinto\b', r'\bsexto\b', r'\bséptimo\b', r'\bseptimo\b', 
        r'\boctavo\b', r'\bnoveno\b', r'\bdécimo\b', r'\bdecimo\b',
        r'\bbásica\b', r'\bbasica\b', r'\bbachillerato\b', r'\baño\b', r'\bgrado\b',
        r'\bcurso\b', r'\bparalelo\b', r'\bdocente\b', r'\btutor\b', r'\btutora\b',
        r'\bescuela\b', r'\bunidad\b', r'\beducativa\b'
    ]
    for jw in junk_words:
        cleaned = re.sub(jw, ' ', cleaned, flags=re.IGNORECASE)
        
    # 3. Eliminar paralelos sueltos (A, B, C...) aislados
    cleaned = re.sub(r'\b[A-Z]\b', ' ', cleaned, flags=re.IGNORECASE)
    
    # 4. Limpiar caracteres especiales sueltos (:, -) y normalizar espacios
    cleaned = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\.]', ' ', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    # Si todo era basura (ej: "Matriz Final Del Tutor"), devolvemos vacío
    if not cleaned or len(cleaned) < 3:
        return ""
    
    # 5. Aplicar Title Case inteligente, respetando acrónimos de títulos
    words = cleaned.split()
    titulos_conocidos = {"mgs.", "mgs", "mgtr.", "mgtr", "lic.", "lcdo.", "lcda.", "ing.", "prof.", "psic.", "tec.", "dr.", "dra."}
    final_words = [w.capitalize() for w in words]
            
    # Validar que no haya quedado SOLO un título sin nombre (ej. "Mgs. Lic.")
    if all(w.lower() in titulos_conocidos for w in final_words):
        return ""
            
    return " ".join(final_words)

def _extract_tutor_from_source(filename: str, sheet_name: str | None = None) -> str:
    """Extrae el nombre del Docente Tutor desde el archivo Excel original o consolidados en el mismo directorio."""
    
    def _search_in_file(fpath: str) -> str:
        try:
            src_wb = load_workbook(fpath, data_only=True)
            src_ws = None
            if sheet_name and sheet_name in src_wb.sheetnames:
                src_ws = src_wb[sheet_name]
            elif sheet_name:
                norm = sheet_name.strip().lower()
                for s in src_wb.sheetnames:
                    if s.strip().lower() == norm:
                        src_ws = src_wb[s]
                        break
            if src_ws is None:
                src_wb.close()
                return ""

            max_r = src_ws.max_row
            # Buscar en toda la hoja, ya que puede estar en el encabezado o al final
            for r in range(1, max_r + 1):
                for c in range(1, 15):
                    cell_val = src_ws.cell(row=r, column=c).value
                    val = str(cell_val or "").strip()
                    val_lower = val.lower()
                    
                    if not val:
                        continue
                        
                    if "tutor" in val_lower:
                        # 1. Dentro de la misma celda (ej. "Tutor: Lcdo. Juan" o "TUTOR MGS.MÓNICA CALLE")
                        cleaned_val = re.sub(r'(?i)^(docente\s+tutor|tutor)\s*:?\s*', '', val).strip()
                        if cleaned_val and len(cleaned_val) > 3:
                            src_wb.close()
                            return _clean_tutor_name(cleaned_val)
                        
                        # 2. Una fila arriba (Formato consolidado)
                        if r > 1:
                            name_above = str(src_ws.cell(row=r - 1, column=c).value or "").strip()
                            if name_above and len(name_above) > 3 and name_above.lower() not in ["promedio"]:
                                src_wb.close()
                                return _clean_tutor_name(name_above)
                                
                        # 3. Columna a la derecha
                        if c < 15:
                            name_right = str(src_ws.cell(row=r, column=c + 1).value or "").strip()
                            if name_right and len(name_right) > 3:
                                src_wb.close()
                                return _clean_tutor_name(name_right)
                                
                        # 4. Una fila abajo
                        if r < max_r:
                            name_below = str(src_ws.cell(row=r + 1, column=c).value or "").strip()
                            if name_below and len(name_below) > 3:
                                src_wb.close()
                                return _clean_tutor_name(name_below)
                                
                    # 5. Heurística para buscar en el encabezado (primeras 15 filas) sin la palabra tutor
                    if r <= 15:
                        # Si contiene un título académico pero no es rectora ni vicerrectora
                        titles = ["mgs.", "mgs ", "lic.", "lcdo.", "lcda.", "ing.", "prof.", "psic."]
                        if any(t in val_lower for t in titles):
                            if not any(x in val_lower for x in ["rector", "vicerrector", "inspector"]):
                                # Podría ser "SÉPTIMO A MGS. LIDIA" o solo el nombre
                                src_wb.close()
                                return _clean_tutor_name(val)
                                
            src_wb.close()
        except Exception:
            pass
        return ""

    # 1. Buscar en el archivo principal
    tutor = _search_in_file(filename)
    if tutor:
        return tutor
        
    # 2. Buscar en otros archivos del mismo directorio (ej. el consolidado)
    parent_dir = os.path.dirname(filename)
    if parent_dir and os.path.isdir(parent_dir):
        for other_file in glob.glob(os.path.join(parent_dir, "*.xlsx")):
            if os.path.abspath(other_file) != os.path.abspath(filename):
                tutor = _search_in_file(other_file)
                if tutor:
                    return tutor

                
    return ""


def _ensure_footer_in_sheet(ws: Worksheet, audit_rows: list[StudentAuditRow], filename: str, sheet_name: str | None = None) -> None:
    """
    Garantiza que exista un footer correcto (Rectora, Vicerrectora, Docente Tutor)
    debajo de las notas con fuentes uniformes de tamaño 12 y espaciado equidistante.
    """
    valid_rows = [ar.excel_row for ar in audit_rows if ar.excel_row is not None]
    if not valid_rows:
        return

    last_student_row = max(valid_rows)
    last_data_row = last_student_row

    # Buscar si existe "Promedio" debajo del último estudiante
    for r in range(last_student_row, last_student_row + 5):
        for c in range(1, 15):
            val = str(ws.cell(row=r, column=c).value or "").strip().lower()
            if "promedio" in val:
                last_data_row = max(last_data_row, r)

    # Revisar si el footer ya existe en la hoja de trabajo actual
    existing_footer_row = None
    for r in range(last_data_row + 1, last_data_row + 15):
        for c in range(1, 15):
            val = str(ws.cell(row=r, column=c).value or "").strip().lower()
            if "rectora" in val or "vicerrectora" in val or "tutor" in val:
                existing_footer_row = r
                break
        if existing_footer_row:
            break

    if existing_footer_row:
        # El footer ya existe en la plantilla original.
        # Aseguramos que todas sus celdas tengan tamaño 12 según las directrices institucionales
        for r in [existing_footer_row - 1, existing_footer_row]:
            ws.row_dimensions[r].height = 22
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                if type(cell).__name__ == "MergedCell":
                    continue
                if cell.value:
                    val_str = str(cell.value).strip()
                    if val_str and not val_str.startswith("="):
                        is_name_row = (r == existing_footer_row - 1)
                        cell.font = Font(
                            name="Calibri",
                            size=12,
                            bold=is_name_row or bool(cell.font and cell.font.bold),
                            color=cell.font.color if cell.font and cell.font.color else None,
                        )
                        if not cell.alignment or cell.alignment.horizontal is None:
                            cell.alignment = Alignment(horizontal="center", vertical="center")
        return

    # Si no existe, creamos uno nuevo dejando exactamente 4 filas vacías de separación
    footer_start_row = last_data_row + 5

    # Datos fijos
    rectora_name = "Ing. Sonia Chicaiza Mgtr."
    rectora_title = "Rectora (e)"

    # Vicerrectora según jornada
    jornada = _detect_jornada(filename)
    if jornada == "VESPERTINA":
        vicerrectora_name = "Lcda. Lucia Uzhca Mgtr."
        vicerrectora_title = "Vicerrectora (e) Vespertina"
    else:
        vicerrectora_name = "Lcda. Isabel Carpio Mgtr."
        vicerrectora_title = "Vicerrectora (e)Matutina"

    # Tutor: extraer del archivo original
    tutor_name = _extract_tutor_from_source(filename, sheet_name)
    if not tutor_name:
        tutor_name = ""
    tutor_title = "Docente Tutor"
    if tutor_name and any(f in tutor_name.lower() for f in ["lcda.", "licda.", "mgs. monica", "mgs. lidia", "mgs. ana", "dra."]):
        tutor_title = "Docente Tutora"

    # Determinar columnas equidistantes para que Docente Tutor no quede pegado a Vicerrectorado
    # y la distancia entre Rectorado y Vicerrectorado sea casi igual a la de Vicerrectorado y Tutor
    max_c = ws.max_column
    if max_c >= 20:
        # Matrices con 20 o más columnas (ej: octavo A, 4AM, 4BM)
        c_rec = 3
        c_vic = 11
        c_tut = 19
        span = 3
    elif max_c >= 14:
        # Matrices medianas (ej: 6BM, 6AM con 15-17 columnas)
        c_rec = 2
        c_vic = 7
        c_tut = 12
        span = 3
    else:
        # Matrices estrechas (<14 columnas)
        step = max(3, (max_c - 2) // 2)
        c_rec = 2
        c_vic = c_rec + step
        c_tut = c_vic + step
        span = max(1, min(2, step - 1))

    # Asegurar que el bloque de tutor no exceda el límite derecho
    if c_tut + span - 1 > max_c:
        span = max(1, max_c - c_tut + 1)

    ws.row_dimensions[footer_start_row].height = 22
    ws.row_dimensions[footer_start_row + 1].height = 22

    def _safe_merge_and_write(r: int, c_start: int, c_end: int, text: str, is_bold: bool) -> None:
        ranges_to_remove = []
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row <= r <= rng.max_row and not (rng.max_col < c_start or rng.min_col > c_end):
                ranges_to_remove.append(rng)
        for rng in ranges_to_remove:
            ws.merged_cells.remove(rng)

        # Convertir cualquier MergedCell residual en Cell estándar para permitir escritura
        from openpyxl.cell import Cell
        for c_idx in range(c_start, c_end + 1):
            if (r, c_idx) in ws._cells and type(ws._cells[(r, c_idx)]).__name__ == "MergedCell":
                ws._cells[(r, c_idx)] = Cell(ws, row=r, column=c_idx)

        if c_end > c_start:
            ws.merge_cells(start_row=r, start_column=c_start, end_row=r, end_column=c_end)

        top_cell = ws.cell(row=r, column=c_start)
        top_cell.value = text
        top_cell.font = Font(name="Calibri", size=12, bold=is_bold)
        top_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Fila de nombres (negrita, tamaño 12)
    _safe_merge_and_write(footer_start_row, c_rec, c_rec + span - 1, rectora_name, True)
    _safe_merge_and_write(footer_start_row, c_vic, c_vic + span - 1, vicerrectora_name, True)
    _safe_merge_and_write(footer_start_row, c_tut, c_tut + span - 1, tutor_name, True)

    # Fila de cargos / títulos (regular, tamaño 12)
    _safe_merge_and_write(footer_start_row + 1, c_rec, c_rec + span - 1, rectora_title, False)
    _safe_merge_and_write(footer_start_row + 1, c_vic, c_vic + span - 1, vicerrectora_title, False)
    _safe_merge_and_write(footer_start_row + 1, c_tut, c_tut + span - 1, tutor_title, False)



def _strip_accents(s: str) -> str:
    """Elimina tildes de una cadena para comparaciones robustas."""
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def _canonical_header_info(raw_val: str, is_last_col: bool = False) -> tuple[str, str | None]:
    """
    Normaliza y corrige la ortografía de cada encabezado de la matriz escolar,
    asegurando mayúsculas impecables, tildes correctas y su respectivo color institucional.
    Devuelve (texto_canonico_corregido, color_hex).
    """
    if raw_val is None:
        return ("", None)
    raw_str = str(raw_val).strip()
    if not raw_str or raw_str == "-":
        return ("", None)

    clean = re.sub(r"\s+", " ", raw_str.replace("\n", " ")).strip()
    unacc = _strip_accents(clean).upper()
    unacc_no_space = re.sub(r"[^A-Z]", "", unacc)

    # 1. Identificadores de nómina y alumnos
    if unacc in ["NO", "NO.", "N", "NRO", "NUMERO", "N°"]:
        return ("N°", "FFD8D8D8")
    if "CEDULA" in unacc:
        return ("CÉDULA", "FFD8D8D8")
    if any(k in unacc for k in ["NOMINA", "ESTUDIANTE", "APELLIDO"]):
        if "NOMINA" in unacc:
            return ("NÓMINA", "FFD8D8D8")
        return ("ESTUDIANTES APELLIDOS / NOMBRES", "FFD8D8D8")

    # 2. Materias principales
    if "LENGUA" in unacc or "LITERATURA" in unacc:
        return ("LENGUA Y LITERATURA", "FFB4C6E7")
    if "MATEM" in unacc:
        return ("MATEMÁTICA", "FF94DCF8")
    if "CIENCIA" in unacc or "NATURAL" in unacc:
        return ("CIENCIAS NATURALES", "FFFFE598")
    if "SOCIAL" in unacc:
        return ("ESTUDIOS SOCIALES", "FFF7CAAC")
    if "INGL" in unacc:
        return ("INGLÉS", "FFDEEAF6")
    if unacc == "ECA":
        return ("ECA", "FFD0CECE")
    if "CULTURAL" in unacc or "ART" in unacc:
        return ("EDUCACIÓN CULTURAL Y ARTÍSTICA", "FFD0CECE")
    if "FISIC" in unacc:
        return ("EDUCACIÓN FÍSICA", "FFFFD966")
    if "ANIM" in unacc_no_space or "LECTURA" in unacc:
        return ("ANIMACIÓN A LA LECTURA", "FFA0FCF3")
    if unacc == "CAI":
        return ("CAI", "FFF8FF9D")
    if "VOCACIONAL" in unacc:
        return ("ORIENTACIÓN VOCACIONAL", "FFB6D7A8")
    if "AULA" in unacc:
        return ("ACOMPAÑAMIENTO INTEGRAL EN EL AULA", "FFF8FF9D")
    if "CIVICA" in unacc or "ACOMPAN" in unacc:
        return ("CÍVICA Y ACOMPAÑAMIENTO INTEGRAL", "FFF8FF9D")

    # 3. Columnas de resumen general
    if "PROMEDIO FINAL" in unacc:
        return ("PROMEDIO FINAL", "FF92D050")
    if "COMPORT" in unacc:
        return ("COMPORTAMIENTO", "FFCC99FF")
    if "PROMOCI" in unacc:
        return ("PROMOCIÓN", "FFBFBFBF")
    if "CUALITAT" in unacc or "CUALIT" in unacc:
        return ("CUALITATIVA", "FFE5E7EB")
    if "PROMEDIO" in unacc:
        # En plantillas donde la última columna fue nombrada erróneamente PROMEDIO en vez de PROMOCIÓN
        if is_last_col:
            return ("PROMOCIÓN", "FFBFBFBF")
        return ("PROMEDIO", "FF92D050")
    if "OBSERVAC" in unacc:
        return ("OBSERVACIONES", "FFBFBFBF")

    return (clean.upper(), None)


def _apply_institutional_colors_to_sheet(
    ws: Worksheet,
    header_row: int,
    audit_rows: list[StudentAuditRow],
    student_col: int | None = None
) -> None:
    """
    Aplica el formato institucional con la paleta de colores oficial a las columnas y cabeceras
    (tal como en la matriz oficial de la institución), corrigiendo inconsistencias ortográficas:
      - Cabeceras de materias con texto hacia arriba (textRotation=90) y su color pastel oficial.
      - Columnas de notas con su respectivo color pastel por materia (ambas columnas si hay nota y cualitativa).
      - Cabeceras de nómina (N°, Cédula, Nombres) en gris institucional (#D8D8D8).
      - Detección y corrección de la subfila NOTA / CUALIT. si está presente en la hoja.
    """
    col_colors: dict[int, str] = {}
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    vertical_align = Alignment(horizontal="center", vertical="center", textRotation=90, wrap_text=True)
    header_font = Font(name="Calibri", size=10, bold=True)
    sub_font = Font(name="Calibri", size=9, bold=True)

    ws.row_dimensions[header_row].height = 145

    valid_rows = [ar.excel_row for ar in audit_rows if ar.excel_row is not None]
    min_student_row = min(valid_rows) if valid_rows else (header_row + 1)
    max_student_row = max(valid_rows) if valid_rows else header_row

    has_subheaders = (min_student_row > header_row + 1)
    sub_row = header_row + 1 if has_subheaders else None
    if sub_row:
        ws.row_dimensions[sub_row].height = 24

    # 1. Primera pasada: Normalizar ortografía de cabeceras y registrar colores de columnas
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=c)
        if type(cell).__name__ == "MergedCell":
            continue
        val_str = str(cell.value or "").strip()
        if not val_str:
            continue
        is_last = (c == ws.max_column or all(ws.cell(row=header_row, column=k).value is None for k in range(c + 1, ws.max_column + 1)))
        canon_name, color_hex = _canonical_header_info(val_str, is_last_col=is_last)
        if canon_name:
            cell.value = canon_name
            cell.font = header_font
            if color_hex:
                col_colors[c] = color_hex
                cell.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                if canon_name in ["N°", "CÉDULA", "NÓMINA", "ESTUDIANTES APELLIDOS / NOMBRES"]:
                    cell.alignment = center_align
                else:
                    cell.alignment = vertical_align
        elif val_str == "-":
            cell.value = None

    # 2. Extender colores a las columnas secundarias de materias combinadas (ej. Col E cuando D y E son LENGUA)
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row == header_row:
            if mr.min_col in col_colors:
                c_color = col_colors[mr.min_col]
                for sc in range(mr.min_col, mr.max_col + 1):
                    col_colors[sc] = c_color

    # 3. Si existe subfila de NOTA y CUALIT., estandarizar el texto y sus colores
    if sub_row:
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=sub_row, column=c)
            if type(cell).__name__ == "MergedCell":
                continue
            sub_val = str(cell.value or "").upper().strip()
            c_color = col_colors.get(c)

            is_numeric_sub = "NOTA" in sub_val
            is_cualit_sub = "CUALIT" in sub_val

            # Si el docente dejó la celda vacía o con texto corrupto, deducir por el tipo de dato del estudiante
            if not is_numeric_sub and not is_cualit_sub and c_color and valid_rows:
                first_st_val = ws.cell(row=min_student_row, column=c).value
                next_st_val = ws.cell(row=min_student_row, column=c + 1).value if c < ws.max_column else None
                if isinstance(first_st_val, (int, float)) and isinstance(next_st_val, str) and next_st_val in ["A+", "A-", "B+", "B-", "C+", "C-", "D", "E"]:
                    is_numeric_sub = True
                elif isinstance(first_st_val, str) and first_st_val in ["A+", "A-", "B+", "B-", "C+", "C-", "D", "E"]:
                    prev_st_val = ws.cell(row=min_student_row, column=c - 1).value if c > 1 else None
                    if isinstance(prev_st_val, (int, float)):
                        is_cualit_sub = True

            if is_numeric_sub:
                cell.value = "NOTA"
                cell.alignment = center_align
                cell.font = sub_font
                if c_color:
                    cell.fill = PatternFill(start_color=c_color, end_color=c_color, fill_type="solid")
            elif is_cualit_sub:
                cell.value = "CUALIT."
                cell.alignment = center_align
                cell.font = sub_font
                if c_color:
                    cell.fill = PatternFill(start_color=c_color, end_color=c_color, fill_type="solid")

    # 4. Aplicar color pastel y tamaño de letra 14pt a todas las celdas de estudiantes (nombres y notas)
    if valid_rows:
        for r in range(min_student_row, max_student_row + 1):
            ws.row_dimensions[r].height = 24  # Altura de fila adecuada para texto de 14pt
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                if type(cell).__name__ == "MergedCell":
                    continue
                color_hex = col_colors.get(c)
                if color_hex:
                    cell.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

                curr_font = cell.font
                is_bold = bool(curr_font and curr_font.bold)
                font_name = curr_font.name if curr_font and curr_font.name else "Calibri"
                font_color = curr_font.color if curr_font and curr_font.color else None
                cell.font = Font(
                    name=font_name,
                    size=12,
                    bold=is_bold,
                    italic=curr_font.italic if curr_font else False,
                    color=font_color,
                )

                s_col = student_col or 2
                if c == s_col:
                    if not cell.alignment or cell.alignment.horizontal != "left":
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    if not cell.alignment or cell.alignment.horizontal is None:
                        cell.alignment = Alignment(horizontal="center", vertical="center")

        # 5. Adaptar dinámicamente el ancho de la columna de estudiantes al nombre más largo de la hoja
        s_col = student_col or 2
        max_name_len = 0
        for r in range(min_student_row, max_student_row + 1):
            val_str = str(ws.cell(row=r, column=s_col).value or "").strip()
            if val_str and not val_str.startswith("=") and not any(k in val_str.lower() for k in ["promedio", "rector", "vicerrec", "tutor"]):
                max_name_len = max(max_name_len, len(val_str))

        if audit_rows:
            for ar in audit_rows:
                if ar.student_name:
                    max_name_len = max(max_name_len, len(ar.student_name.strip()))

        target_student_width = max(38.0, round(max_name_len * 1.15 + 5.0, 1)) if max_name_len > 0 else 38.0
        student_col_letter = get_column_letter(s_col)
        ws.column_dimensions[student_col_letter].width = target_student_width


def highlight_teacher_file(
    teacher_file: str,
    output_file: str,
    audit_rows: list[StudentAuditRow],
    sheet_name: str | None = None,
    reset_marked: bool = False,
    template_file: str | None = None,
    fix_errors: bool = False,
) -> int:
    """
    Abre el archivo de docentes preservando fórmulas, colorea a los estudiantes y celdas
    con error según `audit_rows`, y guarda el resultado en `output_file`.

    COMPORTAMIENTO ACUMULATIVO Y SOPORTE CONSOLIDADO:
    1. Si `output_file` ya existe (y reset_marked=False), se utiliza como base para
       acumular los cambios.
    2. Si `reset_marked=True` o `output_file` no existe:
       Se inicializa a partir de `template_file` (ej: el BS consolidado) si existe,
       o de lo contrario desde `teacher_file`.
    3. Si la hoja `sheet_name` ya existe en el libro (ej: "octavo A"), se edita directamente.
       Si NO existe (ej: "segundo A" proveniente del archivo individual 2AV.xlsx),
       se copia la hoja del archivo individual a `output_file` como una nueva hoja.
    4. Las hojas del libro se reordenan automáticamente según SHEET_ORDER (2do a 10mo).
    5. Se colorean en amarillo los nombres de estudiantes con discrepancias y en rojo
       las celdas exactas de materias o promedio donde está el error.

    Devuelve el total de estudiantes con error que fueron marcados.
    """
    target_sheet_name = sheet_name or "activa"

    if not reset_marked and os.path.exists(output_file):
        base_file = output_file
        logger.info(
            "Archivo marcado existente encontrado (%s). Actualizando hoja: '%s'",
            base_file,
            target_sheet_name,
        )
        wb = load_workbook(base_file, data_only=False)

        # Cargar siempre la versión limpia y fresca de la hoja desde teacher_file
        src_wb = load_workbook(teacher_file, data_only=False)
        src_ws = None
        if sheet_name and sheet_name in src_wb.sheetnames:
            src_ws = src_wb[sheet_name]
        elif sheet_name:
            norm = sheet_name.strip().lower()
            for s in src_wb.sheetnames:
                if s.strip().lower() == norm:
                    src_ws = src_wb[s]
                    break
        if src_ws is None:
            src_ws = src_wb.active

        actual_title = sheet_name or src_ws.title
        if actual_title in wb.sheetnames:
            sheet_idx = wb.sheetnames.index(actual_title)
            del wb[actual_title]
            ws = wb.create_sheet(title=actual_title, index=sheet_idx)
        else:
            ws = wb.create_sheet(title=actual_title)

        _copy_worksheet(src_ws, ws)
        src_wb.close()
        target_sheet_name = actual_title

    elif template_file and os.path.exists(template_file):
        base_file = template_file
        logger.info(
            "Iniciando nuevo archivo marcado desde plantilla consolidada: %s",
            base_file,
        )
        wb = load_workbook(base_file, data_only=False)
        target_sheet_name = sheet_name or wb.active.title
        if target_sheet_name in wb.sheetnames:
            ws = wb[target_sheet_name]
        else:
            src_wb = load_workbook(teacher_file, data_only=False)
            src_ws = src_wb.active
            ws = wb.create_sheet(title=target_sheet_name)
            _copy_worksheet(src_ws, ws)
            src_wb.close()

    else:
        base_file = teacher_file
        logger.info(
            "Iniciando nuevo archivo marcado desde archivo docente: %s",
            base_file,
        )
        wb = load_workbook(base_file, data_only=False)
        target_sheet_name = sheet_name or wb.active.title
        ws = wb[target_sheet_name] if target_sheet_name in wb.sheetnames else wb.active

    # Reordenar hojas preservando 'RESUMEN DISCREPANCIAS' al frente y orden canónico (2do a 10mo)
    front_sheets = [wb[s] for s in wb.sheetnames if s == "RESUMEN DISCREPANCIAS"]
    ordered_sheets = [wb[s] for s in SHEET_ORDER if s in wb.sheetnames and s != "RESUMEN DISCREPANCIAS"]
    other_sheets = [wb[s] for s in wb.sheetnames if s not in SHEET_ORDER and s != "RESUMEN DISCREPANCIAS"]
    wb._sheets = front_sheets + ordered_sheets + other_sheets

    # Localizar encabezados y columnas en la hoja
    header_row, student_col, subject_col_map = _find_teacher_columns(ws)

    # Aplicar la paleta institucional a todas las columnas de la matriz
    _apply_institutional_colors_to_sheet(ws, header_row, audit_rows, student_col=student_col)

    marked_students = 0
    marked_cells = 0

    for audit_row in audit_rows:
        if not audit_row.has_error:
            continue
        if audit_row.excel_row is None:
            continue

        row_idx = audit_row.excel_row

        # 1. Colorear celda de nombre del estudiante (columna de nómina/estudiante)
        # IMPORTANTE: Solo en modo Errores Resaltados (fix_errors=False).
        # En la matriz corregida (fix_errors=True) NO se resaltan los estudiantes.
        if not fix_errors:
            name_cell = ws.cell(row=row_idx, column=student_col)
            if type(name_cell).__name__ != "MergedCell":
                name_cell.fill = STUDENT_ERROR_FILL
        marked_students += 1

        # 2. Colorear cada materia / promedio que tenga error y agregar nota emergente
        for comp in audit_row.comparisons:
            if comp.status == GradeStatus.ERROR:
                if comp.subject_key in subject_col_map:
                    col_idx = subject_col_map[comp.subject_key]
                    grade_cell = ws.cell(row=row_idx, column=col_idx)
                    if fix_errors:
                        # Corregir la nota directamente SIN colores ni comentarios
                        if type(grade_cell).__name__ != "MergedCell":
                            grade_cell.value = comp.official_grade
                            grade_cell.comment = None
                    else:
                        if type(grade_cell).__name__ != "MergedCell":
                            grade_cell.fill = GRADE_ERROR_FILL

                        # Comentario emergente: SOLO muestra Nota Sistema y Diferencia
                        if isinstance(comp.official_grade, (int, float)):
                            sis_val = f"{comp.official_grade:.2f}"
                        elif comp.official_grade is not None and str(comp.official_grade).strip():
                            sis_val = str(comp.official_grade).strip()
                        else:
                            sis_val = "-"
    
                        if isinstance(comp.difference, (int, float)):
                            diff_val = f"{comp.difference:.2f}"
                        elif comp.difference is not None and str(comp.difference).strip():
                            diff_val = str(comp.difference).strip()
                        else:
                            diff_val = "-"
    
                        comment = Comment(f"Sistema: {sis_val}\nDiferencia: {diff_val}", "")
                        comment.width = 160
                        comment.height = 45
                        grade_cell.comment = comment

                    marked_cells += 1
                    logger.debug(
                        "Marcada celda de error [%s, col %d, fila %d]: %s",
                        comp.subject_display,
                        col_idx,
                        row_idx,
                        audit_row.student_name,
                    )

    # Validar y agregar footer si hace falta
    _ensure_footer_in_sheet(ws, audit_rows, teacher_file, sheet_name=target_sheet_name)

    # Adjuntar el logo oficial de la escuela en la parte superior izquierda de cada hoja de notas
    for sheet in wb.worksheets:
        if sheet.title != "RESUMEN DISCREPANCIAS":
            _ensure_school_logo_in_sheet(sheet)

    try:
        wb.save(output_file)
    except PermissionError:
        wb.close()
        raise PermissionError(
            f"No se pudo guardar el archivo marcado '{os.path.basename(output_file)}' porque está ABIERTO en Excel. "
            f"Por favor ciérralo en Excel y vuelve a ejecutar."
        )
    finally:
        wb.close()

    logger.info(
        "Archivo de docentes marcado guardado en: %s (%d estudiantes, %d celdas de nota con error)",
        output_file,
        marked_students,
        marked_cells,
    )
    return marked_students


def create_or_update_discrepancies_sheet(
    marked_file: str,
    discrepancies: list[dict],
    processed_courses: list[str] = None
) -> None:
    """
    Crea o actualiza la hoja 'RESUMEN DISCREPANCIAS' al inicio del libro.
    Usa una tabla única y directa con Emojis visuales y lenguaje humano.
    """
    if not os.path.exists(marked_file):
        return

    wb = load_workbook(marked_file, data_only=False)

    sheet_title = "RESUMEN DISCREPANCIAS"
    if sheet_title in wb.sheetnames:
        del wb[sheet_title]

    ws = wb.create_sheet(title=sheet_title, index=0)
    
    # Estilos estructurales con paleta institucional (ARGB 100% opaco para Excel)
    header_fill = PatternFill(start_color="FF1E3A8A", end_color="FF1E3A8A", fill_type="solid") # Azul oscuro
    header_font = Font(color="FFFFFFFF", bold=True, size=12, name="Calibri")
    course_fill = PatternFill(start_color="FF2563EB", end_color="FF2563EB", fill_type="solid") # Azul
    course_font = Font(color="FFFFFFFF", bold=True, size=11, name="Calibri")
    table_header_fill = PatternFill(start_color="FFE5E7EB", end_color="FFE5E7EB", fill_type="solid") # Gris suave
    table_header_font = Font(color="FF1E3A8A", bold=True)
    success_font = Font(bold=True, color="FF16A34A")
    
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(bottom=Side(style='thin', color='FFCCCCCC'))

    # Título principal
    ws.append(["RESUMEN DE AUDITORÍA (DISCREPANCIAS ENCONTRADAS)"])
    ws.merge_cells("A1:E1")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
    ws.row_dimensions[1].height = 30
    
    ws.append([]) # Fila en blanco
    
    current_row = 3

    if processed_courses is None:
        processed_courses = []
        
    # Agrupar discrepancias por curso
    discs_by_course = defaultdict(list)
    for d in discrepancies:
        discs_by_course[d.get("curso", "-")].append(d)
        
    # Asegurar que los cursos procesados estén en las claves
    for c in processed_courses:
        if c not in discs_by_course:
            discs_by_course[c] = []

    # Iterar por curso
    for curso, c_discs in sorted(discs_by_course.items()):
        # Título del curso
        ws.append([f"▶ CURSO: {curso}"])
        ws.merge_cells(start_row=current_row, end_row=current_row, start_column=1, end_column=5)
        cell = ws.cell(row=current_row, column=1)
        cell.fill = course_fill
        cell.font = course_font
        cell.alignment = left_align
        ws.row_dimensions[current_row].height = 20
        current_row += 1
        
        if not c_discs:
            ws.append(["✅ Todo el curso está perfectamente cuadrado. No hay discrepancias."])
            ws.merge_cells(start_row=current_row, end_row=current_row, start_column=1, end_column=5)
            cell = ws.cell(row=current_row, column=1)
            cell.font = success_font
            cell.alignment = center_align
            ws.row_dimensions[current_row].height = 25
            current_row += 1
        else:
            # Encabezados de tabla
            ws.append(["Estudiante", "Tipo de Problema", "Materia", "Docente", "Sistema"])
            for col_idx in range(1, 6):
                c = ws.cell(row=current_row, column=col_idx)
                c.fill = table_header_fill
                c.font = table_header_font
                c.alignment = center_align
            ws.row_dimensions[current_row].height = 20
            current_row += 1
            
            # Agrupar por estudiante
            discs_by_student = defaultdict(list)
            for d in c_discs:
                discs_by_student[d.get("estudiante", "-")].append(d)
                
            for estudiante, s_discs in sorted(discs_by_student.items()):
                start_row_student = current_row
                
                for idx, d in enumerate(s_discs):
                    estado_original = d.get("estado", "")
                    materia_original = d.get("materia", "")
                    
                    if estado_original == "ERROR":
                        tipo_problema = "❌ ERROR DE NOTA"
                        n_doc = d.get("docente", "-")
                        n_sis = d.get("sistema", "-")
                        mat = materia_original
                    elif estado_original == "DESFASE DE NÓMINA":
                        if materia_original == "SOLO EN DOCENTES":
                            tipo_problema = "⚠️ FALTA EN SISTEMA"
                            mat = "(Nómina)"
                            n_doc = "-"
                            n_sis = "Falta"
                        elif materia_original == "SOLO EN SISTEMA":
                            tipo_problema = "⚠️ FALTA EN DOCENTE"
                            mat = "(Nómina)"
                            n_doc = "Falta"
                            n_sis = "-"
                        else:
                            tipo_problema = "⚠️ DESFASE"
                            mat = "(Nómina)"
                            n_doc = "-"
                            n_sis = "-"
                    else:
                        tipo_problema = "❓ OTRO"
                        mat = materia_original
                        n_doc = "-"
                        n_sis = "-"
                        
                    ws.append([estudiante, tipo_problema, mat, n_doc, n_sis])
                    
                    # Formato de celdas individuales
                    ws.cell(row=current_row, column=3).alignment = left_align
                    ws.cell(row=current_row, column=4).alignment = center_align
                    ws.cell(row=current_row, column=5).alignment = center_align
                    
                    ws.row_dimensions[current_row].height = 20
                    current_row += 1
                    
                end_row_student = current_row - 1
                
                # Combinar celdas si hay más de 1 fila
                if end_row_student > start_row_student:
                    ws.merge_cells(start_row=start_row_student, end_row=end_row_student, start_column=1, end_column=1)
                    ws.merge_cells(start_row=start_row_student, end_row=end_row_student, start_column=2, end_column=2)
                
                # Alineación de las celdas combinadas o únicas
                ws.cell(row=start_row_student, column=1).alignment = center_align
                ws.cell(row=start_row_student, column=2).alignment = center_align
                
                # Bordes abajo de cada estudiante para separar visualmente
                for col_idx in range(1, 6):
                    ws.cell(row=end_row_student, column=col_idx).border = thin_border
                    
        ws.append([]) # Espacio extra entre cursos
        current_row += 1

    # Ajuste de ancho de columnas
    ws.column_dimensions["A"].width = 35  # Estudiante
    ws.column_dimensions["B"].width = 25  # Tipo Problema
    ws.column_dimensions["C"].width = 30  # Materia
    ws.column_dimensions["D"].width = 12  # Docente
    ws.column_dimensions["E"].width = 12  # Sistema

    # Asegurar que esté al inicio
    wb._sheets = [ws] + [s for s in wb._sheets if s != ws]

    # Garantizar el logotipo en todas las hojas del archivo consolidado/marcado
    for sheet in wb.worksheets:
        if sheet.title != "RESUMEN DISCREPANCIAS":
            _ensure_school_logo_in_sheet(sheet)

    try:
        wb.save(marked_file)
    except PermissionError:
        wb.close()
        raise PermissionError(
            f"No se pudo actualizar la hoja RESUMEN DISCREPANCIAS en '{os.path.basename(marked_file)}' "
            f"porque el archivo está ABIERTO en Excel. Por favor ciérralo y vuelve a intentar."
        )
    finally:
        wb.close()

    logger.info("Hoja '%s' actualizada con diseño simplificado en %s.", sheet_title, marked_file)
