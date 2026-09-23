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

from copy import copy
import logging
import os
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from config import SHEET_ORDER
from excel_reader import _find_teacher_columns
from auditor import GradeStatus, StudentAuditRow

logger = logging.getLogger(__name__)

# Paleta de colores para el archivo de docentes:
# - STUDENT_ERROR_FILL: Amarillo / ámbar suave para el nombre del estudiante con error
STUDENT_ERROR_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

# - GRADE_ERROR_FILL: Rojo suave / coral para la nota específica donde está la discrepancia
GRADE_ERROR_FILL = PatternFill(start_color="F1948A", end_color="F1948A", fill_type="solid")


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


def highlight_teacher_file(
    teacher_file: str,
    output_file: str,
    audit_rows: list[StudentAuditRow],
    sheet_name: str | None = None,
    reset_marked: bool = False,
    template_file: str | None = None,
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

    marked_students = 0
    marked_cells = 0

    for audit_row in audit_rows:
        if not audit_row.has_error:
            continue
        if audit_row.excel_row is None:
            continue

        row_idx = audit_row.excel_row

        # 1. Colorear celda de nombre del estudiante (columna de nómina/estudiante)
        name_cell = ws.cell(row=row_idx, column=student_col)
        name_cell.fill = STUDENT_ERROR_FILL
        marked_students += 1

        # 2. Colorear cada materia / promedio que tenga error y agregar nota emergente
        for comp in audit_row.comparisons:
            if comp.status == GradeStatus.ERROR:
                if comp.subject_key in subject_col_map:
                    col_idx = subject_col_map[comp.subject_key]
                    grade_cell = ws.cell(row=row_idx, column=col_idx)
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
    
    # Estilos
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Azul oscuro
    header_font = Font(color="FFFFFF", bold=True, size=12, name="Calibri")
    row_alt_fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid") # Gris claro
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Título principal
    ws.append(["RESUMEN DE AUDITORÍA (DISCREPANCIAS ENCONTRADAS)"])
    ws.merge_cells("A1:D1")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
    ws.row_dimensions[1].height = 30

    # Encabezados de tabla
    headers = ["Curso", "Estudiante", "Tipo de Problema", "Detalle de la Discrepancia"]
    ws.append(headers)
    for cell in ws[2]:
        cell.fill = PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid")
        cell.font = Font(bold=True)
        cell.alignment = center_align
    ws.row_dimensions[2].height = 20

    current_row = 3

    if not discrepancies:
        ws.append(["-", "No se encontraron discrepancias.", "✅ TODO CORRECTO", "Todo coincide perfectamente."])
        for cell in ws[current_row]:
            cell.alignment = center_align
    else:
        # Ordenar por curso y estudiante
        discrepancies.sort(key=lambda x: (x.get("curso", ""), x.get("estudiante", "")))
        
        for i, d in enumerate(discrepancies):
            curso = d.get("curso", "-")
            estudiante = d.get("estudiante", "-")
            estado_original = d.get("estado", "")
            materia_original = d.get("materia", "")
            
            tipo_problema = ""
            detalle = ""
            
            if estado_original == "ERROR":
                tipo_problema = "❌ ERROR DE NOTA"
                n_doc = d.get("docente", "-")
                n_sis = d.get("sistema", "-")
                detalle = f"{materia_original} (Docente: {n_doc} vs Sistema: {n_sis})"
            elif estado_original == "DESFASE DE NÓMINA":
                if materia_original == "SOLO EN DOCENTES":
                    tipo_problema = "⚠️ FALTA EN SISTEMA"
                    detalle = "El estudiante está en la matriz del docente pero no en el reporte oficial."
                elif materia_original == "SOLO EN SISTEMA":
                    tipo_problema = "⚠️ FALTA EN DOCENTE"
                    detalle = "El estudiante está en el reporte oficial pero falta en la matriz del docente."
                else:
                    tipo_problema = "⚠️ DESFASE"
                    detalle = "Inconsistencia en la nómina del estudiante."
            else:
                tipo_problema = "❓ OTRO"
                detalle = f"{materia_original}"

            ws.append([curso, estudiante, tipo_problema, detalle])
            
            fill_color = row_alt_fill if i % 2 == 0 else white_fill
            ws.cell(row=current_row, column=1).alignment = center_align
            ws.cell(row=current_row, column=2).alignment = left_align
            ws.cell(row=current_row, column=3).alignment = center_align
            ws.cell(row=current_row, column=4).alignment = left_align
            
            for col_idx in range(1, 5):
                ws.cell(row=current_row, column=col_idx).fill = fill_color
                
            ws.row_dimensions[current_row].height = 25
            current_row += 1

    # Ajuste de ancho de columnas
    ws.column_dimensions["A"].width = 15  # Curso
    ws.column_dimensions["B"].width = 40  # Estudiante
    ws.column_dimensions["C"].width = 25  # Tipo Problema
    ws.column_dimensions["D"].width = 60  # Detalle

    # Asegurar que esté al inicio
    wb._sheets = [ws] + [s for s in wb._sheets if s != ws]

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
