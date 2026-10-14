"""
excel_preview.py
================
Genera representaciones HTML fieles y visuales de hojas de cálculo de Excel (OpenPyXL)
para previsualizar en Streamlit los archivos generados con sus colores, fuentes,
celdas combinadas, encabezados verticales hacia arriba y comentarios de auditoría.
"""

import base64
import os
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from config import get_school_logo_path

_LOGO_CACHE: str | None = None


def _get_logo_base64() -> str | None:
    global _LOGO_CACHE
    if _LOGO_CACHE is not None:
        return _LOGO_CACHE
    logo_path = get_school_logo_path()
    if logo_path and os.path.isfile(logo_path):
        try:
            with open(logo_path, "rb") as f:
                _LOGO_CACHE = base64.b64encode(f.read()).decode("utf-8")
                return _LOGO_CACHE
        except Exception:
            return None
    return None


# Palabras clave de materias y columnas de evaluación que deben mostrarse con texto vertical hacia arriba
SUBJECT_KEYWORDS = {
    "LENGUA", "MATEMÁTICA", "MATEMATICA", "CIENCIAS", "NATURALES", "ESTUDIOS", "SOCIALES",
    "INGLÉS", "INGLES", "CULTURAL", "ARTÍSTICA", "ARTISTICA", "EDUCACIÓN", "EDUCACION",
    "FÍSICA", "FISICA", "ANIMACIÓN", "ANIMACION", "LECTURA", "INTERDISCIPLINAR", "PROYECTO",
    "VOCACIONAL", "ORIENTACIÓN", "ORIENTACION", "PROMEDIO", "COMPORTAMIENTO",
    "OBSERVACIONES", "PROMOCIÓN", "PROMOCION", "CÍVICA", "CIVICA", "ACOMPAÑAMIENTO",
}



def _is_subject_header_cell(val_str: str, row_idx: int) -> bool:
    """Detecta si una celda corresponde al encabezado de materias para orientar el texto hacia arriba."""
    if not val_str or row_idx < 4 or row_idx > 15:
        return False
    upper = val_str.upper().strip()
    # Descartar columnas de nómina o metadatos de docentes
    if any(k in upper for k in ["NOMINA", "NÓMINA", "APELLIDOS", "ESTUDIANTE", "DOCENTE", "TUTOR", "CONSOLIDADO", "UNIDAD", "AÑO"]):
        return False
    return any(k in upper for k in SUBJECT_KEYWORDS)


def worksheet_to_html(ws: Worksheet, max_rows: int = 55, max_cols: int = 22) -> str:
    """
    Convierte una hoja de cálculo openpyxl en una tabla HTML interactiva
    con diseño idéntico a Excel: colores de fondo (#FFF2CC, #F1948A, colores pastel de materias),
    fuentes, celdas combinadas, texto vertical hacia arriba en materias y comentarios emergentes.
    """
    # Mapear celdas combinadas
    merged_spans = {}
    merged_hidden = set()

    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.min_col, rng.min_row, rng.max_col, rng.max_row
        rowspan = max_row - min_row + 1
        colspan = max_col - min_col + 1
        merged_spans[(min_row, min_col)] = (rowspan, colspan)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if (r, c) != (min_row, min_col):
                    merged_hidden.add((r, c))

    actual_max_row = min(ws.max_row, max_rows)
    actual_max_col = min(ws.max_column, max_cols)

    # Contenedor con scroll tipo Excel
    html = [
        '<div style="overflow: auto; max-height: 560px; border: 1px solid #94a3b8; '
        'border-radius: 8px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); background: #ffffff; margin-top: 8px;">'
        '<table style="border-collapse: collapse; font-family: Calibri, \'Segoe UI\', sans-serif; font-size: 11px; '
        'width: 100%; white-space: nowrap; color: #0f172a !important;">'
    ]

    # Encabezado de columnas de Excel (A, B, C...)
    html.append('<thead><tr style="background: #f1f5f9; color: #475569; position: sticky; top: 0; z-index: 10;">')
    html.append('<th style="border: 1px solid #cbd5e1; padding: 4px 6px; width: 35px; text-align: center; '
                'background: #e2e8f0; font-weight: bold; position: sticky; left: 0; z-index: 11; color: #475569;">#</th>')

    for c in range(1, actual_max_col + 1):
        col_letter = get_column_letter(c)
        html.append(f'<th style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: center; '
                    f'background: #f1f5f9; font-weight: bold; color: #475569;">{col_letter}</th>')
    html.append('</tr></thead><tbody>')

    for r in range(1, actual_max_row + 1):
        # Omitir filas vacías consecutivas hacia el final
        row_has_data = any(ws.cell(row=r, column=c).value is not None for c in range(1, actual_max_col + 1))
        if not row_has_data and r > 18:
            continue

        html.append('<tr>')
        # Número de fila (1, 2, 3...)
        html.append(f'<td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; '
                    f'background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0; '
                    f'background-color: #f1f5f9; z-index: 5;">{r}</td>')

        for c in range(1, actual_max_col + 1):
            if (r, c) in merged_hidden:
                continue

            cell = ws.cell(row=r, column=c)
            val = cell.value
            val_str = ""
            if val is not None:
                if isinstance(val, float):
                    val_str = f"{val:.2f}"
                else:
                    val_str = str(val)

            # Relleno de fondo (fill)
            bg_color = "#ffffff"
            if cell.fill and getattr(cell.fill, "fill_type", None) == "solid":
                color_obj = getattr(cell.fill, "start_color", None)
                if color_obj and color_obj.rgb:
                    rgb_str = str(color_obj.rgb)
                    if len(rgb_str) == 8 and rgb_str != "00000000":
                        bg_color = f"#{rgb_str[-6:]}"
                    elif len(rgb_str) == 6:
                        bg_color = f"#{rgb_str}"

            # Fuente (font)
            font_color = "#0f172a"
            is_bold = False
            if cell.font:
                if cell.font.bold:
                    is_bold = True
                if cell.font.color and cell.font.color.rgb:
                    fc = str(cell.font.color.rgb)
                    if len(fc) >= 6 and fc != "00000000":
                        font_color = f"#{fc[-6:]}"

            # Alineación
            align = "left"
            if cell.alignment and cell.alignment.horizontal:
                align = cell.alignment.horizontal

            # Comentario emergente con tooltip
            comment_badge = ""
            if cell.comment and cell.comment.text:
                comment_clean = str(cell.comment.text).replace('"', '&quot;').replace("\n", " &#10; ")
                comment_badge = (
                    f' <span title="{comment_clean}" style="cursor: help; color: #dc2626; '
                    f'font-size: 11px; margin-left: 2px;">💬</span>'
                )

            # Detectar si debe ir orientado hacia arriba
            is_vertical = False
            if cell.alignment and getattr(cell.alignment, "textRotation", 0) in (90, 255, -90, 180, 270):
                is_vertical = True
            elif _is_subject_header_cell(val_str, r):
                is_vertical = True

            font_size_css = ""
            if not is_vertical:
                if cell.font and getattr(cell.font, "size", None) and int(cell.font.size) >= 12:
                    font_size_css = f" font-size: {int(cell.font.size)}px;"
                elif r >= 8:
                    font_size_css = " font-size: 14px;"

            style = (
                f"border: 1px solid #cbd5e1; padding: 5px 7px; text-align: {align}; "
                f"background-color: {bg_color}; color: {font_color} !important;{font_size_css}"
            )
            if is_bold:
                style += " font-weight: bold;"

            if is_vertical:
                style += " height: 145px; vertical-align: bottom; text-align: center; padding: 6px 3px;"
                content = (
                    f'<div style="writing-mode: vertical-rl; transform: rotate(180deg); '
                    f'white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; '
                    f'display: inline-block; font-weight: bold; font-size: 11px; letter-spacing: 0.3px;">'
                    f'{val_str}</div>'
                )
            elif r == 1 and c == 1 and val_str:
                logo_b64 = _get_logo_base64()
                if logo_b64:
                    content = (
                        f'<div style="display: flex; align-items: center; justify-content: center; position: relative; min-height: 48px; padding: 4px;">'
                        f'<img src="data:image/png;base64,{logo_b64}" style="height: 48px; width: auto; object-fit: contain; position: absolute; left: 10px; top: 50%; transform: translateY(-50%);" alt="Logo">'
                        f'<span style="padding-left: 55px; text-align: center;">{val_str}{comment_badge}</span>'
                        f'</div>'
                    )
                else:
                    content = f"{val_str}{comment_badge}"
            else:
                content = f"{val_str}{comment_badge}"

            span_attr = ""
            if (r, c) in merged_spans:
                r_span, c_span = merged_spans[(r, c)]
                if r_span > 1:
                    span_attr += f' rowspan="{r_span}"'
                if c_span > 1:
                    span_attr += f' colspan="{c_span}"'

            html.append(f'<td style="{style}"{span_attr}>{content}</td>')

        html.append('</tr>')

    html.append('</tbody></table></div>')
    return "".join(html)


def get_preview_legend_html() -> str:
    """Devuelve la barra de leyenda de colores con alta legibilidad en modo claro y oscuro."""
    return """
    <div style="display: flex; flex-wrap: wrap; gap: 10px; margin: 10px 0 14px 0; font-size: 12px; font-family: 'Segoe UI', Roboto, sans-serif;">
      <div style="display: flex; align-items: center; gap: 8px; background: #ffffff !important; color: #0f172a !important; padding: 6px 14px; border-radius: 6px; border: 1px solid #cbd5e1; box-shadow: 0 1px 2px rgba(0,0,0,0.08);">
        <span style="display: inline-block; width: 14px; height: 14px; background-color: #FFF2CC; border: 1px solid #eab308; border-radius: 3px;"></span>
        <span style="color: #0f172a !important;"><strong style="color: #0f172a !important;">Estudiante con Discrepancia:</strong> Nómina resaltada en amarillo</span>
      </div>
      <div style="display: flex; align-items: center; gap: 8px; background: #ffffff !important; color: #0f172a !important; padding: 6px 14px; border-radius: 6px; border: 1px solid #cbd5e1; box-shadow: 0 1px 2px rgba(0,0,0,0.08);">
        <span style="display: inline-block; width: 14px; height: 14px; background-color: #F1948A; border: 1px solid #dc2626; border-radius: 3px;"></span>
        <span style="color: #0f172a !important;"><strong style="color: #0f172a !important;">Calificación con Error:</strong> Nota resaltada en rojo/coral</span>
      </div>
      <div style="display: flex; align-items: center; gap: 8px; background: #ffffff !important; color: #0f172a !important; padding: 6px 14px; border-radius: 6px; border: 1px solid #cbd5e1; box-shadow: 0 1px 2px rgba(0,0,0,0.08);">
        <span style="color: #dc2626; font-size: 13px;">💬</span>
        <span style="color: #0f172a !important;"><strong style="color: #0f172a !important;">Detalle Oficial:</strong> Pasa el ratón sobre 💬 para ver nota sistema y diferencia</span>
      </div>
      <div style="display: flex; align-items: center; gap: 8px; background: #ffffff !important; color: #0f172a !important; padding: 6px 14px; border-radius: 6px; border: 1px solid #cbd5e1; box-shadow: 0 1px 2px rgba(0,0,0,0.08);">
        <span style="display: inline-block; width: 14px; height: 14px; background-color: #93c5fd; border: 1px solid #3b82f6; border-radius: 3px;"></span>
        <span style="color: #0f172a !important;"><strong style="color: #0f172a !important;">Materias:</strong> Texto vertical hacia arriba con colores originales</span>
      </div>
    </div>
    """


def get_sample_preview_html() -> str:
    """
    Devuelve un ejemplo visual fiel e interactivo idéntico a las matrices reales de la institución:
    con membrete, datos del docente/tutor, jornada, materias con texto hacia arriba en colores pastel,
    y resaltados exactos de auditoría (amarillo nómina, rojo notas con error).
    """
    return """
    <div style="overflow: auto; max-height: 560px; border: 1px solid #94a3b8; border-radius: 8px; background: white; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); margin-top: 8px;">
    <table style="border-collapse: collapse; font-family: Calibri, 'Segoe UI', sans-serif; font-size: 11px; width: 100%; white-space: nowrap; color: #0f172a !important;">
      <thead>
        <tr style="background: #f1f5f9; color: #475569; position: sticky; top: 0; z-index: 10;">
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; background: #e2e8f0; width: 35px; text-align: center; position: sticky; left: 0; z-index: 11;">#</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">A</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">B</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">C</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">D</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">E</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">F</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">G</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">H</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">I</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">J</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">K</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">L</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">M</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">N</th>
          <th style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center;">O</th>
        </tr>
      </thead>
      <tbody>
        <!-- Membrete institucional con logotipo a la izquierda -->
        <tr>
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">1</td>
          <td colspan="15" style="border: 1px solid #cbd5e1; padding: 8px 14px; text-align: center; font-weight: bold; font-size: 13px; color: #0f172a; background: #ffffff;">
            """ + (
                f'<div style="display: flex; align-items: center; justify-content: center; position: relative; min-height: 48px; padding: 4px;">'
                f'<img src="data:image/png;base64,{_get_logo_base64()}" style="height: 48px; width: auto; object-fit: contain; position: absolute; left: 10px; top: 50%; transform: translateY(-50%);" alt="Logo">'
                f'<span style="padding-left: 55px; text-align: center;">UNIDAD EDUCATIVA &quot;ESTADOS UNIDOS DE NORTE AM&Eacute;RICA&quot;</span>'
                f'</div>'
                if _get_logo_base64() else
                'UNIDAD EDUCATIVA &quot;ESTADOS UNIDOS DE NORTE AM&Eacute;RICA&quot;'
            ) + """</td>
        </tr>

        <tr>
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">2</td>
          <td colspan="15" style="border: 1px solid #cbd5e1; padding: 3px; text-align: center; font-size: 11px; color: #334155; background: #ffffff;">Correo electr&oacute;nico: eeuuescuela@gmail.com &nbsp;&nbsp; eeuuvicerrectorado@gmail.com</td>
        </tr>
        <tr>
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">3</td>
          <td colspan="15" style="border: 1px solid #cbd5e1; padding: 3px; text-align: center; font-size: 11px; color: #334155; background: #ffffff;">Tel&eacute;fono: (07)289-0677 &nbsp;&nbsp; Distrito: 01D01 &nbsp;&nbsp; Zona: 6</td>
        </tr>
        <tr>
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">4</td>
          <td colspan="15" style="border: 1px solid #cbd5e1; padding: 4px; text-align: center; font-weight: bold; color: #0f172a; background: #ffffff;">A&Ntilde;O LECTIVO: 2025-2026</td>
        </tr>
        <!-- Datos de Docente y Jornada -->
        <tr>
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">5</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; background: #ffffff;"></td>
          <td colspan="6" style="border: 1px solid #cbd5e1; padding: 4px 8px; font-weight: bold; color: #0f172a; background: #ffffff;">Docente: Lcda. Marcia Siavichay</td>
          <td colspan="2" style="border: 1px solid #cbd5e1; padding: 4px; font-weight: bold; text-align: right; color: #0f172a; background: #ffffff;">JORNADA:</td>
          <td colspan="6" style="border: 1px solid #cbd5e1; padding: 4px 8px; font-weight: bold; color: #0f172a; background: #ffffff;">MATUTINA</td>
        </tr>
        <tr>
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">6</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; background: #ffffff;"></td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px; font-weight: bold; color: #0f172a; background: #ffffff;">A&Ntilde;O DE EGB:</td>
          <td colspan="5" style="border: 1px solid #cbd5e1; padding: 4px 8px; font-weight: bold; color: #0f172a; background: #ffffff;">OCTAVO A</td>
          <td colspan="8" style="border: 1px solid #cbd5e1; padding: 4px; background: #ffffff;"></td>
        </tr>
        <!-- Título Consolidado -->
        <tr>
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">7</td>
          <td colspan="15" style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; font-weight: bold; font-size: 12px; background: #FEF08A; color: #713f12;">CONSOLIDADO FINAL DE CALIFICACIONES</td>
        </tr>
        <!-- Encabezados de Tabla con Texto Hacia Arriba -->
        <tr style="height: 150px;">
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">8</td>
          <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; font-weight: bold; background: #D1D5DB; color: #1e293b; width: 35px; vertical-align: middle;">No</td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 12px; text-align: left; font-weight: bold; background: #D1D5DB; color: #1e293b; width: 280px; vertical-align: middle;">NOMINA</td>
          <!-- Materias verticales hacia arriba con sus colores originales -->
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #93C5FD; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">LENGUA Y LITERATURA</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #7DD3FC; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">MATEM&Aacute;TICA</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #FED7AA; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">CIENCIAS NATURALES</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #FBCFE8; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">ESTUDIOS SOCIALES</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #E2E8F0; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">INGL&Eacute;S</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #FDE047; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">CULTURAL Y ART&Iacute;STICA</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #67E8F9; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">EDUCACI&Oacute;N F&Iacute;SICA</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #A7F3D0; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">ANIMACI&Oacute;N A LA LECTURA</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #FEF08A; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">INTERDISCIPLINAR</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #BBF7D0; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">ORIENTACI&Oacute;N VOCACIONAL</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #86EFAC; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">PROMEDIO FINAL</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #D8B4FE; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">COMPORTAMIENTO</div>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 6px 4px; text-align: center; vertical-align: bottom; background: #CBD5E1; color: #0f172a;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; text-align: left; height: 135px; margin: 0 auto; display: inline-block; font-weight: bold;">OBSERVACIONES</div>
          </td>
        </tr>
        <!-- Estudiante 1 (Normal) -->
        <tr style="font-size: 14px;">
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">9</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">1</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px;">APOLO CACAY EMILY ANALIA</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.97</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.45</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.02</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.67</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.57</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.38</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">8.45</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A-</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center; font-weight: bold;">7.64</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">S</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: center;">PROMOVIDO</td>
        </tr>
        <!-- Estudiante 2 (CON DISCREPANCIAS RESALTADAS) -->
        <tr style="font-size: 14px;">
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">10</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">2</td>
          <!-- Resaltado nómina en amarillo -->
          <td style="border: 1px solid #eab308; padding: 4px 8px; background-color: #FFF2CC; font-weight: bold; color: #854d0e;">ARCE ORELLANA DARWIN FERNANDO</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.78</td>
          <!-- Resaltado nota en rojo coral con comentario -->
          <td style="border: 1px solid #dc2626; padding: 4px; text-align: center; background-color: #F1948A; color: #7f1d1d; font-weight: bold;">
            7.10 <span title="Docente: 7.10 &#10; Sistema: 9.94 &#10; Diferencia: -2.84" style="cursor: help; color: #dc2626;">💬</span>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.84</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.95</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.98</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.78</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">10.00</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <!-- Resaltado promedio en rojo coral con comentario -->
          <td style="border: 1px solid #dc2626; padding: 4px; text-align: center; background-color: #F1948A; color: #7f1d1d; font-weight: bold;">
            7.32 <span title="Docente: 7.32 &#10; Sistema: 9.89 &#10; Diferencia: -2.57" style="cursor: help; color: #dc2626;">💬</span>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">S</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: center;">PROMOVIDO</td>
        </tr>
        <!-- Estudiante 3 (Normal) -->
        <tr style="font-size: 14px;">
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">11</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">3</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px;">ARIAS CALLE ANTHONY ALEXANDER</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.37</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.10</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.00</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.00</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.00</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">7.67</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">8.12</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">B-</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A-</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center; font-weight: bold;">7.32</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">F</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: center;">PROMOVIDO</td>
        </tr>
        <!-- Estudiante 4 (CON DISCREPANCIAS RESALTADAS) -->
        <tr style="font-size: 14px;">
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">12</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">4</td>
          <!-- Resaltado nómina en amarillo -->
          <td style="border: 1px solid #eab308; padding: 4px 8px; background-color: #FFF2CC; font-weight: bold; color: #854d0e;">ATANCURI LOPEZ CHELSEA ABIGAIL</td>
          <!-- Resaltado nota en rojo coral con comentario -->
          <td style="border: 1px solid #dc2626; padding: 4px; text-align: center; background-color: #F1948A; color: #7f1d1d; font-weight: bold;">
            8.50 <span title="Docente: 8.50 &#10; Sistema: 9.95 &#10; Diferencia: -1.45" style="cursor: help; color: #dc2626;">💬</span>
          </td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.60</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.50</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.59</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.68</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.74</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.76</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center; font-weight: bold;">9.68</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">S</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: center;">PROMOVIDO</td>
        </tr>
        <!-- Estudiante 5 (Normal) -->
        <tr style="font-size: 14px;">
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">13</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">5</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px;">AVILA CORTE MARIA GRACIA</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.79</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.35</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.50</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.76</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.50</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.76</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.77</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A-</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center; font-weight: bold;">9.63</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">S</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: center;">PROMOVIDO</td>
        </tr>
        <!-- Estudiante 6 (Normal) -->
        <tr style="font-size: 14px;">
          <td style="border: 1px solid #cbd5e1; padding: 3px 6px; text-align: center; background: #f8fafc; font-weight: bold; color: #64748b; position: sticky; left: 0;">14</td>

          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">6</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px;">BARROS CARDENAS DAVID ALEXANDER</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">8.99</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.07</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">8.13</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">8.56</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">8.90</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">9.11</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">8.54</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A-</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">A+</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center; font-weight: bold;">8.75</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center;">S</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px 8px; text-align: center;">PROMOVIDO</td>
        </tr>
      </tbody>
    </table>
    </div>
    """
