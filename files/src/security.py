import io
import os
import re
from pathlib import Path
import zipfile
import openpyxl
from streamlit.runtime.uploaded_file_manager import UploadedFile

MAX_FILE_SIZE_MB = 30
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

# Límites defensivos contra ataques DoS / Zip Bomb
MAX_UNCOMPRESSED_MB = 120
MAX_UNCOMPRESSED_BYTES = MAX_UNCOMPRESSED_MB * 1024 * 1024
MAX_ZIP_ENTRIES = 2500
MAX_COMPRESSION_RATIO = 100.0


def sanitize_filename(raw_name: str, fallback_prefix: str = "archivo") -> str:
    """
    Sanitiza nombres de archivos para prevenir ataques de Path Traversal (Zip Slip / Directory Traversal).
    Elimina secuencias peligrosas (../, ..\\, bytes nulos) y conserva solo caracteres seguros.
    """
    if not raw_name or not isinstance(raw_name, str):
        return f"{fallback_prefix}.xlsx"

    # 1. Extraer solo el nombre base descartando cualquier ruta relativa o absoluta
    base_name = os.path.basename(raw_name.replace("\\", "/"))
    # Eliminar bytes nulos
    base_name = base_name.replace("\x00", "").strip()

    # 2. Descartar puntos iniciales (evitar archivos ocultos o '../')
    base_name = re.sub(r"^\.+", "", base_name)

    # 3. Filtrar caracteres no permitidos en sistemas de archivos
    safe_name = re.sub(r'[^a-zA-Z0-9_\-\.\sáéíóúÁÉÍÓÚñÑ]', '_', base_name)
    safe_name = re.sub(r'\s+', ' ', safe_name).strip()

    # 4. Asegurar que termine en .xlsx
    if not safe_name.lower().endswith(".xlsx"):
        safe_name = f"{safe_name}.xlsx"

    # Truncar nombres excesivamente largos (máx 120 caracteres)
    if len(safe_name) > 120:
        stem = Path(safe_name).stem[:110]
        safe_name = f"{stem}.xlsx"

    return safe_name if safe_name != ".xlsx" else f"{fallback_prefix}.xlsx"


def validate_excel_file(uploaded_file: UploadedFile) -> tuple[bool, str]:
    """
    Implementa 5 capas de seguridad para asegurar que el archivo subido sea un Excel .xlsx válido y seguro:
    1. Tamaño y memoria del archivo.
    2. Nombre de archivo y extensión .xlsx.
    3. Firma digital de bytes mágicos (PK\x03\x04).
    4. Inspección profunda del archivo ZIP (Prevención de Zip Bomb y macros maliciosas).
    5. Validación estructural ligera con OpenPyXL.

    Devuelve (True, "") si es seguro.
    Devuelve (False, "mensaje de error amigable") si falla en alguna capa.
    """
    # -------------------------------------------------------------
    # CAPA 1: Validación de Memoria y Tamaño
    # -------------------------------------------------------------
    file_size = getattr(uploaded_file, "size", 0)
    if file_size == 0:
        return False, "El archivo está completamente vacío (0 bytes)."

    if file_size > MAX_FILE_SIZE_BYTES:
        return False, f"El archivo excede el límite de tamaño permitido ({MAX_FILE_SIZE_MB} MB). Por favor, reduce el tamaño del archivo."

    # -------------------------------------------------------------
    # CAPA 2: Validación de Extensión y Nombre
    # -------------------------------------------------------------
    file_name = getattr(uploaded_file, "name", "")
    if not file_name or not file_name.lower().endswith(".xlsx"):
        return False, "Extensión no permitida. Solo se admiten libros de Excel en formato moderno (.xlsx)."

    # -------------------------------------------------------------
    # CAPA 3: Validación de Firma Digital (Magic Bytes)
    # -------------------------------------------------------------
    uploaded_file.seek(0)
    magic_bytes = uploaded_file.read(4)

    if magic_bytes != b'PK\x03\x04':
        return False, "Formato inválido. El archivo no posee una cabecera de libro Excel válida (¿quizás es una imagen o PDF renombrado a .xlsx?)."

    # -------------------------------------------------------------
    # CAPA 4: Inspección de Archivo ZIP (Prevención Zip Bomb y Macros)
    # -------------------------------------------------------------
    try:
        uploaded_file.seek(0)
        file_bytes = uploaded_file.read()
        file_stream = io.BytesIO(file_bytes)

        with zipfile.ZipFile(file_stream, "r") as zf:
            entries = zf.infolist()
            if len(entries) > MAX_ZIP_ENTRIES:
                return False, f"El archivo contiene una cantidad inusualmente alta de elementos internos ({len(entries)}). Rechazado por seguridad."

            total_uncompressed = 0
            for entry in entries:
                # Detección de macros incrustadas (vbaProject.bin)
                entry_name_lower = entry.filename.lower()
                if "vbaproject.bin" in entry_name_lower or entry_name_lower.endswith((".exe", ".dll", ".bat", ".cmd", ".vbs")):
                    return False, "El archivo contiene código ejecutable o macros no permitidas (.bin/.exe). Por seguridad solo se admiten hojas de cálculo estándar."

                total_uncompressed += entry.file_size
                if total_uncompressed > MAX_UNCOMPRESSED_BYTES:
                    return False, f"El contenido descomprimido del archivo supera el límite de seguridad ({MAX_UNCOMPRESSED_MB} MB). Posible archivo malicioso (Zip Bomb)."

            # Ratio de compresión desproporcionado
            if file_size > 0:
                ratio = total_uncompressed / file_size
                if ratio > MAX_COMPRESSION_RATIO and total_uncompressed > 10 * 1024 * 1024:
                    return False, f"Tasa de compresión sospechosa ({ratio:.1f}:1). Rechazado preventivamente para proteger el servidor."

    except zipfile.BadZipFile:
        return False, "El archivo está corrupto internamente (Error de compresión ZIP). Por favor, abre el archivo en Excel y guárdalo de nuevo."
    except Exception as e:
        return False, f"Error al verificar la integridad del archivo: {str(e)}"

    # -------------------------------------------------------------
    # CAPA 5: Validación Estructural (OpenPyXL Test Ligero)
    # -------------------------------------------------------------
    try:
        file_stream.seek(0)
        wb = openpyxl.load_workbook(file_stream, read_only=True)
        # Verificar que tenga al menos una hoja
        if not wb.sheetnames:
            wb.close()
            return False, "El libro de Excel no contiene ninguna hoja de cálculo."
        wb.close()
    except Exception as e:
        return False, f"Estructura de Excel ilegible o dañada: {str(e)}"

    # Restaurar el puntero para que la aplicación pueda usar el archivo normalmente
    uploaded_file.seek(0)
    return True, ""

