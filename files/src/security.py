import io
import zipfile
import openpyxl
from streamlit.runtime.uploaded_file_manager import UploadedFile

MAX_FILE_SIZE_MB = 30
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

def validate_excel_file(uploaded_file: UploadedFile) -> tuple[bool, str]:
    """
    Implementa 3 capas de seguridad para asegurar que el archivo subido sea un Excel válido.
    Devuelve (True, "") si es seguro.
    Devuelve (False, "mensaje de error amigable") si falla en alguna capa.
    """
    
    # -------------------------------------------------------------
    # CAPA 1: Validación de Memoria y Tamaño
    # -------------------------------------------------------------
    if uploaded_file.size == 0:
        return False, "El archivo está completamente vacío (0 bytes)."
        
    if uploaded_file.size > MAX_FILE_SIZE_BYTES:
        return False, f"El archivo excede el límite de tamaño permitido ({MAX_FILE_SIZE_MB} MB). Por favor, reduce el tamaño del archivo."

    # -------------------------------------------------------------
    # CAPA 2: Validación de Firma Digital (Magic Bytes)
    # -------------------------------------------------------------
    # Los archivos .xlsx son en realidad archivos comprimidos ZIP.
    # Todo ZIP válido comienza con los bytes mágicos: PK\x03\x04 (50 4B 03 04)
    # Leemos los primeros 4 bytes sin consumir mucha memoria.
    uploaded_file.seek(0)
    magic_bytes = uploaded_file.read(4)
    
    if magic_bytes != b'PK\x03\x04':
        return False, "Formato inválido. El archivo parece no ser un Excel genuino (¿quizás es una foto o texto renombrado a .xlsx?)."

    # -------------------------------------------------------------
    # CAPA 3: Validación Estructural (OpenPyXL Test Ligero)
    # -------------------------------------------------------------
    # Intentamos abrir la estructura del archivo. Si está corrupto, lanzará error.
    try:
        uploaded_file.seek(0)
        # Usamos BytesIO para no bloquear el archivo físico
        file_stream = io.BytesIO(uploaded_file.read())
        
        # Leemos solo la estructura básica usando read_only=True para que sea instantáneo y no consuma RAM
        wb = openpyxl.load_workbook(file_stream, read_only=True)
        wb.close()
    except zipfile.BadZipFile:
        return False, "El archivo está corrupto internamente (Error de compresión ZIP). Por favor, intenta guardarlo nuevamente desde Excel."
    except Exception as e:
        return False, f"Error desconocido al leer la estructura del Excel: {str(e)}"
        
    # Restaurar el puntero para que la aplicación pueda usar el archivo normalmente después de la validación
    uploaded_file.seek(0)
    return True, ""
