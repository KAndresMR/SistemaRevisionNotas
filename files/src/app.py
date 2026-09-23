import streamlit as st
import os
import shutil
import tempfile
from pathlib import Path
import zipfile
import pandas as pd

from collections import defaultdict

from config import FULL_MAP
from engine import run_audit
from auditor import GradeStatus, StudentMatchStatus
from teacher_highlighter import create_or_update_discrepancies_sheet
from security import validate_excel_file
from excel_reader import ExcelStructureError

st.set_page_config(page_title="Auditoría de Notas", page_icon="📊", layout="wide")

# Diseño Premium Básico para Streamlit
st.markdown("""
<style>
    .main { background-color: #f4f6f9; }
    h1, h2, h3 { color: #1e3a8a; font-family: 'Inter', sans-serif; }
    .stButton>button { background-color: #2563eb; color: white; border-radius: 8px; border: none; padding: 10px 24px; font-weight: bold; transition: 0.3s; }
    .stButton>button:hover { background-color: #1d4ed8; color: white; transform: scale(1.02); }
    .stDownloadButton>button { background-color: #10b981; color: white; border-radius: 8px; border: none; font-weight: bold; }
    .stDownloadButton>button:hover { background-color: #059669; color: white; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Dashboard de Revisión de Notas")
st.markdown("Sube tu matriz de docentes y los reportes oficiales para generar la auditoría automática al instante.")

# Diccionario combinado para búsqueda rápida

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Docente(s)")
    teacher_files = st.file_uploader("Sube uno o varios archivos (ej. notas_8AM.xlsx o Matriz Consolidada)", type=['xlsx'], accept_multiple_files=True)
    
    jornada = st.selectbox("Selecciona la Jornada", ["Matutina", "Vespertina"])

with col2:
    st.subheader("2. Reportes Oficiales")
    report_files = st.file_uploader("Sube uno o varios reportes (ej. 8AM.xlsx, 9AM.xlsx)", type=['xlsx'], accept_multiple_files=True)

if st.button("🚀 Ejecutar Auditoría") and teacher_files and report_files:
    # ---------------------------------------------------------
    # VALIDACIÓN DE SEGURIDAD Y FORMATO
    # ---------------------------------------------------------
    for t_file in teacher_files:
        is_valid, error_msg = validate_excel_file(t_file)
        if not is_valid:
            st.error(f"⚠️ Archivo Docente ('{t_file.name}') inválido: {error_msg}")
            st.stop()
            
    for report in report_files:
        is_valid, error_msg = validate_excel_file(report)
        if not is_valid:
            st.error(f"⚠️ Reporte Oficial ('{report.name}') inválido: {error_msg}")
            st.stop()

    def guess_course_from_filename(filename: str, fmap: dict) -> str | None:
        norm_name = filename.upper().replace(" ", "").replace("_", "").replace("-", "")
        # Buscar de mayor a menor longitud para evitar que "8A" intercepte a "8AM" si existiera
        sorted_keys = sorted(fmap.keys(), key=len, reverse=True)
        for key in sorted_keys:
            if key.upper() in norm_name:
                return key
        return None

    with st.spinner("Procesando archivos..."):
        import gc
        import io
        # Crear directorio temporal
        temp_dir = Path(tempfile.mkdtemp())
        download_data = None
        
        try:
            # 1. Emparejamiento Inteligente (Smart Match)
            teacher_map = {}
            for t_file in teacher_files:
                c_key = guess_course_from_filename(t_file.name, FULL_MAP)
                if c_key:
                    teacher_map[c_key] = t_file
                else:
                    # Si no encuentra curso en el nombre, asume que es la matriz consolidada
                    teacher_map["CONSOLIDATED"] = t_file
                    
            all_discrepancies = []
            file_discrepancies = defaultdict(list)
            course_summary = {}
            processed_courses = []
            
            progress_bar = st.progress(0)
            
            # Directorio temporal interno para los archivos marcados individuales
            marked_dir = temp_dir / "resultados_marcados"
            marked_dir.mkdir(exist_ok=True)
            
            for idx, report in enumerate(report_files):
                course_key = guess_course_from_filename(report.name, FULL_MAP)
                
                if not course_key:
                    st.warning(f"⚠️ El archivo '{report.name}' tiene un nombre irreconocible. Se omite.")
                    continue
                    
                sheet_name, course_name = FULL_MAP[course_key]
                
                # 2. Cargas Asimétricas: Buscar pareja
                matched_t_file = teacher_map.get(course_key) or teacher_map.get("CONSOLIDATED")
                
                if not matched_t_file:
                    st.warning(f"⚠️ Falta archivo de docente para '{course_key}'. Omitiendo este curso.")
                    continue
                
                # Guardar archivos emparejados en subdirectorios separados
                # para evitar colisiones de nombre (ej. ambos se llaman 4AM.xlsx)
                teacher_subdir = temp_dir / "docentes"
                teacher_subdir.mkdir(exist_ok=True)
                teacher_path = teacher_subdir / matched_t_file.name
                if not teacher_path.exists():
                    with open(teacher_path, "wb") as f:
                        f.write(matched_t_file.getbuffer())
                
                official_subdir = temp_dir / "oficiales"
                official_subdir.mkdir(exist_ok=True)
                report_path = official_subdir / report.name
                with open(report_path, "wb") as f:
                    f.write(report.getbuffer())
                    
                output_file = temp_dir / f"auditoria_{course_key}.xlsx"
                # Nombre del archivo de salida basado en el archivo original del docente
                marked_teacher_output = marked_dir / f"Matriz_Marcada_{matched_t_file.name}"
                
                try:
                    marked_path, audit_rows = run_audit(
                        teacher_file=str(teacher_path),
                        official_file=str(report_path),
                        output_file=str(output_file),
                        course_name=course_name,
                        teacher_sheet=sheet_name,
                        teacher_marked_output=str(marked_teacher_output),
                        reset_marked_teacher=False, # <-- Falso para que acumule las hojas
                        template_file=str(teacher_path),
                    )
                    
                    # Métricas
                    n_err = sum(1 for r in audit_rows if r.has_error)
                    n_disc = sum(sum(1 for c in r.comparisons if c.status == GradeStatus.ERROR) for r in audit_rows)
                    n_desf = sum(1 for r in audit_rows if r.has_desfase)
                    
                    course_summary[course_key] = {
                        "total_students": len(audit_rows),
                        "students_with_error": n_err,
                        "total_discrepancies": n_disc,
                        "students_with_desfase": n_desf,
                    }
                    
                    for audit_row in audit_rows:
                        if not audit_row.comparisons:
                            detail = "SOLO EN DOCENTES" if audit_row.match_status == StudentMatchStatus.ONLY_IN_TEACHERS else "SOLO EN SISTEMA"
                            disc = {
                                "curso": course_key, "estudiante": audit_row.student_name,
                                "materia": detail, "docente": "-", "sistema": "-", "diferencia": "-", "estado": "DESFASE DE NÓMINA"
                            }
                            all_discrepancies.append(disc)
                            file_discrepancies[str(marked_teacher_output)].append(disc)
                        else:
                            if audit_row.match_status == StudentMatchStatus.MATCHED_BY_NAME:
                                disc = {
                                    "curso": course_key, "estudiante": audit_row.student_name,
                                    "materia": "(NÓMINA)", "docente": "-", "sistema": "-", "diferencia": "-", "estado": "DESFASE DE NÓMINA"
                                }
                                all_discrepancies.append(disc)
                                file_discrepancies[str(marked_teacher_output)].append(disc)
                            for comp in audit_row.comparisons:
                                if comp.status == GradeStatus.ERROR:
                                    disc = {
                                        "curso": course_key, "estudiante": audit_row.student_name,
                                        "materia": comp.subject_display, "docente": comp.teacher_grade,
                                        "sistema": comp.official_grade, "diferencia": comp.difference, "estado": "ERROR"
                                    }
                                    all_discrepancies.append(disc)
                                    file_discrepancies[str(marked_teacher_output)].append(disc)
                    
                    processed_courses.append(course_key)
                    
                except ExcelStructureError as e:
                    st.warning(f"⚠️ El archivo de {course_key} tiene errores estructurales (ej. falta columna de NÓMINA): {str(e)}. Se omitirá.")
                except Exception as e:
                    st.error(f"Error procesando {course_key}: {e}")
                    
                progress_bar.progress((idx + 1) / len(report_files))
                
            # Post-proceso: Insertar hoja de discrepancias dentro de cada archivo marcado
            if file_discrepancies:
                for f_path, discrepancies in file_discrepancies.items():
                    create_or_update_discrepancies_sheet(f_path, discrepancies)
            
            # 3. Empaquetar resultados en un ZIP volátil
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                # Agregar todos los archivos docentes marcados
                for f in marked_dir.glob("*.xlsx"):
                    zip_file.write(f, arcname=f.name)
            
            download_data = zip_buffer.getvalue()
            
            st.success("✅ ¡Auditoría completada exitosamente!")
            
            # Mostrar resumen en UI
            st.subheader("Resumen de Errores")
            if all_discrepancies:
                df_errores = pd.DataFrame(all_discrepancies)
                st.dataframe(df_errores, use_container_width=True)
                
                # Gráfica rápida
                errores_por_curso = df_errores['curso'].value_counts()
                st.bar_chart(errores_por_curso)
            else:
                st.info("¡Felicidades! No se encontraron discrepancias en los cursos procesados.")

        finally:
            # LIMPIEZA EXTREMA: Borrar todos los archivos físicos inmediatamente
            shutil.rmtree(temp_dir, ignore_errors=True)
            # Limpiar RAM de los diccionarios pesados de openpyxl
            gc.collect()

    # Renderizar el botón de descarga (ZIP) usando la RAM volátil
    if download_data:
        st.download_button(
            label="📥 Descargar Paquete de Resultados (ZIP)",
            data=download_data,
            file_name="Resultados_Auditoria_Notas.zip",
            mime="application/zip"
        )
