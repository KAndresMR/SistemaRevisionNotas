# 📊 Sistema de Auditoría de Calificaciones (Dashboard Web)

Bienvenido al repositorio del **Sistema de Auditoría de Calificaciones**, una herramienta automatizada diseñada para comparar las notas ingresadas por los docentes en sus matrices de Excel frente a los reportes oficiales generados por el sistema académico, detectando discrepancias y errores con precisión milimétrica.

## 🚀 Arquitectura del Proyecto

El sistema ha sido refactorizado para operar bajo una arquitectura limpia y centralizada en la web mediante **Streamlit**. Todo el código fuente reside en el directorio `src/`.

### Estructura de Directorios
```text
SistemaRevisionNotas/
├── README.md               # Este documento
├── .venv/                  # Entorno virtual de Python
└── src/                    # Código fuente de la aplicación
    ├── app.py              # Interfaz gráfica de usuario (Streamlit)
    ├── auditor.py          # Lógica CORE: normalización de nombres, emparejamiento y comparación
    ├── config.py           # Variables globales, diccionarios de cursos y alias de materias
    ├── engine.py           # Orquestador del flujo de lectura, comparación y escritura
    ├── excel_reader.py     # Lógica de extracción de datos puros de Excel (openpyxl)
    └── teacher_highlighter.py # Lógica de formato visual para reportar errores en Excel
```

## 🛠️ Tecnologías y Librerías Clave

- **Python 3.10+**
- **Streamlit (`import streamlit as st`)**: Utilizado en `app.py` para construir un dashboard web responsivo, donde los usuarios cargan múltiples archivos simultáneamente mediante componentes drag-and-drop (`st.file_uploader`), sin necesidad de usar la consola.
- **OpenPyXL (`import openpyxl`)**: El motor subyacente (`excel_reader.py` y `teacher_highlighter.py`) que nos permite leer valores y fórmulas (usando `data_only=False`), y modificar colores de celdas y comentarios (`PatternFill`, `Comment`) sin dañar el formato original del documento.
- **Pandas (`import pandas as pd`)**: Utilizado superficialmente en `app.py` para inyectar datos tubulares en el Dashboard.

---

## 🧠 Lógica de Comparación: ¿Cómo funciona por debajo?

El flujo de ejecución (`src/engine.py`) orquesta cuatro pasos fundamentales:

### 1. Extracción de Datos (`excel_reader.py`)
El sistema lee dinámicamente los Excel buscando palabras clave definidas en `src/config.py` (ej. `"NÓMINA"`, `"ESTUDIANTE"`) para identificar la fila de encabezados y mapear las columnas correctas de cada materia. Retorna listas de objetos `Student` (definidos en `auditor.py`).

### 2. Emparejamiento de Estudiantes (`auditor.py -> match_students`)
Compara el listado de docentes contra el listado del sistema. 
- Utiliza **Normalización de Nombres** (`normalize_name`): Convierte todo a mayúsculas, elimina tildes, signos de puntuación y espacios extras.
- Aplica una heurística de **Match Exacto** y luego una de **Tolerancia a Nombres Truncados** (ej. "JUAN PEREZ" emparejará con "JUAN ALEJANDRO PEREZ" si la fila coincide lógicamente).
- Clasifica a los estudiantes en tres estados: `PERFECT_MATCH`, `ONLY_IN_TEACHERS`, `ONLY_IN_SYSTEM`.

### 3. Comparación de Calificaciones (`auditor.py -> compare_students`)
Por cada estudiante emparejado, se iteran las materias definidas en `config.py` (ej. Matemáticas, Lenguaje, etc.).
- Las celdas en blanco (`None`) o con texto se evalúan contra un comportamiento estándar.
- Se aplica una tolerancia numérica (definida en `config.TOLERANCE = 0.02`) para evitar falsos positivos por redondeo.
- Si `abs(nota_docente - nota_sistema) > TOLERANCE`, se emite una alerta (`GradeStatus.ERROR`).

### 4. Resaltado Visual y Reporte Ejecutivo (`teacher_highlighter.py`)
En lugar de crear archivos de texto planos, el sistema inyecta colores directamente sobre el Excel original que subió el profesor.
- Colorea la celda exacta de la nota con error de color **Rojo oscuro** (`#991B1B`).
- Crea la hoja interactiva **"RESUMEN DISCREPANCIAS"**, que agrupa todos los fallos por estudiante utilizando *Dictionaries* en memoria, ahorrando espacio horizontal y entregando un reporte sumamente profesional.

---

## 💻 Guía de Inicio Rápido (Para Desarrolladores)

Si deseas probar el código en tu entorno local o realizar modificaciones a la lógica de negocio, sigue estos pasos:

1. **Activar el Entorno Virtual**
   Abre una terminal en la raíz del proyecto y ejecuta:
   ```bash
   source .venv/bin/activate
   ```

2. **Levantar el Dashboard**
   Navega a la carpeta `src/` y lanza el servidor de Streamlit:
   ```bash
   cd src/
   streamlit run app.py
   ```
   *Esto abrirá automáticamente `http://localhost:8501` en tu navegador.*

3. **Prueba Práctica**
   - En la sección **"Matriz de Docentes"**, arrastra el consolidado maestro (`BS MAT ANUAL MATRIZ TUTOR.xlsx`).
   - En la sección **"Reportes Oficiales"**, sube uno o varios reportes de Carmenta (ej. `8AM.xlsx`, `9AM.xlsx`).
   - Haz clic en **"Ejecutar Auditoría"**. Verás cómo procesa los archivos en memoria y te devuelve un único archivo `Auditoria_Matriz_Docentes_Resultados.xlsx` descargable y perfectamente marcado.

---
*Desarrollado para optimizar tiempos administrativos académicos garantizando fiabilidad de datos.*
