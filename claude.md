# # Unificador de Planillas Excel

## Contexto del proyecto
Flask-based Excel unifier tool. Permite subir múltiples archivos Excel/CSV,
seleccionar hojas y columnas, revisar problemas de datos y descargar un 
archivo unificado.

## Stack
- **Backend:** Flask 2.x, Pandas, OpenPyXL
- **Frontend:** HTML/CSS/JS vanilla (sin framework)
- **Puerto local:** 5000

## Archivos clave
- `app.py` — backend completo (rutas, lógica de merge, análisis)
- `templates/index.html` — frontend completo (UI, JS)

## Comandos
- `python app.py` — iniciar servidor de desarrollo
- `pip install -r requirements.txt` — instalar dependencias

## Features actuales
- Subida múltiple de archivos (.xlsx, .xls, .csv)
- Selección de hoja por archivo
- Selección y reordenamiento de columnas (drag & drop)
- Archivo prioritario para resolución de duplicados
- Revisión de valores similares (homologación)
- Revisión de duplicados (con selección manual)
- Revisión de problemas de mayúsculas/minúsculas
- Confirmación y descarga del Excel unificado

## Rutas API
- `POST /upload` — carga archivos, retorna columnas y estilos
- `POST /preview` — analiza datos, retorna similitudes/duplicados/casing
- `POST /confirm` — aplica correcciones y genera Excel
- `POST /merge` — fast path (backward compatibility)

## Convenciones
- Comentarios en español
- Mantener consistencia visual entre secciones sim/dup/cas
- Plan primero, código después en cambios grandes