# Generador de Contratos · Filadd

Herramienta interna para generar contratos en masa a partir de una plantilla Word y un Excel con datos.

## Archivos

```
app.py                  → aplicación principal
requirements.txt        → dependencias Python
packages.txt            → dependencias sistema (LibreOffice para PDF)
.streamlit/config.toml  → configuración de tema y servidor
```

## Deploy en Streamlit Cloud

1. Subir esta carpeta a un repositorio GitHub (puede ser privado)
2. Ir a [share.streamlit.io](https://share.streamlit.io)
3. Conectar el repo y seleccionar `app.py` como entry point
4. Deploy — Streamlit instala todo automáticamente

> `packages.txt` instala LibreOffice en el servidor para exportar PDFs.
> Sin él, los contratos se generan en `.docx`.

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Variables soportadas en plantillas

Cualquier texto entre doble llave: `{{NOMBRE_VARIABLE}}`
- Solo letras, números y guión bajo
- La columna del Excel debe tener el mismo nombre (sin llaves)
- Funciona en cuerpo, tablas, encabezados y pies de página
