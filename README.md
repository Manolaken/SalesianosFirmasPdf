# Aplicación para Insertar Firmas en PDFs

Esta aplicación de escritorio permite automatizar la inserción de firmas en archivos PDF, buscando campos resaltados y añadiendo las firmas correspondientes de alumnos y profesores.

## Características

- **Interfaz gráfica amigable** con tkinter
- **Detección automática** de campos resaltados:
  - Amarillo: Teléfono o nombre del alumno
  - Rojo/Naranja: Nombre del profesor
- **Inserción inteligente de firmas** justo encima de los nombres
- **Soporte para múltiples PDFs** en batch
- **Descarga automática de firmas** desde URLs o rutas locales
- **Log detallado** de todas las operaciones
- **Cache de firmas** para optimizar el procesamiento
- **Búsqueda flexible** con coincidencias parciales

## Requisitos previos

- Python 3.8 o superior
- Sistema operativo: Windows, macOS o Linux

## Instalación

### 1. Instalar Python

Si no tienes Python instalado, descárgalo desde [python.org](https://www.python.org/downloads/)

### 2. Crear entorno virtual (recomendado)

```bash
# En Windows
python -m venv venv
venv\Scripts\activate

# En macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

## Uso

### Ejecutar la aplicación

```bash
python insertar_firmas_pdf.py
```

### Pasos para procesar PDFs

1. **Añadir PDFs**: Haz clic en "📄 Añadir PDFs" y selecciona uno o más archivos PDF
2. **Seleccionar Excel**: Haz clic en "📊 Seleccionar Excel" y selecciona el archivo con las firmas
3. **(Opcional) Carpeta de salida**: Haz clic en "📁 Seleccionar Carpeta" si quieres guardar los PDFs en una ubicación específica
4. **Ajustar configuración**: Modifica el tamaño y margen de las firmas si es necesario
5. **Procesar**: Haz clic en "▶ Procesar PDFs"

## Formato del archivo Excel

El archivo Excel debe tener exactamente estas columnas:

| Columna | Descripción | Ejemplo |
|---------|-------------|---------|
| **id** | Identificador único | 1, 2, 3... |
| **nombre** | Nombre de la persona | ivan, mariano |
| **apellido1** | Primer apellido | coma, alcalde |
| **apellido2** | Segundo apellido | lopez, gracia |
| **telefono** | Teléfono (solo para alumnos) | 659645517 o NaN |
| **ruta** | Ruta completa al archivo de firma | C:\...\firma.jpg |

### Reglas importantes:

1. **Si tiene teléfono** → Se considera ALUMNO
2. **Si NO tiene teléfono (NaN/vacío)** → Se considera PROFESOR
3. El nombre completo se construye automáticamente: `NOMBRE APELLIDO1 APELLIDO2`
4. Las rutas pueden ser absolutas o relativas al Excel

### Ejemplo de datos:

```
id | nombre  | apellido1 | apellido2 | telefono   | ruta
---|---------|-----------|-----------|------------|---------------------------
1  | ivan    | coma      | lopez     | 659645517  | firmas/ivancomalopez.jpg
2  | mariano | alcalde   | gracia    | NaN        | firmas/marianoalcalde.jpg
```

## Formato de las rutas de firmas

La aplicación soporta tres tipos de rutas:

1. **Rutas absolutas**: `C:\Users\usuario\firmas\firma.png` (Windows) o `/home/usuario/firmas/firma.png` (Linux/Mac)
2. **Rutas relativas al Excel**: `firmas/firma.png` (busca en la carpeta "firmas" junto al Excel)
3. **URLs**: `https://example.com/firmas/firma.png`

## Configuración

### Ajustar tamaño de las firmas

En la sección "Configuración" puedes ajustar:
- **Ancho de firma**: 50-300 píxeles (predeterminado: 120)
- **Alto de firma**: 30-200 píxeles (predeterminado: 60)
- **Margen superior**: 0-50 píxeles (predeterminado: 5)

## Requisitos del PDF

Los PDFs deben tener:
- **Campos resaltados en amarillo**: Para teléfono o nombre del alumno
- **Campos resaltados en rojo/naranja**: Para nombre del profesor

La aplicación detecta automáticamente estos campos y coloca las firmas justo encima.

## Salida

Los PDFs procesados se guardan con el sufijo `_firmado.pdf`:
- Original: `documento.pdf`
- Procesado: `documento_firmado.pdf`

## Solución de problemas

### No se detectan los campos resaltados
- Verifica que los campos estén resaltados correctamente (amarillo para alumnos, rojo/naranja para profesores)
- Los resaltados deben ser "anotaciones" de tipo highlight, no solo fondo de color

### No se encuentran las firmas
- Verifica que las rutas en el Excel sean correctas
- Si usas rutas relativas, deben estar en relación a la ubicación del archivo Excel
- Para URLs, verifica que sean accesibles

### Error al cargar el Excel
- Verifica que el archivo sea .xlsx o .xls
- Verifica que las columnas tengan nombres descriptivos

## Mejoras futuras

- [ ] Soporte para Google Drive
- [ ] Soporte para otras nubes (OneDrive, Dropbox)
- [ ] Generación automática de Excel desde PDFs
- [ ] Previsualización antes de procesar

## Autor

Aplicación desarrollada para automatizar el proceso de firma de documentos de planes de formación.

## Licencia

Este software es de uso interno y educativo.
