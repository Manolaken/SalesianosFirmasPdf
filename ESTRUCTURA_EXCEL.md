# Estructura del Excel LibroFirmas.xlsx

## Columnas del Excel

El archivo Excel debe tener exactamente esta estructura:

| id | nombre | apellido1 | apellido2 | telefono | ruta |
|----|--------|-----------|-----------|----------|------|
| 1  | manuel | peiro     | perez     | 661825400 | C:\...\manuelpeiroperez.jpg |
| 2  | mariano| alcalde   | gracia    | (vacío)   | C:\...\marianoalcaldegracia.jpg |
| 3  | ivan   | coma      | lopez     | 659645517 | C:\...\ivancomalopez.jpg |

### Descripción de columnas:

- **id**: Identificador único (numérico)
- **nombre**: Nombre de la persona (texto)
- **apellido1**: Primer apellido (texto)
- **apellido2**: Segundo apellido (texto)
- **telefono**: Número de teléfono (numérico, puede estar vacío)
- **ruta**: Ruta completa al archivo de la firma (texto)

## Cómo funciona la búsqueda

La aplicación crea dos diccionarios de búsqueda:

### 1. Búsqueda por Teléfono
```
{
  "661825400": "C:\\...\\manuelpeiroperez.jpg",
  "659645517": "C:\\...\\ivancomalopez.jpg"
}
```

### 2. Búsqueda por Nombre Completo
```
{
  "MANUEL PEIRO PEREZ": "C:\\...\\manuelpeiroperez.jpg",
  "MANUEL PEIRO": "C:\\...\\manuelpeiroperez.jpg",  # También sin 2º apellido
  "MARIANO ALCALDE GRACIA": "C:\\...\\marianoalcaldegracia.jpg",
  "MARIANO ALCALDE": "C:\\...\\marianoalcaldegracia.jpg",
  "IVAN COMA LOPEZ": "C:\\...\\ivancomalopez.jpg",
  "IVAN COMA": "C:\\...\\ivancomalopez.jpg"
}
```

## Proceso de búsqueda en el PDF

Cuando la aplicación encuentra un campo resaltado en el PDF:

### Para Alumnos (campo amarillo):

1. **Si contiene un teléfono** (ej: 659645517):
   - Busca directamente en el diccionario de teléfonos
   - Si encuentra: usa esa firma
   - Si no encuentra: intenta buscar por nombre

2. **Si contiene un nombre** (ej: IVÁN COMA LÓPEZ):
   - Busca en el diccionario de nombres
   - Primero intenta coincidencia exacta
   - Si no encuentra, busca coincidencia parcial
   - Compara también sin el segundo apellido

### Para Profesores (campo rojo/naranja):

1. **Busca por nombre** (ej: MARIANO ALCALDE):
   - Busca en el diccionario de nombres
   - Primero intenta coincidencia exacta
   - Si no encuentra, busca coincidencia parcial
   - Acepta coincidencias si:
     - El nombre del PDF está contenido en algún nombre del Excel
     - Algún nombre del Excel está contenido en el nombre del PDF
     - Coinciden el nombre y el primer apellido

## Tipos de Rutas Soportadas

La aplicación soporta tres tipos de rutas en la columna **ruta**:

### 1. Rutas Absolutas de Windows
```
C:\Users\Manolaken\SALESIANOS\2GSH\fotos\ivancomalopez.jpg
```

### 2. Rutas Relativas (al Excel)
Si el Excel está en: `C:\Users\Manolaken\Documentos\LibroFirmas.xlsx`

Y la firma en: `C:\Users\Manolaken\Documentos\firmas\ivan.jpg`

Entonces en el Excel puedes poner:
```
firmas\ivan.jpg
```
O:
```
.\firmas\ivan.jpg
```

### 3. URLs (futuro soporte para Google Drive)
```
https://drive.google.com/uc?id=XXXXXXXXXXXXX
```

## Ejemplos de Coincidencias

### Ejemplo 1: Coincidencia por Teléfono
- **PDF (amarillo)**: 659645517
- **Excel**: telefono = 659645517, nombre = ivan, apellido1 = coma
- **Resultado**: ✓ Encuentra la firma de Ivan Coma

### Ejemplo 2: Coincidencia por Nombre Completo
- **PDF (amarillo)**: IVÁN COMA LÓPEZ
- **Excel**: nombre = ivan, apellido1 = coma, apellido2 = lopez
- **Resultado**: ✓ Encuentra la firma (coincidencia exacta)

### Ejemplo 3: Coincidencia Parcial (sin segundo apellido)
- **PDF (rojo)**: MARIANO ALCALDE
- **Excel**: nombre = mariano, apellido1 = alcalde, apellido2 = gracia
- **Resultado**: ✓ Encuentra la firma (coincidencia parcial: nombre + apellido1)

### Ejemplo 4: Profesor sin Teléfono
- **PDF (rojo)**: MARIANO ALCALDE GRACIA
- **Excel**: telefono = NaN, nombre = mariano, apellido1 = alcalde
- **Resultado**: ✓ Encuentra la firma por nombre

## Consejos para Preparar el Excel

1. **Nombres en minúsculas**: La aplicación los convierte automáticamente a mayúsculas
2. **Teléfonos sin espacios**: Solo números (ej: 659645517, no 659 64 55 17)
3. **Rutas válidas**: Verifica que los archivos existan en las rutas especificadas
4. **Sin celdas vacías innecesarias**: Si no hay teléfono, deja la celda vacía (no pongas 0 o "N/A")
5. **Formatos de imagen**: Usa JPG o PNG para las firmas

## Verificación del Excel

Para verificar que tu Excel está bien configurado, puedes usar el script `test_excel.py`:

```bash
python test_excel.py
```

Este script te mostrará:
- Todas las personas en el Excel
- Los diccionarios de búsqueda generados
- Las rutas de las firmas
- Si las rutas existen en el sistema

## Actualizar para Google Drive (futuro)

Cuando quieras migrar a Google Drive:

1. Sube todas las firmas a una carpeta en Google Drive
2. Para cada firma, obtén el enlace directo de descarga
3. Actualiza la columna **ruta** con las URLs de Drive
4. La aplicación detectará automáticamente que son URLs y las descargará

Formato de URL de Google Drive:
```
https://drive.google.com/uc?export=download&id=ID_DEL_ARCHIVO
```

Donde `ID_DEL_ARCHIVO` es el ID que aparece en el enlace compartido de Drive.
