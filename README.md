# TXT-A-XLSX

Conversor original recuperado de los archivos del proyecto (septiembre de 2026). El script **no cambia ni elimina los TXT de origen**.

## Funciones

1. **Individual**: un XLSX por TXT, manteniendo subcarpetas.
2. **Por carpeta**: un XLSX por carpeta, con cada TXT en una hoja y una hoja `INDICE`.
3. **Unificado**: todos los TXT en un XLSX, con hojas individuales e `INDICE` de rutas de origen.

Detecta tabuladores, comas, punto y coma o `|`; intenta codificaciones habituales de Windows y preserva los valores como texto, incluidos ceros iniciales.

## Uso

```bash
python -m pip install -r requirements.txt
python txt-to-xlsx-menu-v2.py
```

Se abre una ventana para seleccionar carpeta y modo. Requiere Python 3.10+ y `tkinter` instalado.

## Validación del respaldo

Se comprobó la sintaxis con `py_compile` y los tres modos con tres TXT sintéticos en distintas carpetas. Se verificó la preservación de ceros iniciales y que los archivos de origen no cambiaran. **No se han usado archivos operativos reales para esta prueba**.

Código recuperado de la versión previa del usuario; mantener historial de cambios y validar en Windows antes de uso operativo.
