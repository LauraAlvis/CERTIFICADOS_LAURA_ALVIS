# CERTIFICADOS_LAURA_ALVIS
En este repositorio encontrará las diferentes herramientas que manejé para realizar el código que me permitió generar los diplomas de participación de la muestra de ingeniería industrial, teniendo en cuenta la base de datos proporcionada.

# 📜 Generador Masivo de Certificados (Python/Pillow) 

Script de Python diseñado para la automatización de la creación de certificados personalizados en formato PDF a partir de una lista de datos en un archivo Excel (`.xlsx`) y una plantilla de imagen (`.png`). Esta versión incluye un **pre-procesamiento robusto de datos** y manejo de múltiples estudiantes por celda.

-----

## 🚀 Características Principales

  * **Pre-procesamiento Robusto:** Analiza todo el archivo de datos para crear un mapa de códigos válidos, incluso si el código está en una fila diferente a la del proyecto.
  * **Manejo de Grupos:** Soporta el registro de **múltiples estudiantes y códigos en una sola celda**, siempre y cuando estén separados por un guion (`-`).
  * **Prevención de Duplicados:** Evita generar el mismo certificado más de una vez para un estudiante dentro del mismo espacio académico.
  * **Manejo Dinámico de Texto:** Ajusta automáticamente el tamaño y el espaciado del nombre del proyecto.
  * **Validación de Código:** Un código es válido si tiene más de 2 caracteres y no es numéricamente igual a cero.
  * **Registro de Faltantes:** Registra los estudiantes sin código válido en un archivo Excel de "Datos Faltantes".

-----

## ⚙️ Estructura del Repositorio

Asegúrate de que la estructura de carpetas de tu proyecto sea la siguiente:

```
.
├── CERTIFICADOS/              # 📁 Carpeta de SALIDA (Contiene los PDF generados)
├── DATOS/                     # 📁 Carpeta de ENTRADA (Contiene el archivo de datos)
│   └── datos.xlsx             # 📄 Archivo Excel con la información de los estudiantes
├── DATOS FALTANTES/           # 📁 Carpeta de SALIDA (Contiene el Excel de estudiantes sin código)
├── PLANTILLA/                 # 📁 Carpeta de ENTRADA (Contiene la imagen de la plantilla)
│   └── plantilla.png.png      # 🖼️ Imagen PNG de la plantilla del certificado
├── FUENTES/                   # 📁 Carpeta de ENTRADA (Contiene los archivos de fuentes TrueType)
│   ├── times.ttf
│   ├── ITCEDSCR.TTF
│   └── ... (otras fuentes)
└── generador_certificados.py  # 🐍 El script principal
```

-----

## 📥 Requisitos

Este proyecto requiere **Python 3.x** y las siguientes librerías.

Para instalar las dependencias, ejecuta el siguiente comando:

```bash
pip install pandas Pillow openpyxl
```

-----

## 📝 Preparación de Archivos

### 1\. Archivo de Datos (`datos.xlsx`)

El archivo Excel debe estar ubicado en la carpeta `DATOS/` y debe contener las siguientes columnas, **leyendo todos los datos como texto (`str`) para evitar errores de tipo**.

| Columna en Excel | Descripción | Formato de Datos |
| :--- | :--- | :--- |
| **Selecciona el espacio académico** | Nombre del espacio (materia, curso, etc.). | Texto |
| **Nombre del Proyecto** | Título completo del proyecto. | Texto |
| **Nombre completo del estudiante N** | Nombre del estudiante (donde N va de 1 a 4). | **Texto** (Soporta múltiples nombres separados por `-`) |
| **Código del estudiante N** | Código o ID del estudiante (donde N va de 1 a 4). | **Texto** (Soporta múltiples códigos separados por `-`) |

> ⚠️ **Múltiples Estudiantes en la Misma Celda:**
> Si tienes varios estudiantes en una sola celda (ej. si dos estudiantes presentaron el mismo proyecto), sepáralos usando un **guion (`-`)** tanto en la columna de nombres como en la de códigos, asegurando que el orden sea consistente.
>
>   * **`Nombre completo del estudiante 1`**: `Juan Pérez - María Gómez`
>   * **`Código del estudiante 1`**: `12345678 - 87654321`

### 2\. Plantilla y Fuentes

  * Ubica la imagen PNG de la plantilla en la carpeta `PLANTILLA/`.
  * Coloca todos los archivos `.ttf` de las fuentes necesarias en la carpeta `FUENTES/`.

-----

## ▶️ Uso

1.  Asegúrate de que tu ruta de salida en el código Python esté correcta (ver nota de la ruta abajo).

2.  Asegúrate de haber completado la **Preparación de Archivos**.

3.  Ejecuta el script principal:

    ```bash
    python generador_certificados.py
    ```

### Resultados de la Ejecución

  * Los certificados generados se guardarán en la carpeta `CERTIFICADOS/` como archivos **PDF**, usando una combinación del nombre del estudiante, las primeras 30 letras del proyecto y el código para generar el nombre del archivo.
  * Si se encuentran estudiantes con nombre pero sin un código válido, se creará el archivo **`DATOS FALTANTES/estudiantes_sin_codigo.xlsx`**.

-----

## ⚠️ Nota Importante sobre la Ruta de Salida

El código utiliza actualmente una **ruta absoluta** para la carpeta de certificados, la cual **debes modificar** para que funcione en tu máquina o servidor:

```python
carpeta_certificados = r"C:\Users\ING\Desktop\Estudio\6 SEMESTRE\CIENCIA DE DATOS\3ER CORTE\CERTIFICADOS" # ¡MODIFICAR ESTA RUTA!
```

**Recomendación:** Para hacerlo portable, cámbiala a una ruta relativa si deseas que los certificados se guarden dentro del mismo directorio del proyecto:

```python
carpeta_certificados = "CERTIFICADOS_GENERADOS" # Ejemplo de ruta relativa
```
