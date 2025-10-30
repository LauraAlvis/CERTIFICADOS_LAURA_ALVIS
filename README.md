# CERTIFICADOS_LAURA_ALVIS
En este repositorio encontrará las diferentes herramientas que manejé para realizar el código que me permitió generar los diplomas de participación de la muestra de ingeniería industrial, teniendo en cuenta la base de datos proporcionada.
¡Claro\! Con gusto te ayudaré a generar un archivo **README** para tu repositorio de GitHub. El código que proporcionaste es un script de Python para la **Generación Masiva de Certificados** a partir de un archivo Excel y una plantilla de imagen.

Aquí tienes un **README** estructurado y profesional que puedes usar:

# 📜 Generador Masivo de Certificados (Python/Pillow)

Script de Python diseñado para la automatización de la creación de certificados personalizados en formato PDF a partir de una lista de datos en un archivo Excel (`.xlsx`) y una plantilla de imagen (`.png`). Utiliza las librerías `pandas`, `Pillow` y `openpyxl` para el procesamiento de datos, la manipulación de imágenes y la gestión de archivos.

-----

## 🚀 Características Principales

  * **Procesamiento de Datos:** Lee los datos de estudiantes, proyectos y espacios académicos desde un archivo Excel.
  * **Generación de Certificados:** Crea un certificado individual en formato **PDF** por cada estudiante con datos válidos.
  * **Manejo Dinámico de Texto:** Ajusta automáticamente el tamaño y el espaciado del nombre del proyecto en el certificado para acomodar títulos largos.
  * **Validación de Datos:**
      * Verifica la presencia de campos clave (Nombre, Proyecto, Código, Espacio Académico).
      * Registra a los estudiantes sin código válido en un archivo Excel de "Datos Faltantes" para su posterior corrección.
  * **Estructura de Carpetas:** Organiza los archivos de entrada y salida en carpetas específicas.

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
│   └── ... (otras fuentes usadas: Bodoni Bd BT Bold Italic.ttf, Bodoni Bd BT Bold.ttf, timesbd.TTF)
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

El archivo Excel debe estar ubicado en la carpeta `DATOS/` y contener **obligatoriamente** las siguientes columnas:

| Columna en Excel | Descripción |
| :--- | :--- |
| **Selecciona el espacio académico** | Nombre del espacio (materia, curso, etc.) donde se realizó el proyecto. |
| **Nombre del Proyecto** | Título completo del proyecto. |
| **Nombre completo del estudiante N** | Nombre del estudiante (donde N va de 1 a 4). |
| **Código del estudiante N** | Código o ID del estudiante (donde N va de 1 a 4). |

### 2\. Plantilla del Certificado (`plantilla.png.png`)

  * Ubica la imagen PNG de la plantilla en la carpeta `PLANTILLA/`.
  * Asegúrate de que la plantilla sea compatible con las **coordenadas** y **tamaños de fuente** definidos en el script.
  * El script está optimizado para un tamaño de imagen aproximado a un **A3 a 300dpi (2480x1754 píxeles)**, por lo que si tu plantilla es diferente, deberás ajustar las variables de posición (`posicion_...`) y el punto de anclaje central (`center_x`) dentro de la función `generar_certificado`.

### 3\. Archivos de Fuente

  * Coloca todos los archivos `.ttf` mencionados en el script (ej. `times.ttf`, `ITCEDSCR.TTF`, etc.) dentro de la carpeta `FUENTES/`.

-----

## ▶️ Uso

1.  Asegúrate de haber completado los pasos de **Requisitos** y **Preparación de Archivos**.

2.  Ejecuta el script principal:

    ```bash
    python generador_certificados.py
    ```

### Resultados de la Ejecución

  * Los certificados generados se guardarán en la carpeta `CERTIFICADOS/` como archivos **PDF**.
  * Si se encuentran estudiantes con nombre pero sin código o con un código inválido (longitud $\le 2$), el script generará el archivo `DATOS FALTANTES/estudiantes_sin_codigo.xlsx`.

-----

## ⚠️ Nota Importante sobre Rutas

El código utiliza una ruta absoluta para la carpeta de certificados, que **debe ser modificada** antes de usar el script en otro entorno:

```python
carpeta_certificados = r"C:\Users\ING\Desktop\Estudio\6 SEMESTRE\CIENCIA DE DATOS\3ER CORTE\CERTIFICADOS" # ¡CAMBIAR ESTA RUTA!
```

**Recomendación:** Cambia esta línea a una ruta relativa para mayor portabilidad, por ejemplo:

```python
carpeta_certificados = "CERTIFICADOS_GENERADOS"
```

-----

¿Te gustaría que te ayude a **generar el archivo `requirements.txt`** con las dependencias exactas del proyecto?
