import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime
from openpyxl import Workbook
import textwrap
from openpyxl.utils.dataframe import dataframe_to_rows

def generar_certificados():
    # Definir rutas de las carpetas
    carpeta_datos = "DATOS"
    # Usamos una ruta absoluta para la carpeta de certificados
    carpeta_certificados = r"C:\Users\ING\Desktop\Estudio\6 SEMESTRE\CIENCIA DE DATOS\3ER CORTE\CERTIFICADOS"
    carpeta_plantilla = "PLANTILLA"
    carpeta_datos_faltantes = "DATOS FALTANTES"

    # Archivo de plantilla
    plantilla_path = os.path.join(carpeta_plantilla, "plantilla.png.png")

    # Archivo de datos Excel
    excel_path = os.path.join(carpeta_datos, "datos.xlsx")

    # Crear carpetas si no existen
    if not os.path.exists(carpeta_certificados):
        os.makedirs(carpeta_certificados)
    if not os.path.exists(carpeta_datos_faltantes):
        os.makedirs(carpeta_datos_faltantes)

    # Leer el archivo Excel
    try:
        # Especificar dtype para las columnas de código para evitar que pandas las interprete como números
        dtype_mapping = {
            "Código del estudiante 1": str,
            "Código del estudiante 2": str,
            "Código del estudiante 3": str,
            "Código del estudiante 4": str,
        }
        df = pd.read_excel(excel_path, dtype=dtype_mapping)
        # Limpiar posibles espacios en blanco en los códigos
        for col in dtype_mapping.keys():
            df[col] = df[col].str.strip()
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo Excel en la ruta: {excel_path}")
        return
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return

    # Columnas necesarias
    columnas_necesarias = [
        "Selecciona el espacio académico",
        "Nombre del Proyecto",
        "Nombre completo del estudiante 1",
        "Código del estudiante 1",
        "Nombre completo del estudiante 2",
        "Código del estudiante 2",
        "Nombre completo del estudiante 3",
        "Código del estudiante 3",
        "Nombre completo del estudiante 4",
        "Código del estudiante 4"
    ]

    # Verificar si todas las columnas necesarias están presentes
    if not all(col in df.columns for col in columnas_necesarias):
        print("Error: No todas las columnas necesarias están presentes en el archivo Excel.")
        return

    # Lista para almacenar estudiantes sin código
    estudiantes_sin_codigo = []

    # Iterar sobre las filas del DataFrame
    for index, row in df.iterrows():
        # Función auxiliar para validar los datos
        # Un dato es válido si no es nulo, no está en blanco y no es solo un punto "."
        def es_dato_valido(dato):
            if pd.isnull(dato): return False
            texto = str(dato).strip()
            return texto != "" and texto != "."

        # Obtener datos generales del proyecto
        espacio_academico = str(row["Selecciona el espacio académico"]).strip().upper() if es_dato_valido(row["Selecciona el espacio académico"]) else None
        nombre_proyecto = str(row["Nombre del Proyecto"]).strip().upper() if es_dato_valido(row["Nombre del Proyecto"]) else None

        # Iterar sobre los estudiantes (hasta 4)
        for i in range(1, 5):
            nombre_estudiante_col = f"Nombre completo del estudiante {i}"
            codigo_estudiante_col = f"Código del estudiante {i}"

            # Obtener datos del estudiante (inicialmente, sin validación de longitud de código)
            nombre_estudiante = str(row[nombre_estudiante_col]).strip() if es_dato_valido(row[nombre_estudiante_col]) else None
            codigo_estudiante = str(row[codigo_estudiante_col]).strip() if es_dato_valido(row[codigo_estudiante_col]) else None

            # --- Validaciones de datos antes de generar el certificado ---

            # 1. Validar que los datos esenciales (nombre, proyecto, espacio) estén presentes y sean válidos
            if not nombre_estudiante or not espacio_academico or not nombre_proyecto:
                # Si el nombre del estudiante no es válido, no lo rastreamos en estudiantes_sin_codigo
                if nombre_estudiante: # Si el nombre es válido, pero falta otro dato esencial
                    print(f"Advertencia: Datos esenciales (proyecto/espacio académico) faltantes o inválidos para {nombre_estudiante}. No se generará el certificado.")
                continue # Saltar la generación de certificado para este estudiante

            # 2. Validar el código del estudiante (presencia)
            if not codigo_estudiante:
                if nombre_estudiante: # Si el nombre es válido, pero el código no
                    estudiantes_sin_codigo.append({"Nombre": nombre_estudiante})
                    print(f"Advertencia: Código de estudiante faltante o inválido para {nombre_estudiante}. No se generará el certificado.")
                continue # Saltar la generación de certificado

            # 3. Nueva validación: longitud del código del estudiante
            if len(codigo_estudiante) <= 2:
                if nombre_estudiante: # Si el nombre es válido, pero el código es demasiado corto
                    estudiantes_sin_codigo.append({"Nombre": nombre_estudiante})
                    print(f"Advertencia: El código del estudiante {nombre_estudiante} ({codigo_estudiante}) tiene 2 o menos caracteres. No se generará el certificado.")
                continue # Saltar la generación de certificado

            # Si todas las validaciones pasan, generar el certificado
            # Generar el certificado
            try:
                generar_certificado(
                    plantilla_path,
                    carpeta_certificados,
                    nombre_estudiante,
                    codigo_estudiante,
                    nombre_proyecto,
                    espacio_academico
                )
            except Exception as e:
                print(f"Error al generar el certificado para {nombre_estudiante}: {e}")

    # Crear archivo Excel con estudiantes sin código
    if estudiantes_sin_codigo:
        crear_excel_estudiantes_sin_codigo(estudiantes_sin_codigo, carpeta_datos_faltantes)

def generar_certificado(plantilla_path, carpeta_certificados, nombre_estudiante, codigo_estudiante, nombre_proyecto, espacio_academico):
    # Cargar la plantilla
    try:
        # Abrimos la imagen en modo RGBA para manejar transparencias si las hubiera
        img = Image.open(plantilla_path).convert("RGBA")
    except FileNotFoundError:
        print(f"Error: No se encontró la plantilla en la ruta: {plantilla_path}")
        return
    except Exception as e:
        print(f"Error al cargar la plantilla: {e}")
        return

    draw = ImageDraw.Draw(img)

    # Definir fuentes y tamaños
    carpeta_fuentes = "FUENTES"
    font_path = os.path.join(carpeta_fuentes, "times.ttf") # Fuente por defecto
    font_estudiante_path = os.path.join(carpeta_fuentes, "ITCEDSCR.TTF") # Fuente para el nombre del estudiante
    font_participacion_path = os.path.join(carpeta_fuentes, "Bodoni Bd BT Bold Italic.ttf")
    font_diploma_path = os.path.join(carpeta_fuentes, "Bodoni Bd BT Bold.ttf")
    font_negrilla_path = os.path.join(carpeta_fuentes, "timesbd.TTF") # Fuente negrita times
    try:
        font_diploma = ImageFont.truetype(font_diploma_path , 80)
        font_participacion = ImageFont.truetype(font_participacion_path , 60)
        font_intro = ImageFont.truetype(font_path, 40)
        font_nombre_estudiante = ImageFont.truetype(font_estudiante_path, 95) # Usando la fuente cursiva y un tamaño ajustado
        font_proyecto_nombre_large = ImageFont.truetype(font_negrilla_path, 50) # Fuente para proyectos de 1 o 2 líneas
        font_proyecto_nombre_small = ImageFont.truetype(font_negrilla_path, 45) # Fuente para proyectos de 3 líneas
        font_proyecto_nombre_xsmall = ImageFont.truetype(font_negrilla_path, 40) # Fuente para proyectos muy largos (más de 3 líneas)
        font_codigo_estudiante = ImageFont.truetype(font_path, 40) # Nueva fuente para el código del estudiante
        font_espacio_nombre = ImageFont.truetype(font_negrilla_path, 45)
        font_fecha = ImageFont.truetype(font_path, 40)
        
    except IOError:
        print(f"Error: No se encontró un archivo de fuente. Asegúrate de que los archivos de fuente estén en la carpeta '{carpeta_fuentes}'.")
        return
    except Exception as e:
        print(f"Error al cargar la fuente: {e}")
        return
    
    # Coordenada X para centrar el texto (ancho de la imagen / 2)
    # La imagen de la plantilla tiene un tamaño de 2480x1754 píxeles (aproximadamente A3 a 300dpi)
    # Si tu plantilla es diferente, ajusta estas coordenadas.
    # Basado en la imagen proporcionada, el ancho es 2480.
    # center_x = 2480 / 2 = 1240
    center_x = img.width / 2

    # Definir colores

    # Definir colores
    color = 'rgb(0, 0, 0)'
    color_amarillo = 'rgb(214, 175, 36)'

    # Definir posiciones (ajusta estas coordenadas según tu plantilla)
    # El primer valor es la distancia desde la izquierda (eje X)
    # El segundo valor es la distancia desde arriba (eje Y)
    # El anchor="mm" centra el texto en esas coordenadas (middle-middle)
    posicion_diploma = (center_x, 300) # Más arriba
    posicion_participacion = (center_x, 400) # Más arriba
    posicion_intro_estudiante = (center_x, 500)
    posicion_nombre_estudiante = (center_x, 570)
    posicion_codigo_y_proyecto = (center_x, 650) # Posición para el texto combinado
    
    # Posiciones iniciales para elementos que se moverán dinámicamente
    initial_y_nombre_proyecto = 720 # Ajustado hacia arriba
    initial_y_intro_espacio = 820 # Ajustado hacia arriba
    initial_y_nombre_espacio = 900 # Ajustado hacia arriba
    initial_y_fecha = 990 # Ajustado hacia arriba

    # Obtener la fecha actual
    fecha_actual = datetime.now().strftime("%d de %B de %Y").replace("January", "Enero").replace("February", "Febrero").replace("March", "Marzo").replace("April", "Abril").replace("May", "Mayo").replace("June", "Junio").replace("July", "Julio").replace("August", "Agosto").replace("September", "Septiembre").replace("October", "Octubre").replace("November", "Noviembre").replace("December", "Diciembre")

    # Escribir el texto en la imagen
    # --- Bloque de texto superior ---
    draw.text(posicion_diploma, "DIPLOMA", fill=color_amarillo, font=font_diploma, anchor="mm")
    draw.text(posicion_participacion, "PARTICIPACIÓN MUESTRA DE INGENIERÍA", fill=color_amarillo, font=font_participacion, anchor="mm")
    draw.text(posicion_intro_estudiante, "La facultad de Ingeniería Industrial Seccional Villavicencio, hace constar que el estudiante:", fill=color, font=font_intro, anchor="mm")
    draw.text(posicion_nombre_estudiante, nombre_estudiante, fill=color, font=font_nombre_estudiante, anchor="mm")    
    texto_combinado = f"con código o ID: {codigo_estudiante}, participó como ponente en modalidad póster con el proyecto:"
    draw.text(posicion_codigo_y_proyecto, texto_combinado, fill=color, font=font_intro, anchor="mm")
    
    # --- Bloque de nombre del proyecto (dinámico) ---
    # Intentar con la fuente grande primero
    current_font_proyecto = font_proyecto_nombre_large
    width_wrap = 45 # Ancho para la fuente grande
    lineas_proyecto = textwrap.wrap(nombre_proyecto, width=width_wrap)

    # Si el proyecto es muy largo (más de 2 líneas), reducir la fuente progresivamente
    if len(lineas_proyecto) > 2: # Si con la fuente grande ocupa más de 2 líneas
        current_font_proyecto = font_proyecto_nombre_small
        width_wrap = 55 # Ancho para la fuente pequeña (más caracteres por línea)
        lineas_proyecto = textwrap.wrap(nombre_proyecto, width=width_wrap)

        # Si con la fuente mediana aún ocupa más de 3 líneas
        if len(lineas_proyecto) > 3: 
            current_font_proyecto = font_proyecto_nombre_xsmall
            width_wrap = 65 # Ancho para la fuente más pequeña
            lineas_proyecto = textwrap.wrap(nombre_proyecto, width=width_wrap)

    y_actual_proyecto = initial_y_nombre_proyecto
    
    # Calcular la altura de una sola línea de texto del proyecto (para el offset)
    # Usamos un texto de referencia para obtener la altura de la fuente
    single_line_height_ref = current_font_proyecto.getbbox("Ejemplo de texto")[3] - current_font_proyecto.getbbox("Ejemplo de texto")[1]
    line_spacing = 10 # Espacio entre líneas del proyecto

    for linea in lineas_proyecto:
        draw.text((center_x, y_actual_proyecto), linea, fill=color, font=current_font_proyecto, anchor="mm")
        text_height = current_font_proyecto.getbbox(linea)[3] - current_font_proyecto.getbbox(linea)[1]
        y_actual_proyecto += text_height + line_spacing # Mover a la siguiente línea

    # Calcular el offset para el texto subsiguiente
    # El offset es la altura adicional que tomó el bloque del proyecto más allá de una sola línea
    total_project_block_height = y_actual_proyecto - initial_y_nombre_proyecto
    single_line_project_block_height = single_line_height_ref + line_spacing # Altura que ocuparía una sola línea
    
    offset_y = total_project_block_height - single_line_project_block_height # Diferencia de altura
    if offset_y < 0: offset_y = 0 # Asegurarse de que el offset no sea negativo

    # --- Bloque de espacio académico y fecha (dinámico) ---
    draw.text((center_x, initial_y_intro_espacio + offset_y), "estudio realizado en el espacio académico de:", fill=color, font=font_intro, anchor="mm")
    draw.text((center_x, initial_y_nombre_espacio + offset_y), espacio_academico, fill=color, font=font_espacio_nombre, anchor="mm")
    draw.text((center_x, initial_y_fecha + offset_y), f" El certificado se expide el {fecha_actual}.", fill=color, font=font_fecha, anchor="mm")

    # Crear una imagen de fondo blanca para el PDF
    pdf_canvas = Image.new('RGB', img.size, (255, 255, 255))
    # Pegar la imagen del certificado (con su posible transparencia) sobre el fondo blanco
    pdf_canvas.paste(img, mask=img.split()[3]) # El canal Alpha (transparencia) se usa como máscara

    # Guardar el certificado como PDF
    nombre_archivo = f"{nombre_estudiante.replace(' ', '_')}_{codigo_estudiante}.pdf"
    ruta_archivo = os.path.join(carpeta_certificados, nombre_archivo)
    pdf_canvas.save(ruta_archivo, "PDF", resolution=100.0)

    print(f"Certificado generado para {nombre_estudiante} en: {ruta_archivo}")

def crear_excel_estudiantes_sin_codigo(estudiantes_sin_codigo, carpeta_destino):
    # Crear un nuevo libro de trabajo de OpenPyXL
    wb = Workbook()
    ws = wb.active
    ws.title = "Estudiantes sin código"

    # Encabezados
    ws.append(["Nombre del Estudiante"])

    # Agregar los nombres de los estudiantes sin código al archivo Excel
    for estudiante in estudiantes_sin_codigo:
        ws.append([estudiante["Nombre"]])

    # Guardar el archivo Excel
    excel_path = os.path.join(carpeta_destino, "estudiantes_sin_codigo.xlsx")
    try:
        wb.save(excel_path)
        print(f"Archivo Excel de estudiantes sin código generado en: {excel_path}")
    except Exception as e:
        print(f"Error al guardar el archivo Excel: {e}")

# Ejecutar la función principal
if __name__ == "__main__":
    generar_certificados()
