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
        # Forzar a que TODAS las columnas se lean como texto (string) para evitar errores de tipo.
        # Esto es más robusto y previene que pandas interprete datos como números incorrectamente.
        df = pd.read_excel(excel_path, dtype=str)
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
    # Conjunto para rastrear certificados generados y evitar duplicados (estudiante, espacio_academico)
    certificados_generados = set()

    # --- PASO 1: Pre-procesamiento para crear un mapa de códigos válidos ---
    mapa_codigos_validos = {}
    print("Analizando datos y creando mapa de códigos de estudiantes...")

    # Función unificada para limpiar y validar datos. Devuelve el dato limpio o None si es inválido.
    def obtener_dato_limpio(dato):
        if pd.isnull(dato):
            return None
        texto_limpio = str(dato).strip()
        if texto_limpio.lower() in ["", "nan", "."]:
            return None
        return texto_limpio

    for index, row in df.iterrows():
        for i in range(1, 5):
            nombres_str = obtener_dato_limpio(row[f"Nombre completo del estudiante {i}"])
            codigos_str = obtener_dato_limpio(row[f"Código del estudiante {i}"])

            if not nombres_str or not codigos_str:
                continue

            nombres_lista = [n.strip() for n in nombres_str.split('-')]
            codigos_lista = [c.strip() for c in codigos_str.split('-')]

            # Procesar cada par de nombre/código, incluso si no hay guion
            for nombre_estudiante, codigo_estudiante in zip(nombres_lista, codigos_lista):
                if not nombre_estudiante or not codigo_estudiante:
                    continue

                # Un código es válido si existe, tiene más de 2 caracteres y no es numéricamente igual a cero.
                codigo_es_valido = False
                if len(codigo_estudiante) > 2:
                    if not (codigo_estudiante.isdigit() and int(codigo_estudiante) == 0):
                        codigo_es_valido = True

                if codigo_es_valido:
                    # Si el estudiante no está en el mapa, añadimos su código válido.
                    nombre_estudiante_upper = nombre_estudiante.upper()
                    if nombre_estudiante_upper not in mapa_codigos_validos:
                        mapa_codigos_validos[nombre_estudiante_upper] = codigo_estudiante

    # Iterar sobre las filas del DataFrame
    for index, row in df.iterrows():
        # Obtener datos generales del proyecto
        nombre_proyecto = obtener_dato_limpio(row["Nombre del Proyecto"])
        espacio_academico = obtener_dato_limpio(row["Selecciona el espacio académico"])
        if not espacio_academico:
            espacio_academico = "ESPACIO ACADÉMICO NO ESPECIFICADO"

        # Iterar sobre los estudiantes (hasta 4)
        for i in range(1, 5):
            nombres_str = obtener_dato_limpio(row[f"Nombre completo del estudiante {i}"])
            if not nombres_str:
                continue

            nombres_lista = [n.strip() for n in nombres_str.split('-')]

            for nombre_estudiante_original in nombres_lista:
                if not nombre_estudiante_original:
                    continue

                # --- Validaciones de datos antes de generar el certificado ---
                # 1. Validar que los datos esenciales (nombre y proyecto) estén presentes.
                if not nombre_proyecto:
                    if nombre_estudiante_original:
                        print(f"Advertencia: Nombre de proyecto faltante o inválido para {nombre_estudiante_original} en la fila {index+2}. No se generará el certificado.")
                    continue

                # 2. Obtener el código correcto desde el mapa de códigos válidos.
                nombre_estudiante_upper = nombre_estudiante_original.upper()
                codigo_estudiante = mapa_codigos_validos.get(nombre_estudiante_upper)

                if not codigo_estudiante:
                    estudiantes_sin_codigo.append({"Nombre": nombre_estudiante_original})
                    print(f"Error: No se pudo encontrar un código válido para {nombre_estudiante_original} en todo el archivo. No se generará el certificado para este proyecto.")
                    continue

                # 3. Verificar si ya se generó un certificado para este estudiante en este espacio académico
                espacio_academico_upper = espacio_academico.upper()
                identificador_certificado = (nombre_estudiante_upper, espacio_academico_upper)

                if identificador_certificado in certificados_generados:
                    print(f"Info: Certificado duplicado para '{nombre_estudiante_original}' en el espacio académico '{espacio_academico}'. Omitiendo.")
                    continue

                # Generar el certificado
                try:
                    generar_certificado(
                        plantilla_path,
                        carpeta_certificados,
                        nombre_estudiante_original,
                        codigo_estudiante,
                        nombre_proyecto.upper(),
                        espacio_academico_upper
                    )
                    certificados_generados.add(identificador_certificado)
                except Exception as e:
                    print(f"Error al generar el certificado para {nombre_estudiante_original}: {e}")

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
    draw.text(posicion_nombre_estudiante, nombre_estudiante.title(), fill=color, font=font_nombre_estudiante, anchor="mm")    
    texto_combinado = f"con código o ID: {codigo_estudiante}, participó como ponente en modalidad póster con el proyecto:"
    draw.text(posicion_codigo_y_proyecto, texto_combinado, fill=color, font=font_intro, anchor="mm")
    
    # --- Bloque de nombre del proyecto (dinámico) ---
    # Esta función ahora se encarga de todo el proceso de dibujado del proyecto
    def dibujar_texto_multilinea(texto, y_inicial):
        # Intentar con la fuente grande primero
        fuente_actual = font_proyecto_nombre_large
        ancho_wrap = 45
        lineas = textwrap.wrap(texto, width=ancho_wrap)

        # Si el proyecto es muy largo, reducir la fuente progresivamente
        if len(lineas) > 2:
            fuente_actual = font_proyecto_nombre_small
            ancho_wrap = 55
            lineas = textwrap.wrap(texto, width=ancho_wrap)
        if len(lineas) > 3:
            fuente_actual = font_proyecto_nombre_xsmall
            ancho_wrap = 65
            lineas = textwrap.wrap(texto, width=ancho_wrap)

        y_actual = y_inicial
        espacio_linea = 10
        altura_linea_unica = fuente_actual.getbbox("A")[3] - fuente_actual.getbbox("A")[1]

        for linea in lineas:
            draw.text((center_x, y_actual), linea, fill=color, font=fuente_actual, anchor="mm")
            y_actual += altura_linea_unica + espacio_linea
        
        return y_actual - y_inicial # Devuelve la altura total del bloque de texto

    altura_bloque_proyecto = dibujar_texto_multilinea(nombre_proyecto, initial_y_nombre_proyecto)

    # Calcular el offset para el texto subsiguiente
    # El offset es la altura adicional que tomó el bloque del proyecto más allá de una sola línea
    altura_linea_ref = font_proyecto_nombre_large.getbbox("A")[3] - font_proyecto_nombre_large.getbbox("A")[1] + 10
    offset_y = altura_bloque_proyecto - altura_linea_ref
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
    # Incluir el nombre del proyecto para asegurar que cada archivo sea único
    
    # Sanitizar el nombre del archivo para eliminar caracteres inválidos (como ':')
    def sanitizar_nombre_archivo(nombre):
        caracteres_invalidos = '<>:"/\\|?*'
        for char in caracteres_invalidos:
            nombre = nombre.replace(char, '')
        return nombre

    nombre_archivo_base = f"{nombre_estudiante.replace(' ', '_')}_{nombre_proyecto.replace(' ', '_')[:30]}_{codigo_estudiante}"
    nombre_archivo_sanitizado = sanitizar_nombre_archivo(nombre_archivo_base) + ".pdf"
    ruta_archivo = os.path.join(carpeta_certificados, nombre_archivo_sanitizado)
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
