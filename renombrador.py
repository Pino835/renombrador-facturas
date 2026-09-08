# -*- coding: utf-8 -*-
"""
renombrador.py
-------------------------------------------------------------------------
"RENOMBRADOR" - Automatizacion de cuentas por cobrar.

QUE HACE:
  1. Revisa, uno por uno, los PDF escaneados que estan en Escritorio\\PRUEBA.
  2. A cada uno le hace OCR para encontrar el "codigo de cliente" y el
     numero de "factura" (formato CMER-XXXXXXX).
  3. Si logra leer los dos datos, guarda una copia del PDF ya renombrada
     ([codigo_cliente]-[factura].pdf) en Escritorio\\RESULTADO\\<fecha de hoy>\\
  4. Si NO logra leer alguno de los dos datos, copia el PDF a
     Escritorio\\POR_REVISAR y sigue con el siguiente archivo (no
     interrumpe el proceso automatico).
  5. Al terminar de revisar TODOS los archivos automaticos, recien ahi
     abre, uno por uno, una pantalla de revision manual para cada PDF que
     quedo en POR_REVISAR: muestra dos recortes de la factura (donde
     deberian estar el codigo de cliente y la factura) y dos campos para
     digitarlos a mano. Al aceptar, guarda el archivo en RESULTADO\\<fecha>
     y borra la copia de POR_REVISAR.
  6. Registra cada resultado (automatico o manual) en un log en:
         Escritorio\\PROYECTO_RENOMBRADOR\\log_renombrador.json

SOBRE EL BOTON "CANCELAR" (recomendacion incluida en este programa):
  Se agrego porque, sin el, la gerente estaria obligada a inventar o
  adivinar un dato con tal de poder cerrar la ventana. "Cancelar" NO
  cierra el programa completo: solo omite ese archivo por ahora y lo deja
  intacto en POR_REVISAR (no se borra, no se renombra), para revisarlo
  con calma mas tarde -- ya sea abriendo el PDF directamente o volviendo
  a ejecutar este programa, que lo va a volver a mostrar en la pantalla
  de revision manual. El programa sigue de inmediato con el siguiente
  archivo pendiente. Si en algun momento prefieren que "Cancelar" detenga
  TODO el proceso en lugar de solo saltar el archivo, es el cambio
  senalado con "CAMBIAR AQUI" mas abajo.
-------------------------------------------------------------------------
"""

import re
import os
import sys
import csv
import json
import shutil
from pathlib import Path
from datetime import datetime

import fitz  # PyMuPDF
import pytesseract
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk, messagebox

# =========================================================================
# CONFIGURACION - ajuste aqui si algo cambia
# =========================================================================

# Tesseract se puede haber instalado en distintos lugares segun como haya
# quedado el equipo (instalacion normal en Program Files, o instalacion
# automatica sin permisos de administrador que hace "instalar_dependencias.py"
# en la carpeta del usuario). Se revisan todas las ubicaciones conocidas.
_CANDIDATOS_TESSERACT = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Tesseract-OCR", "tesseract.exe"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
]

_tesseract_encontrado = shutil.which("tesseract")
if not _tesseract_encontrado:
    for _ruta_candidata in _CANDIDATOS_TESSERACT:
        if _ruta_candidata and os.path.exists(_ruta_candidata):
            _tesseract_encontrado = _ruta_candidata
            break

if _tesseract_encontrado:
    pytesseract.pytesseract.tesseract_cmd = _tesseract_encontrado

# Aviso si falta el idioma Español: el programa igual "funciona" sin el
# (Tesseract no lo rechaza, solo lee con el modelo de Ingles), pero la
# precision baja notablemente en facturas con tildes y textos en Español.
# Esto ya se soluciona solo al correr "instalar_dependencias.py"; este
# aviso es solo para que quede claro en pantalla si por algun motivo el
# idioma no quedo instalado.
if _tesseract_encontrado:
    _ruta_spa = os.path.join(os.path.dirname(_tesseract_encontrado), "tessdata", "spa.traineddata")
    if not os.path.exists(_ruta_spa):
        print("ADVERTENCIA: no se encontro el idioma Español para el OCR")
        print(f"  ({_ruta_spa})")
        print("  La lectura de las facturas va a ser menos precisa de lo normal.")
        print("  Vuelva a ejecutar 'ejecutar_1_instalar_dependencias.bat' para instalarlo.")

IDIOMAS_OCR = "spa+eng"

ESCRITORIO = Path.home() / "Desktop"
CARPETA_ENTRADA = ESCRITORIO / "PRUEBA"
CARPETA_POR_REVISAR = ESCRITORIO / "POR_REVISAR"
CARPETA_RESULTADO = ESCRITORIO / "RESULTADO"
CARPETA_PROYECTO = ESCRITORIO / "PROYECTO_RENOMBRADOR"

FECHA_HOY = datetime.now().strftime("%d-%m-%Y")
CARPETA_RESULTADO_HOY = CARPETA_RESULTADO / FECHA_HOY
LOG_PATH = CARPETA_PROYECTO / "log_renombrador.json"
LOG_PATH_CSV_ANTIGUO = CARPETA_PROYECTO / "log_renombrador.csv"  # log viejo, solo para migrar el historial una vez

# Lista de codigos de cliente ya confirmados por una persona (se llena
# sola: cada vez que se acepta un codigo en la ventana de revision manual,
# se agrega aqui si no estaba). Sirve para corregir automaticamente
# lecturas futuras de un cliente que ya conocemos y que el OCR a veces lee
# mal (ej. "MCMO001" en vez de "MCM001"). Es un archivo de texto simple,
# un codigo por linea -- se puede abrir y editar a mano si hace falta.
CODIGOS_CONFIRMADOS_PATH = CARPETA_PROYECTO / "codigos_cliente_confirmados.txt"

DPI_RENDER = 300  # resolucion para convertir el PDF a imagen (mas alto = mejor OCR, mas lento)

# Confianza minima (0-100) que debe tener la lectura del OCR para guardarse
# de forma automatica. Con pruebas reales se vio que el OCR a veces "SI"
# encuentra un codigo de cliente, pero lo lee mal (ej. confunde un "0" con
# una "O", o una "I" con un "1") sin que eso se note solo con que "encontro
# algo" -- y a veces con una confianza bastante alta (ej. "CTLIOOO" en vez
# de "CTLI000" se leyo con 76-85% de confianza en varias pruebas). Si la
# confianza de la lectura queda por debajo de este umbral, el archivo se
# manda a revision manual (con el dato ya precargado, solo para que se
# confirme con la vista) en lugar de guardarse solo.
#
# MAS ALTO = mas estricto = menos errores como el de arriba se cuelan,
#            pero mas archivos van a revision manual.
# MAS BAJO = menos estricto = menos trabajo manual, pero mas riesgo de que
#            un dato mal leido se guarde solo sin que nadie lo note.
#
# Se subio de 75 a 80 despues de encontrar el caso de "CTLIOOO", que con
# 75 pasaba de largo. Esta confianza NO distingue perfecto entre aciertos
# y errores (a veces el OCR se equivoca "con seguridad"), asi que subir
# mucho mas este numero no garantiza atrapar todo -- solo reduce el riesgo.
UMBRAL_CONFIANZA_MINIMA = 80.0

# La factura SIEMPRE tiene el formato "CMER-xxxxxxx" (letras y/o numeros
# despues de CMER), asi que se busca ese patron de forma directa en
# cualquier parte de la pagina: es muy especifico y practicamente no se
# confunde con otro texto de la factura (ej. los parrafos de terminos).
PATRON_CMER = re.compile(r'CMER[\s\-]{0,2}([A-Za-z0-9]{3,15})', re.IGNORECASE)

# Respaldo por si la factura no sigue el patron CMER: se busca solo la
# etiqueta "Factura" en la mitad superior del documento (el valor se toma
# aparte, palabra por palabra, ver _valor_tras_posicion).
PATRON_FACTURA_ETIQUETA = re.compile(r'Factura\s*[:\-]?', re.IGNORECASE)

# El codigo de cliente no tiene un patron unico (son letras y/o numeros),
# por lo que se busca siempre junto a su etiqueta ("Cod. Cliente",
# "Codigo de Cliente", "Cliente No.", etc.), solo en la mitad superior
# del documento para no confundirlo con la palabra "cliente" que puede
# aparecer mas abajo en avisos legales ("Estimado cliente..."). Aqui solo
# se detecta DONDE esta la etiqueta; el valor se busca aparte (ver
# _valor_tras_posicion) porque el OCR a veces mete un caracter basura
# entre los dos puntos y el valor (ej. "Cliente: � REMJ003").
PATRON_CLIENTE_ETIQUETA = re.compile(
    r'(?:C[oóeé0]d\.?\s*(?:de\s*)?Cliente|Cliente\s*(?:No\.?|N[°º]|#))\s*[:\-]?',
    re.IGNORECASE,
)

LIMITE_SUPERIOR_RATIO = 0.5  # solo se buscan las etiquetas en el 50% superior de la pagina
CARACTERES_INVALIDOS = re.compile(r'[\\/:*?"<>|]')


# =========================================================================
# OCR: convertir el PDF a imagen y reconstruir el texto por lineas
# =========================================================================

def pdf_a_imagen(ruta_pdf: Path) -> Image.Image:
    """Convierte la primera pagina de un PDF a una imagen PIL de alta resolucion."""
    documento = fitz.open(ruta_pdf)
    pagina = documento[0]
    zoom = DPI_RENDER / 72  # 72 es el DPI base de PDF
    matriz = fitz.Matrix(zoom, zoom)
    pix = pagina.get_pixmap(matrix=matriz)
    modo = "RGB" if pix.alpha == 0 else "RGBA"
    imagen = Image.frombytes(modo, (pix.width, pix.height), pix.samples)
    documento.close()
    return imagen.convert("RGB")


def limpiar_texto_extraido(valor: str) -> str:
    """Deja el valor extraido listo para usarse en un nombre de archivo."""
    valor = valor.strip()
    valor = CARACTERES_INVALIDOS.sub("", valor)
    return valor


def _superposicion_vertical(a_top, a_bottom, b_top, b_bottom):
    inicio = max(a_top, b_top)
    fin = min(a_bottom, b_bottom)
    return max(0, fin - inicio)


def obtener_lineas_ocr(imagen: Image.Image, config: str = ""):
    """
    Corre el OCR palabra por palabra (con su posicion) y reconstruye las
    filas reales de la pagina segun la posicion vertical de cada palabra,
    en vez de usar el agrupamiento de "linea" que entrega Tesseract.

    Esto es necesario porque en facturas con varias columnas (ej. datos
    del cliente a la izquierda, datos de la factura a la derecha, en la
    misma fila visual) Tesseract a veces separa esa misma fila en dos
    "lineas" internas distintas aunque esten a la misma altura -- lo cual
    hacia que la etiqueta ("Cod. Cliente:") y su valor ("LTJ000") quedaran
    en textos separados y no se pudieran leer juntos.

    El parametro "config" se pasa directo a Tesseract (ej. "--psm 7" para
    tratar la imagen como una sola linea de texto corta -- se usa en la
    segunda pasada sobre un recorte pequeno, ver "refinar_lectura").

    Devuelve una lista de filas ordenadas de arriba hacia abajo, cada una
    con su texto, posicion (bbox), palabras individuales y confianza
    promedio del OCR.
    """
    datos = pytesseract.image_to_data(
        imagen, lang=IDIOMAS_OCR, config=config, output_type=pytesseract.Output.DICT
    )

    palabras = []
    total = len(datos["text"])
    for i in range(total):
        texto = datos["text"][i].strip()
        if not texto:
            continue
        try:
            conf = float(datos["conf"][i])
        except (TypeError, ValueError):
            conf = -1.0
        palabras.append({
            "texto": texto,
            "left": datos["left"][i],
            "top": datos["top"][i],
            "width": datos["width"][i],
            "height": datos["height"][i],
            "conf": conf,
        })

    palabras.sort(key=lambda p: p["top"])

    # Se agrupan las palabras en filas por solapamiento vertical (no por
    # coordenadas exactas), para que funcione aunque el alto de cada
    # palabra varie un poco.
    filas = []
    for p in palabras:
        p_top, p_bottom = p["top"], p["top"] + p["height"]
        fila_encontrada = None
        for fila in filas:
            alto_menor = min(p["height"], fila["bottom"] - fila["top"])
            if alto_menor <= 0:
                continue
            solapado = _superposicion_vertical(p_top, p_bottom, fila["top"], fila["bottom"])
            if solapado / alto_menor >= 0.4:
                fila_encontrada = fila
                break
        if fila_encontrada is None:
            filas.append({"top": p_top, "bottom": p_bottom, "palabras": [p]})
        else:
            fila_encontrada["palabras"].append(p)
            fila_encontrada["top"] = min(fila_encontrada["top"], p_top)
            fila_encontrada["bottom"] = max(fila_encontrada["bottom"], p_bottom)

    resultado = []
    for fila in filas:
        ps = sorted(fila["palabras"], key=lambda p: p["left"])

        # Se guarda el rango de caracteres que ocupa cada palabra dentro
        # del texto de la fila, para poder recortar despues solo las
        # palabras de una coincidencia (etiqueta + valor) en vez de toda
        # la fila del documento.
        offset = 0
        for p in ps:
            p["char_inicio"] = offset
            p["char_fin"] = offset + len(p["texto"])
            offset = p["char_fin"] + 1  # +1 por el espacio separador del join

        texto_fila = " ".join(p["texto"] for p in ps)
        top = min(p["top"] for p in ps)
        bottom = max(p["top"] + p["height"] for p in ps)
        left = min(p["left"] for p in ps)
        right = max(p["left"] + p["width"] for p in ps)
        confs = [p["conf"] for p in ps if p["conf"] >= 0]
        confianza = (sum(confs) / len(confs)) if confs else None
        resultado.append({
            "texto": texto_fila,
            "palabras": ps,
            "top": top, "bottom": bottom, "left": left, "right": right,
            "confianza": confianza,
        })

    resultado.sort(key=lambda l: l["top"])
    return resultado


def _bbox_de(linea):
    return (linea["left"], linea["top"], linea["right"], linea["bottom"])


def _bbox_de_rango(linea, inicio_char, fin_char):
    """
    Bbox de solo las palabras de la linea que caen dentro del rango de
    caracteres [inicio_char, fin_char) de linea["texto"]. Se usa para que
    el recorte de la coincidencia (ej. "Cod. Cliente: ILM001") no incluya
    el resto de la fila del documento (ej. una columna vecina como
    "# Identificacion: ...").
    """
    palabras_en_rango = [
        p for p in linea["palabras"]
        if p["char_fin"] > inicio_char and p["char_inicio"] < fin_char
    ]
    if not palabras_en_rango:
        return _bbox_de(linea)
    izquierda = min(p["left"] for p in palabras_en_rango)
    arriba = min(p["top"] for p in palabras_en_rango)
    derecha = max(p["left"] + p["width"] for p in palabras_en_rango)
    abajo = max(p["top"] + p["height"] for p in palabras_en_rango)
    return (izquierda, arriba, derecha, abajo)


def _valor_tras_posicion(linea, pos_char, min_len=3, max_len=15):
    """
    Busca, entre las palabras de la fila que empiezan en o despues de
    'pos_char' (posicion dentro de linea["texto"]), la primera que -al
    quitarle simbolos raros- quede como un codigo razonable (letras y/o
    numeros de min_len a max_len caracteres). Se usa para leer el valor
    que sigue a una etiqueta ("Cliente:", "Factura:") sin exigir que este
    pegado caracter por caracter, ya que el OCR a veces mete un simbolo
    basura de por medio (ej. "Cliente: � REMJ003").
    Devuelve (valor_limpio, palabra) o (None, None) si no se encontro nada.
    """
    for p in linea["palabras"]:
        if p["char_inicio"] < pos_char:
            continue
        candidato = re.sub(r'[^A-Za-z0-9]', '', p["texto"])
        if min_len <= len(candidato) <= max_len:
            return candidato, p
    return None, None


def buscar_factura(lineas, alto_pagina=None):
    # alto_pagina=None se usa para la segunda pasada (sobre un recorte ya
    # acotado): ahi no tiene sentido restringir a la "mitad superior".
    limite = alto_pagina * LIMITE_SUPERIOR_RATIO if alto_pagina is not None else None

    # 1) Patron directo "CMER-xxxx": muy especifico, se busca en toda la pagina.
    for linea in lineas:
        m = PATRON_CMER.search(linea["texto"])
        if m:
            valor = limpiar_texto_extraido(m.group(0))
            valor = re.sub(r'^CMER[\s\-]*', 'CMER-', valor, flags=re.IGNORECASE).upper()
            return valor, linea["confianza"], _bbox_de_rango(linea, m.start(), m.end())

    # 2) Respaldo: etiqueta "Factura" en la mitad superior del documento.
    for linea in lineas:
        if limite is not None and linea["top"] > limite:
            continue
        m = PATRON_FACTURA_ETIQUETA.search(linea["texto"])
        if m:
            valor, palabra_valor = _valor_tras_posicion(linea, m.end(), min_len=4)
            if valor:
                bbox = _bbox_de_rango(linea, m.start(), palabra_valor["char_fin"])
                return valor, linea["confianza"], bbox

    return None, None, None


def buscar_cliente(lineas, alto_pagina=None):
    limite = alto_pagina * LIMITE_SUPERIOR_RATIO if alto_pagina is not None else None
    for linea in lineas:
        if limite is not None and linea["top"] > limite:
            continue
        m = PATRON_CLIENTE_ETIQUETA.search(linea["texto"])
        if m:
            valor, palabra_valor = _valor_tras_posicion(linea, m.end())
            if valor:
                bbox = _bbox_de_rango(linea, m.start(), palabra_valor["char_fin"])
                return valor, linea["confianza"], bbox
    return None, None, None


def recorte_desde_bbox(imagen: Image.Image, bbox, margen=30):
    if bbox is None:
        return None
    left, top, right, bottom = bbox
    ancho, alto = imagen.size
    izquierda = max(0, left - margen)
    arriba = max(0, top - margen)
    derecha = min(ancho, right + margen)
    abajo = min(alto, bottom + margen)
    return imagen.crop((izquierda, arriba, derecha, abajo))


def _agrandar_recorte(recorte: Image.Image, factor=4, lado_maximo=3500):
    """
    Agranda un recorte pequeno para que Tesseract tenga mas pixeles con
    que trabajar (los codigos cortos se leen mejor asi que a tamano
    original). Si el recorte ya era grande (ej. el respaldo de "arriba de
    la pagina" cuando no se encontro nada), se limita para no agrandarlo
    de mas y volverlo lento.
    """
    ancho, alto = recorte.size
    nuevo_ancho, nuevo_alto = ancho * factor, alto * factor
    if max(nuevo_ancho, nuevo_alto) > lado_maximo:
        ajuste = lado_maximo / max(ancho, alto)
        nuevo_ancho, nuevo_alto = int(ancho * ajuste), int(alto * ajuste)
    nuevo_ancho, nuevo_alto = max(1, nuevo_ancho), max(1, nuevo_alto)
    return recorte.resize((nuevo_ancho, nuevo_alto), Image.LANCZOS)


def refinar_lectura(imagen_pagina: Image.Image, bbox, tipo: str):
    """
    Segunda pasada de OCR, mas precisa que el primer barrido de la pagina
    completa: recorta SOLO la zona pequena donde ese primer barrido
    encontro el dato (con margen), la agranda, y le vuelve a hacer OCR con
    una configuracion de Tesseract pensada para una sola linea corta de
    texto ("--psm 7"), en vez de analizar toda la pagina de una vez.

    Se hizo porque se detectaron casos donde el barrido de pagina completa
    inventaba un caracter de mas en codigos cortos (ej. leia "MCMO001" en
    vez de "MCM001", una "O" que no existe en el documento) con una
    confianza alta -- una segunda lectura enfocada solo en esa zona suele
    evitar ese tipo de error.

    Devuelve (valor, confianza) si logra leer algo valido en el recorte, o
    (None, None) si no (en ese caso se conserva el resultado del primer
    barrido, sin regresion).
    """
    recorte = recorte_desde_bbox(imagen_pagina, bbox, margen=15)
    if recorte is None:
        return None, None

    recorte_grande = _agrandar_recorte(recorte)
    lineas_recorte = obtener_lineas_ocr(recorte_grande, config="--psm 7")
    if not lineas_recorte:
        return None, None

    if tipo == "cliente":
        valor, confianza, _bbox = buscar_cliente(lineas_recorte, alto_pagina=None)
    else:
        valor, confianza, _bbox = buscar_factura(lineas_recorte, alto_pagina=None)

    return valor, confianza


def cargar_codigos_confirmados():
    """Devuelve el conjunto de codigos de cliente ya confirmados a mano alguna vez."""
    if not CODIGOS_CONFIRMADOS_PATH.exists():
        return set()
    with open(CODIGOS_CONFIRMADOS_PATH, "r", encoding="utf-8") as f:
        return {linea.strip() for linea in f if linea.strip()}


def agregar_codigo_confirmado(codigo):
    """Agrega un codigo a la lista de confirmados, si no estaba ya."""
    if not codigo:
        return
    CARPETA_PROYECTO.mkdir(parents=True, exist_ok=True)
    if codigo in cargar_codigos_confirmados():
        return
    with open(CODIGOS_CONFIRMADOS_PATH, "a", encoding="utf-8") as f:
        f.write(codigo + "\n")


def sembrar_codigos_confirmados_desde_log():
    """
    La primera vez que se usa esta funcion (todavia no existe el archivo
    de codigos confirmados), se aprovechan las correcciones manuales que
    ya quedaron guardadas en ejecuciones anteriores del log, en vez de
    empezar la lista desde cero.
    """
    if CODIGOS_CONFIRMADOS_PATH.exists():
        return
    codigos = set()
    for fila in _leer_log():
        if fila.get("metodo") == "manual" and fila.get("estado") == "OK" and fila.get("cliente"):
            codigos.add(fila["cliente"].strip())
    if codigos:
        CARPETA_PROYECTO.mkdir(parents=True, exist_ok=True)
        with open(CODIGOS_CONFIRMADOS_PATH, "w", encoding="utf-8") as f:
            for codigo in sorted(codigos):
                f.write(codigo + "\n")


# Unicos caracteres que se permite insertar/quitar/sustituir "gratis" al
# comparar contra la lista de confirmados. Con una lista grande (miles de
# codigos reales), permitir CUALQUIER sustitucion de un solo caracter es
# peligroso: por ejemplo, "JAS003" (bien leido) esta a una sola letra de
# "AAS003" (otro cliente real que si esta en la lista), y con una
# distancia de edicion normal se "corregiria" a ese OTRO cliente -- error
# real que se detecto en pruebas. Por eso aqui solo se permiten, a costo
# cero, los cambios que sabemos que comete el OCR en este tipo de codigos
# (confundir O con 0, o I con 1, o meter/quitar una de esas letras de
# mas). Cualquier otro cambio (otra letra, otro numero) NO se permite.
_EQUIVALENTES_OCR = {("O", "0"), ("0", "O"), ("I", "1"), ("1", "I")}
_CARACTERES_AMBIGUOS_OCR = {"O", "0", "I", "1"}


def _distancia_restringida(a, b):
    """
    Como una distancia de edicion, pero SOLO permite (a costo 0) los
    cambios tipicos de la confusion O/0 e I/1 del OCR -- insertar, quitar
    o sustituir unicamente esos caracteres. Cualquier otro cambio (otra
    letra distinta, otro digito distinto) cuesta "infinito", es decir, no
    esta permitido. Si el resultado es 0, quiere decir que "a" y "b" son
    el mismo codigo, solo que con esas confusiones tipicas del OCR de por
    medio.
    """
    a, b = a.upper(), b.upper()
    m, n = len(a), len(b)
    INFINITO = float("inf")

    def costo_insertar_o_quitar(caracter):
        return 0 if caracter in _CARACTERES_AMBIGUOS_OCR else INFINITO

    def costo_sustituir(ca, cb):
        if ca == cb or (ca, cb) in _EQUIVALENTES_OCR:
            return 0
        return INFINITO

    fila_anterior = [0] * (n + 1)
    for j in range(1, n + 1):
        fila_anterior[j] = fila_anterior[j - 1] + costo_insertar_o_quitar(b[j - 1])

    for i in range(1, m + 1):
        fila_actual = [fila_anterior[0] + costo_insertar_o_quitar(a[i - 1])] + [0] * n
        for j in range(1, n + 1):
            fila_actual[j] = min(
                fila_anterior[j] + costo_insertar_o_quitar(a[i - 1]),      # quitar de "a"
                fila_actual[j - 1] + costo_insertar_o_quitar(b[j - 1]),    # falta en "a"
                fila_anterior[j - 1] + costo_sustituir(a[i - 1], b[j - 1]),
            )
        fila_anterior = fila_actual

    return fila_anterior[n]


def corregir_con_lista_confirmados(valor, codigos_confirmados):
    """
    Si "valor" (lo que leyo el OCR) es el MISMO codigo que uno ya
    confirmado a mano antes, solo que con una confusion tipica de O/0 o
    I/1 de por medio, se corrige a ese codigo confirmado. No corrige ante
    cualquier otra diferencia (otra letra, otro numero), ni si hay dos
    codigos confirmados que calzan igual (para no adivinar entre dos
    clientes reales parecidos).
    Devuelve (valor_final, se_corrigio).
    """
    if not valor or not codigos_confirmados:
        return valor, False
    if valor in codigos_confirmados:
        return valor, False

    candidatos = [codigo for codigo in codigos_confirmados if _distancia_restringida(valor, codigo) == 0]
    if len(candidatos) != 1:
        return valor, False  # ninguno calza, o hay mas de uno: no se adivina

    return candidatos[0], True


def extraer_datos(imagen: Image.Image):
    """Devuelve un diccionario con cliente, factura, su confianza OCR y su bbox (para recortes)."""
    ancho, alto = imagen.size
    lineas = obtener_lineas_ocr(imagen)
    factura, conf_factura, bbox_factura = buscar_factura(lineas, alto)
    cliente, conf_cliente, bbox_cliente = buscar_cliente(lineas, alto)

    # Segunda pasada enfocada: si el barrido de la pagina completa encontro
    # una zona aproximada, se vuelve a leer SOLO esa zona (agrandada) para
    # confirmar o corregir el valor. Ver "refinar_lectura" para el porque.
    if bbox_cliente is not None:
        valor_refinado, confianza_refinada = refinar_lectura(imagen, bbox_cliente, "cliente")
        if valor_refinado:
            cliente, conf_cliente = valor_refinado, confianza_refinada

    if bbox_factura is not None:
        valor_refinado, confianza_refinada = refinar_lectura(imagen, bbox_factura, "factura")
        if valor_refinado:
            factura, conf_factura = valor_refinado, confianza_refinada

    # Tercer filtro: si el codigo de cliente leido esta muy cerca de uno
    # que ya se confirmo a mano antes (ej. para el mismo cliente en una
    # factura anterior), se corrige solo. Esto ataca directamente el caso
    # de "MCMO001" vs "MCM001": una vez confirmado una vez, las siguientes
    # facturas de ese mismo cliente se corrigen automaticamente.
    if cliente:
        codigos_confirmados = cargar_codigos_confirmados()
        cliente_corregido, se_corrigio = corregir_con_lista_confirmados(cliente, codigos_confirmados)
        if se_corrigio:
            cliente = cliente_corregido
            conf_cliente = 100.0  # coincide con un codigo ya confirmado a mano

    return {
        "cliente": cliente, "conf_cliente": conf_cliente, "bbox_cliente": bbox_cliente,
        "factura": factura, "conf_factura": conf_factura, "bbox_factura": bbox_factura,
    }


def obtener_recortes_para_revision(imagen: Image.Image, datos: dict):
    """Devuelve (recorte_codigo_cliente, recorte_factura) para la pantalla manual."""
    ancho, alto = imagen.size
    recorte_cliente = recorte_desde_bbox(imagen, datos["bbox_cliente"])
    recorte_factura = recorte_desde_bbox(imagen, datos["bbox_factura"])

    # Respaldo: si no se ubico ni siquiera la etiqueta, se muestra la
    # parte superior del documento (donde normalmente estan estos datos).
    if recorte_cliente is None:
        recorte_cliente = imagen.crop((0, 0, ancho, int(alto * 0.35)))
    if recorte_factura is None:
        recorte_factura = imagen.crop((0, 0, ancho, int(alto * 0.35)))

    return recorte_cliente, recorte_factura


# =========================================================================
# PANTALLA DE REVISION MANUAL
# =========================================================================

class VentanaRevisionManual(tk.Toplevel):
    """
    Pantalla que aparece solo cuando el programa NO pudo extraer el codigo
    de cliente y/o la factura de un PDF. Muestra los recortes de la
    imagen para que la gerente los lea a simple vista y los digite.
    """

    ANCHO_MAX_IMAGEN = 650
    ALTO_MAX_IMAGEN = 200

    def __init__(self, master, nombre_archivo, recorte_cliente, recorte_factura,
                 codigo_cliente_previo, factura_previa):
        super().__init__(master)
        self.title("Revision manual necesaria - RENOMBRADOR")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._on_cancelar)  # cerrar con la X = Cancelar

        self.resultado = None  # ("aceptar", codigo, factura) o ("cancelar", None, None)

        contenedor = ttk.Frame(self, padding=16)
        contenedor.pack(fill="both", expand=True)

        ttk.Label(
            contenedor,
            text=f"No se pudieron leer todos los datos de:\n{nombre_archivo}",
            font=("Segoe UI", 11, "bold"),
            justify="center",
        ).pack(pady=(0, 12))

        # --- Recorte / campo: Codigo cliente ---
        ttk.Label(contenedor, text="Recorte con el codigo de cliente:").pack(anchor="w")
        self._foto_cliente = self._preparar_foto(recorte_cliente)
        ttk.Label(contenedor, image=self._foto_cliente, relief="solid").pack(pady=(2, 8))

        ttk.Label(contenedor, text="Codigo de cliente:").pack(anchor="w")
        self.campo_cliente = ttk.Entry(contenedor, width=40)
        self.campo_cliente.pack(fill="x", pady=(0, 14))
        if codigo_cliente_previo:
            self.campo_cliente.insert(0, codigo_cliente_previo)

        # --- Recorte / campo: Factura ---
        ttk.Label(contenedor, text="Recorte con el numero de factura:").pack(anchor="w")
        self._foto_factura = self._preparar_foto(recorte_factura)
        ttk.Label(contenedor, image=self._foto_factura, relief="solid").pack(pady=(2, 8))

        ttk.Label(contenedor, text="Numero de factura:").pack(anchor="w")
        self.campo_factura = ttk.Entry(contenedor, width=40)
        self.campo_factura.pack(fill="x", pady=(0, 14))
        if factura_previa:
            self.campo_factura.insert(0, factura_previa)

        self.mensaje_error = ttk.Label(contenedor, text="", foreground="red")
        self.mensaje_error.pack()

        # --- Botones ---
        fila_botones = ttk.Frame(contenedor)
        fila_botones.pack(pady=(8, 0))

        ttk.Button(fila_botones, text="Cancelar (omitir este archivo)",
                   command=self._on_cancelar).pack(side="left", padx=6)
        ttk.Button(fila_botones, text="Aceptar",
                   command=self._on_aceptar).pack(side="left", padx=6)

        self.grab_set()  # bloquea el resto del programa hasta que se responda
        self.campo_cliente.focus_set()

    def _preparar_foto(self, imagen_pil: Image.Image):
        copia = imagen_pil.copy()
        copia.thumbnail((self.ANCHO_MAX_IMAGEN, self.ALTO_MAX_IMAGEN))
        return ImageTk.PhotoImage(copia)

    def _on_aceptar(self):
        codigo = self.campo_cliente.get().strip()
        factura = self.campo_factura.get().strip()

        # Los dos campos son obligatorios antes de aceptar.
        if not codigo or not factura:
            self.mensaje_error.config(text="Ambos campos son obligatorios.")
            return

        codigo = limpiar_texto_extraido(codigo)
        factura = limpiar_texto_extraido(factura)
        self.resultado = ("aceptar", codigo, factura)
        self.destroy()

    def _on_cancelar(self):
        # CAMBIAR AQUI si prefiere que "Cancelar" detenga todo el programa
        # en vez de solo omitir el archivo actual.
        self.resultado = ("cancelar", None, None)
        self.destroy()


# =========================================================================
# LOG (JSON) - un registro por cada archivo procesado, automatico o manual
# =========================================================================

def _leer_log():
    """Devuelve la lista de registros del log, o [] si no existe o esta vacio/dañado."""
    if not LOG_PATH.exists():
        return []
    with open(LOG_PATH, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _escribir_log(registros):
    CARPETA_PROYECTO.mkdir(parents=True, exist_ok=True)
    with open(LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(registros, f, ensure_ascii=False, indent=2)


def migrar_log_csv_a_json():
    """
    La primera vez que se usa el log en formato .json (todavia no existe),
    si existe el .csv de antes (formato anterior) se aprovecha ese
    historial en vez de empezar de cero.
    """
    if LOG_PATH.exists() or not LOG_PATH_CSV_ANTIGUO.exists():
        return
    registros = []
    with open(LOG_PATH_CSV_ANTIGUO, "r", encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            confianza_txt = fila.get("confianza_ocr_pct", "")
            try:
                confianza = float(confianza_txt) if confianza_txt else None
            except ValueError:
                confianza = None
            registros.append({
                "fecha_hora": fila.get("fecha_hora", ""),
                "archivo_original": fila.get("archivo_original", ""),
                "archivo_nuevo": fila.get("archivo_nuevo", ""),
                "cliente": fila.get("cliente", ""),
                "factura": fila.get("factura", ""),
                "metodo": fila.get("metodo", ""),
                "confianza_ocr_pct": confianza,
                "estado": fila.get("estado", ""),
            })
    if registros:
        _escribir_log(registros)


def registrar_log(archivo_original, archivo_nuevo, cliente, factura, metodo, confianza, estado):
    registros = _leer_log()
    registros.append({
        "fecha_hora": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        "archivo_original": archivo_original,
        "archivo_nuevo": archivo_nuevo or "",
        "cliente": cliente or "",
        "factura": factura or "",
        "metodo": metodo,
        "confianza_ocr_pct": round(confianza, 1) if confianza is not None else None,
        "estado": estado,
    })
    _escribir_log(registros)


# =========================================================================
# PROGRAMA PRINCIPAL
# =========================================================================

def preparar_carpetas():
    CARPETA_PROYECTO.mkdir(parents=True, exist_ok=True)
    CARPETA_POR_REVISAR.mkdir(parents=True, exist_ok=True)
    CARPETA_RESULTADO_HOY.mkdir(parents=True, exist_ok=True)
    migrar_log_csv_a_json()
    sembrar_codigos_confirmados_desde_log()

    if not CARPETA_ENTRADA.exists():
        CARPETA_ENTRADA.mkdir(parents=True, exist_ok=True)
        messagebox.showinfo(
            "Carpeta PRUEBA creada",
            f"Se creo la carpeta:\n{CARPETA_ENTRADA}\n\n"
            "Coloque ahi los PDF escaneados y vuelva a ejecutar el programa."
        )
        sys.exit(0)


def nombre_disponible(carpeta: Path, nombre_base: str) -> str:
    """Evita sobreescribir si ya existe un archivo con ese nombre."""
    destino = carpeta / f"{nombre_base}.pdf"
    if not destino.exists():
        return destino.name
    contador = 2
    while (carpeta / f"{nombre_base} ({contador}).pdf").exists():
        contador += 1
    return f"{nombre_base} ({contador}).pdf"


def guardar_resultado(ruta_origen: Path, cliente: str, factura: str) -> str:
    nombre_base = f"{cliente}-{factura}"
    nombre_final = nombre_disponible(CARPETA_RESULTADO_HOY, nombre_base)
    destino = CARPETA_RESULTADO_HOY / nombre_final
    shutil.copy2(ruta_origen, destino)  # copia (no borra el original de PRUEBA)
    return nombre_final


def fase_1_automatica():
    """Revisa PRUEBA uno por uno: lo que se puede leer solo, se guarda ya renombrado."""
    archivos_pdf = sorted(CARPETA_ENTRADA.glob("*.pdf")) + sorted(CARPETA_ENTRADA.glob("*.PDF"))
    archivos_pdf = sorted(set(archivos_pdf))

    total_ok = 0
    total_enviados_a_revisar = 0

    for ruta_pdf in archivos_pdf:
        try:
            imagen = pdf_a_imagen(ruta_pdf)
            datos = extraer_datos(imagen)
        except Exception as error:
            # No se pudo ni abrir/leer el PDF: se manda igual a revision manual
            destino = CARPETA_POR_REVISAR / ruta_pdf.name
            shutil.copy2(ruta_pdf, destino)
            registrar_log(ruta_pdf.name, None, None, None, "automatico",
                          None, f"ERROR AL LEER ({error})")
            total_enviados_a_revisar += 1
            continue

        confianza = None
        confs = [c for c in (datos["conf_cliente"], datos["conf_factura"]) if c is not None]
        if confs:
            confianza = sum(confs) / len(confs)

        # Ademas de encontrar los dos datos, se exige una confianza minima
        # del OCR para aceptarlos sin revision. Se vio con facturas reales
        # que el OCR a veces SI encuentra un codigo, pero lo lee mal (ej.
        # confunde "I" con "1", o mete una "O" de mas) -- eso se refleja
        # en una confianza mas baja de lo normal.
        confianza_suficiente = (
            (datos["conf_cliente"] is None or datos["conf_cliente"] >= UMBRAL_CONFIANZA_MINIMA)
            and (datos["conf_factura"] is None or datos["conf_factura"] >= UMBRAL_CONFIANZA_MINIMA)
        )

        if datos["cliente"] and datos["factura"] and confianza_suficiente:
            nombre_final = guardar_resultado(ruta_pdf, datos["cliente"], datos["factura"])
            registrar_log(ruta_pdf.name, nombre_final, datos["cliente"], datos["factura"],
                          "automatico", confianza, "OK")
            total_ok += 1
        else:
            destino = CARPETA_POR_REVISAR / ruta_pdf.name
            shutil.copy2(ruta_pdf, destino)
            if datos["cliente"] and datos["factura"]:
                # Se encontraron los datos pero con baja confianza: se deja
                # anotado en el log para que quede claro por que se mando
                # a revision manual (no es que no se pudiera leer nada).
                registrar_log(ruta_pdf.name, None, datos["cliente"], datos["factura"],
                              "automatico", confianza, "BAJA CONFIANZA - a revision manual")
            total_enviados_a_revisar += 1

    return total_ok, total_enviados_a_revisar


def fase_2_manual(raiz: tk.Tk):
    """Revisa, uno por uno, lo que quedo en POR_REVISAR (de esta ejecucion o de antes)."""
    total_manual_ok = 0
    total_omitidos = 0

    pendientes = sorted(CARPETA_POR_REVISAR.glob("*.pdf")) + sorted(CARPETA_POR_REVISAR.glob("*.PDF"))
    pendientes = sorted(set(pendientes))

    for ruta_pdf in pendientes:
        try:
            imagen = pdf_a_imagen(ruta_pdf)
            datos = extraer_datos(imagen)
        except Exception:
            imagen = None
            datos = {"cliente": None, "factura": None, "bbox_cliente": None, "bbox_factura": None}

        if imagen is not None:
            recorte_cliente, recorte_factura = obtener_recortes_para_revision(imagen, datos)
        else:
            # No se pudo ni renderizar la imagen: se muestran recortes en blanco
            recorte_cliente = Image.new("RGB", (600, 150), "white")
            recorte_factura = Image.new("RGB", (600, 150), "white")

        ventana = VentanaRevisionManual(
            raiz, ruta_pdf.name, recorte_cliente, recorte_factura,
            datos.get("cliente"), datos.get("factura"),
        )
        raiz.wait_window(ventana)

        if ventana.resultado is None or ventana.resultado[0] == "cancelar":
            registrar_log(ruta_pdf.name, None, datos.get("cliente"), datos.get("factura"),
                          "manual", None, "OMITIDO (queda en POR_REVISAR)")
            total_omitidos += 1
            continue

        _, cliente, factura = ventana.resultado
        agregar_codigo_confirmado(cliente)  # queda disponible para corregir futuras lecturas de este cliente
        nombre_final = guardar_resultado(ruta_pdf, cliente, factura)
        registrar_log(ruta_pdf.name, nombre_final, cliente, factura, "manual", None, "OK")
        ruta_pdf.unlink()  # se borra de POR_REVISAR porque ya quedo guardado en RESULTADO
        total_manual_ok += 1

    return total_manual_ok, total_omitidos


def mostrar_resumen(total_ok, total_manual_ok, total_omitidos):
    mensaje = (
        f"Proceso terminado.\n\n"
        f"Renombrados automaticamente: {total_ok}\n"
        f"Renombrados con revision manual: {total_manual_ok}\n"
        f"Omitidos (siguen en POR_REVISAR): {total_omitidos}\n\n"
        f"Carpeta de resultado de hoy:\n{CARPETA_RESULTADO_HOY}\n\n"
        f"Detalle de esta y otras ejecuciones:\n{LOG_PATH}"
    )
    messagebox.showinfo("Resumen del proceso", mensaje)


def main():
    raiz = tk.Tk()
    raiz.withdraw()  # no se muestra una ventana principal vacia, solo las de revision manual

    preparar_carpetas()

    total_ok, enviados_a_revisar = fase_1_automatica()
    total_manual_ok, total_omitidos = fase_2_manual(raiz)

    mostrar_resumen(total_ok, total_manual_ok, total_omitidos)
    raiz.destroy()


if __name__ == "__main__":
    main()
