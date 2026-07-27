# -*- coding: utf-8 -*-
"""
============================================================================
CASCADAS HOTEL - AUDITORIA DIARIA
============================================================================
Herramienta local para el cierre diario de boletas y facturas.

Corre en el computador de recepcion: un servidor pequeno en Python que hace
el trabajo con el escaner y los archivos, y una pantalla en el navegador.

Cubre el cierre completo:
  - escanear desde el Canon (WIA), eligiendo tipo y numero de documento
  - agregar paginas al mismo documento (voucher, transferencia, pasaporte)
  - archivar en la carpeta del dia: JPEG si es una pagina, PDF si son varias
  - la grilla de valores, con validaciones antes de cerrar
  - escribir la hoja del dia en la planilla de auditoria
  - abrir el correo a contabilidad con todo adjunto
  - buscar documentos de cualquier dia por su numero

El avance del dia se guarda solo: si se apaga el computador, al volver a
abrir esta todo donde estaba.

Se ejecuta con:  python auditoria.py
Para probar la pantalla sin escaner:  python auditoria.py --demo
============================================================================
"""

import datetime
import http.server
import io
import json
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import threading
import traceback
import unicodedata
import webbrowser

# ---------------------------------------------------------------------------
# Pillow es obligatorio (arma los PDF). pywin32 solo en Windows.
# ---------------------------------------------------------------------------
try:
    from PIL import Image, ImageDraw
except ImportError:
    print("\nFalta la libreria Pillow. Abre PowerShell y ejecuta:\n")
    print("    pip install pillow pywin32\n")
    sys.exit(1)

MODO_DEMO = "--demo" in sys.argv
ES_WINDOWS = os.name == "nt"

CARPETA_APP = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CONFIG = os.path.join(CARPETA_APP, "config.json")
CARPETA_TRABAJO = os.path.join(CARPETA_APP, "trabajo")

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

TIPOS = {
    "boleta": {"etiqueta": "Boleta", "exento": False},
    "factura": {"etiqueta": "Factura", "exento": False},
    "exportacion": {"etiqueta": "Factura de exportación", "exento": True},
}

CONFIG_POR_DEFECTO = {
    "carpeta_base": "",       # se busca sola la primera vez (ver buscar_carpeta_escaneados)
    "archivo_excel": "",      # la planilla de auditoria, tambien se busca sola
    "resolucion": 200,
    "color": "gris",          # 'gris' o 'color'
    "calidad_jpeg": 65,
    # Formas de pago sugeridas. Se puede escribir cualquier otra: las nuevas se
    # van sumando solas a esta lista para tenerlas a mano el dia siguiente.
    "formas_pago": ["EFECTIVO", "DB", "CREDITO", "MC", "VISA", "VISA USD",
                    "TRANSFERENCIA", "MC Y TRANSFERENCIA"],
    # --- Correo a contabilidad ---
    "correo_para": "jisla@hotelantofagasta.cl",
    "correo_cc": "npizarro@cascadashotel.cl; administracion@cascadashotel.cl; "
                 "contabilidad@hotelantofagasta.cl",
    "correo_saludo": "Estimada Josefa",
    # Quien firma el correo. La lista es del hotel; el elegido queda por
    # computador, porque no siempre audita la misma persona.
    "remitentes": [],
}

# Prefijo con que se nombran los adjuntos del correo, como los recibe
# contabilidad hoy: "B 137.854.jpeg".
PREFIJO_ADJUNTO = {"boleta": "B", "factura": "F", "exportacion": "FE"}

# ---------------------------------------------------------------------------
# Planilla de auditoria: encabezados en la fila 3, datos desde la 4,
# columnas C a J. El TOTAL de la planilla es el monto AFECTO (neto + IVA):
# la propina va exenta y se controla por otro lado, no entra aqui.
# ---------------------------------------------------------------------------
EXCEL_FILA_DATOS = 4
COL_TIPO, COL_FOLIO, COL_NETO, COL_IVA = 3, 4, 5, 6         # C, D, E, F
COL_TOTAL, COL_PAGO, COL_AUTORIZACION, COL_OBS = 7, 8, 9, 10  # G, H, I, J

# Como se escribe cada tipo en la columna TIPO DE DOCUMENTO.
TIPO_EN_PLANILLA = {
    "boleta": "BOLETA",
    "factura": "FACTURA",
    "exportacion": "FACTURA EXPORTACION",
}

IVA_TASA = 0.19


# ===========================================================================
#  CONFIGURACION
# ===========================================================================
def leer_config():
    cfg = dict(CONFIG_POR_DEFECTO)
    if os.path.exists(ARCHIVO_CONFIG):
        try:
            with open(ARCHIVO_CONFIG, "r", encoding="utf-8") as f:
                cfg.update(json.load(f) or {})
        except Exception:
            pass
    # Una version anterior guardo mal el correo de contabilidad. Como el ajuste
    # ya esta en el disco, hay que corregirlo tambien ahi.
    if "cascadasantofagasta.cl" in (cfg.get("correo_cc") or ""):
        cfg["correo_cc"] = cfg["correo_cc"].replace("cascadasantofagasta.cl", "hotelantofagasta.cl")
        try:
            guardar_config(cfg)
        except Exception:
            pass
    return cfg


def guardar_config(cfg):
    with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ===========================================================================
#  ENCONTRAR LA CARPETA DE ESCANEADOS
#  La ruta lleva enes y tildes y esta en el Escritorio, que ademas puede
#  estar redirigido a OneDrive. Escribirla a mano es la parte mas facil de
#  equivocar, asi que la herramienta la busca sola la primera vez.
# ===========================================================================
def _raices_probables():
    raices = []
    perfil = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    for sub in ("Desktop", "Escritorio", "OneDrive\\Desktop", "OneDrive\\Escritorio",
                "Documents", "Documentos"):
        raices.append(os.path.join(perfil, sub))
    raices.append(perfil)
    one = os.environ.get("OneDrive") or os.environ.get("OneDriveConsumer")
    if one:
        raices.extend([os.path.join(one, "Desktop"), os.path.join(one, "Escritorio"), one])
    return [r for r in raices if os.path.isdir(r)]


def _parece_carpeta_escaneados(nombre):
    n = _sin_tildes(nombre)
    return "escaneado" in n and ("boleta" in n or "factura" in n)


def buscar_carpeta_escaneados():
    """Devuelve la carpeta de escaneados si logra reconocerla, o cadena vacia.
    Solo mira unos pocos niveles: no recorre el disco entero."""
    for raiz in _raices_probables():
        for base, carpetas, _archivos in os.walk(raiz):
            profundidad = base[len(raiz):].count(os.sep)
            if profundidad >= 3:
                carpetas[:] = []
                continue
            carpetas[:] = [c for c in carpetas if not c.startswith(".")]
            for carpeta in carpetas:
                if _parece_carpeta_escaneados(carpeta):
                    return os.path.join(base, carpeta)
    return ""


def buscar_planilla_auditoria(pistas=()):
    """Ubica la planilla de auditoria (.xlsx) sin recorrer el disco entero."""
    raices = [p for p in pistas if p and os.path.isdir(p)]
    for p in list(raices):
        padre = os.path.dirname(p)
        if padre and os.path.isdir(padre):
            raices.append(padre)
    raices.extend(_raices_probables())
    for raiz in raices:
        for base, carpetas, archivos in os.walk(raiz):
            if base[len(raiz):].count(os.sep) >= 2:
                carpetas[:] = []
                continue
            carpetas[:] = [c for c in carpetas if not c.startswith(".")]
            for archivo in archivos:
                if archivo.startswith("~$") or not archivo.lower().endswith((".xlsx", ".xlsm")):
                    continue
                if "auditoria" in _sin_tildes(archivo):
                    return os.path.join(base, archivo)
    return ""


def elegir_archivo_con_ventana(inicial=""):
    """Abre el selector de archivos de Windows para elegir la planilla."""
    codigo = (
        "import sys, tkinter as tk\n"
        "from tkinter import filedialog\n"
        "r = tk.Tk(); r.withdraw(); r.attributes('-topmost', True)\n"
        "ruta = filedialog.askopenfilename(title='Elige la planilla de auditoria', "
        "initialdir=sys.argv[1] if len(sys.argv) > 1 else '', "
        "filetypes=[('Planillas de Excel', '*.xlsx *.xlsm'), ('Todos', '*.*')])\n"
        "sys.stdout.write(ruta or '')\n"
    )
    try:
        salida = subprocess.run([sys.executable, "-c", codigo, inicial or ""],
                                capture_output=True, timeout=180)
        return salida.stdout.decode("utf-8", "replace").strip()
    except Exception:
        traceback.print_exc()
        return ""


def elegir_carpeta_con_ventana(inicial=""):
    """Abre el buscador de carpetas de Windows. Se lanza como proceso aparte
    para no mezclar la ventana con el servidor."""
    codigo = (
        "import sys, tkinter as tk\n"
        "from tkinter import filedialog\n"
        "r = tk.Tk(); r.withdraw(); r.attributes('-topmost', True)\n"
        "ruta = filedialog.askdirectory(title='Elige la carpeta de escaneados', "
        "initialdir=sys.argv[1] if len(sys.argv) > 1 else '')\n"
        "sys.stdout.write(ruta or '')\n"
    )
    try:
        salida = subprocess.run([sys.executable, "-c", codigo, inicial or ""],
                                capture_output=True, timeout=180)
        return salida.stdout.decode("utf-8", "replace").strip()
    except Exception:
        traceback.print_exc()
        return ""


# ===========================================================================
#  CARPETA DEL DIA
#  Se respeta la forma de escribir que ya existe en el disco: si la carpeta
#  del mes esta como "julio", no se crea otra "Julio" al lado.
# ===========================================================================
def _sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn").lower().strip()


def _subcarpeta(padre, nombre_por_defecto, alternativas=(), crear=True):
    """Busca una subcarpeta que ya represente lo mismo; si no, usa el nombre
    por defecto. Devuelve la ruta (creandola solo si `crear`)."""
    buscados = [_sin_tildes(nombre_por_defecto)] + [_sin_tildes(a) for a in alternativas]
    if os.path.isdir(padre):
        try:
            for entrada in sorted(os.listdir(padre)):
                ruta = os.path.join(padre, entrada)
                if os.path.isdir(ruta) and _sin_tildes(entrada) in buscados:
                    return ruta
        except OSError:
            pass
    ruta = os.path.join(padre, nombre_por_defecto)
    if crear:
        os.makedirs(ruta, exist_ok=True)
    return ruta


def carpeta_del_dia(carpeta_base, fecha, crear=True):
    anio = _subcarpeta(carpeta_base, str(fecha.year), crear=crear)
    mes = _subcarpeta(anio, MESES[fecha.month - 1], crear=crear)
    # El dia puede estar como "27", "7" o "27-07-2026": se acepta cualquiera.
    dia = _subcarpeta(mes, "%02d" % fecha.day,
                      alternativas=(str(fecha.day), fecha.strftime("%d-%m-%Y"), fecha.strftime("%Y-%m-%d")),
                      crear=crear)
    return dia


# ===========================================================================
#  ESTADO DEL DIA  (se guarda en disco a cada cambio)
# ===========================================================================
class EstadoDia(object):
    def __init__(self, fecha):
        self.fecha = fecha
        self.carpeta_trabajo = os.path.join(CARPETA_TRABAJO, fecha.isoformat())
        os.makedirs(self.carpeta_trabajo, exist_ok=True)
        self.archivo = os.path.join(self.carpeta_trabajo, "estado.json")
        self.lock = threading.Lock()
        self.documentos = []
        self.siguiente_id = 1
        self._cargar()

    def _cargar(self):
        if not os.path.exists(self.archivo):
            return
        try:
            with open(self.archivo, "r", encoding="utf-8") as f:
                datos = json.load(f)
            self.documentos = datos.get("documentos", [])
            self.siguiente_id = datos.get("siguienteId", len(self.documentos) + 1)
        except Exception:
            traceback.print_exc()

    def guardar(self):
        tmp = self.archivo + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"documentos": self.documentos, "siguienteId": self.siguiente_id},
                      f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.archivo)   # reemplazo atomico: nunca queda a medias

    def buscar(self, doc_id):
        for d in self.documentos:
            if d["id"] == doc_id:
                return d
        return None


# ===========================================================================
#  ESCANER (WIA)
#  Todas las llamadas al escaner pasan por un unico hilo: COM no se lleva
#  bien con que varios hilos lo usen a la vez.
# ===========================================================================
WIA_INTENT = 6146          # tipo de imagen deseada
WIA_XRES, WIA_YRES = 6147, 6148
WIA_XPOS, WIA_YPOS = 6149, 6150
WIA_XEXT, WIA_YEXT = 6151, 6152

INTENT_COLOR = 1
INTENT_GRIS = 2

FORMATO_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
FORMATO_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
FORMATO_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"

# Medidas del vidrio en pulgadas (carta / A4). Se usan solo si el driver no
# informa su propio maximo.
ANCHO_PULGADAS = 8.5
ALTO_PULGADAS = 11.7


class ErrorEscaner(Exception):
    pass


class Escaner(object):
    """Envuelve WIA. En modo demo genera una hoja de prueba."""

    def __init__(self):
        self.peticiones = queue.Queue()
        self.hilo = threading.Thread(target=self._bucle, daemon=True)
        self.hilo.start()

    # -- API publica (se llama desde los hilos del servidor) --
    def _pedir(self, accion, *args):
        respuesta = queue.Queue(1)
        self.peticiones.put((accion, args, respuesta))
        ok, valor = respuesta.get()
        if not ok:
            raise ErrorEscaner(valor)
        return valor

    def listar(self):
        return self._pedir("listar")

    def escanear(self, destino, resolucion, color):
        return self._pedir("escanear", destino, resolucion, color)

    # -- Hilo dedicado --
    def _bucle(self):
        com_iniciado = False
        if ES_WINDOWS and not MODO_DEMO:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                com_iniciado = True
            except Exception:
                traceback.print_exc()
        try:
            while True:
                accion, args, respuesta = self.peticiones.get()
                try:
                    if accion == "listar":
                        respuesta.put((True, self._listar()))
                    elif accion == "escanear":
                        respuesta.put((True, self._escanear(*args)))
                    else:
                        respuesta.put((False, "Acción desconocida: %s" % accion))
                except Exception as e:
                    traceback.print_exc()
                    respuesta.put((False, str(e) or e.__class__.__name__))
        finally:
            if com_iniciado:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _listar(self):
        if MODO_DEMO or not ES_WINDOWS:
            return [{"id": "demo", "nombre": "Escaner simulado (modo demo)"}]
        import win32com.client
        gestor = win32com.client.Dispatch("WIA.DeviceManager")
        equipos = []
        for info in gestor.DeviceInfos:
            try:
                if int(info.Type) != 1:      # 1 = escaner
                    continue
                nombre = ""
                for p in info.Properties:
                    if p.Name == "Name":
                        nombre = str(p.Value)
                equipos.append({"id": str(info.DeviceID), "nombre": nombre or "Escáner"})
            except Exception:
                continue
        return equipos

    def _conectar(self, win32com_client):
        gestor = win32com_client.Dispatch("WIA.DeviceManager")
        for info in gestor.DeviceInfos:
            try:
                if int(info.Type) == 1:
                    return info.Connect()
            except Exception:
                continue
        raise ErrorEscaner("No se encontró ningún escáner conectado. "
                           "Revisa que esté encendido y que el cable USB esté puesto.")

    def _escanear(self, destino, resolucion, color):
        if MODO_DEMO or not ES_WINDOWS:
            return self._escanear_demo(destino, resolucion)

        import win32com.client
        dispositivo = self._conectar(win32com.client)
        item = dispositivo.Items[1]

        def poner(pid, valor):
            """Ajusta una propiedad si el driver la admite; si no, sigue."""
            try:
                for p in item.Properties:
                    if int(p.PropertyID) == pid:
                        p.Value = valor
                        return True
            except Exception:
                pass
            return False

        def maximo(pid, por_defecto):
            try:
                for p in item.Properties:
                    if int(p.PropertyID) == pid:
                        return int(p.SubTypeMax)
            except Exception:
                pass
            return por_defecto

        poner(WIA_INTENT, INTENT_COLOR if color == "color" else INTENT_GRIS)
        poner(WIA_XRES, int(resolucion))
        poner(WIA_YRES, int(resolucion))
        # El area hay que fijarla DESPUES de la resolucion: varios drivers la
        # dejan del tamano anterior y el escaneo sale recortado.
        poner(WIA_XPOS, 0)
        poner(WIA_YPOS, 0)
        poner(WIA_XEXT, maximo(WIA_XEXT, int(resolucion * ANCHO_PULGADAS)))
        poner(WIA_YEXT, maximo(WIA_YEXT, int(resolucion * ALTO_PULGADAS)))

        imagen = None
        ultimo_error = None
        for formato in (FORMATO_JPEG, FORMATO_BMP, FORMATO_PNG):
            try:
                imagen = item.Transfer(formato)
                break
            except Exception as e:
                ultimo_error = e
        if imagen is None:
            raise ErrorEscaner("El escáner no entregó la imagen (%s)" % ultimo_error)

        temporal = destino + ".wia"
        if os.path.exists(temporal):
            os.remove(temporal)
        imagen.SaveFile(temporal)

        # Se normaliza a JPEG con Pillow: asi el PDF pesa parecido siempre,
        # sin importar en que formato lo haya entregado el driver.
        with Image.open(temporal) as img:
            img = img.convert("RGB" if color == "color" else "L")
            img.save(destino, "JPEG", quality=90, dpi=(resolucion, resolucion))
        os.remove(temporal)
        return destino

    def _escanear_demo(self, destino, resolucion):
        """Hoja de prueba, para poder revisar la pantalla sin escaner."""
        ancho = int(resolucion * ANCHO_PULGADAS / 2)
        alto = int(resolucion * ALTO_PULGADAS / 2)
        img = Image.new("L", (ancho, alto), 246)
        d = ImageDraw.Draw(img)
        d.rectangle([30, 30, ancho - 30, alto - 30], outline=170, width=3)
        marca = datetime.datetime.now().strftime("%d/%m/%Y  %H:%M:%S")
        d.text((60, 70), "ESCANEO SIMULADO", fill=60)
        d.text((60, 95), marca, fill=90)
        for i in range(14):
            y = 150 + i * 26
            d.line([60, y, ancho - 60 - (i % 4) * 40, y], fill=200, width=2)
        img.save(destino, "JPEG", quality=90, dpi=(resolucion, resolucion))
        return destino


# ===========================================================================
#  ARMADO DEL PDF
# ===========================================================================
def nombre_documento(doc, cantidad_paginas):
    """Nombre del archivo en la carpeta del dia, como se ha archivado siempre:
    'B 137.854.jpeg'. Un documento de una sola pagina queda en JPEG; con varias
    tiene que ser PDF, porque un JPEG no admite mas de una."""
    prefijo = PREFIJO_ADJUNTO.get(doc["tipo"], "")
    base = ("%s %s" % (prefijo, _numero_con_puntos(doc.get("numero")))).strip()
    base = re.sub(r'[\\/:*?"<>|]', "-", base) or ("documento %d" % doc["id"])
    return base + (".jpeg" if cantidad_paginas == 1 else ".pdf")


def archivo_actual(doc):
    return doc.get("archivoSalida") or doc.get("archivoPdf")


def construir_salida(doc, carpeta_destino, cfg):
    """Rehace el archivo del documento y borra el que hubiera antes (el nombre
    cambia si se corrige el numero, y la extension si cambia la cantidad de
    paginas)."""
    paginas = [p["archivo"] for p in doc.get("paginas", []) if os.path.exists(p["archivo"])]
    if not paginas:
        borrar_salida_anterior(doc)
        return None

    os.makedirs(carpeta_destino, exist_ok=True)
    destino = os.path.join(carpeta_destino, nombre_documento(doc, len(paginas)))

    if len(paginas) == 1:
        # La pagina ya esta en JPEG desde el escaneo: se copia tal cual.
        temporal = destino + ".tmp"
        shutil.copy2(paginas[0], temporal)
        os.replace(temporal, destino)
    else:
        calidad = int(cfg.get("calidad_jpeg", 65))
        resolucion = int(cfg.get("resolucion", 200))
        imagenes = []
        try:
            for ruta in paginas:
                img = Image.open(ruta)
                img.load()
                imagenes.append(img.convert("RGB" if cfg.get("color") == "color" else "L"))
            # Se escribe primero un temporal: si algo falla a medio camino, el
            # archivo que ya estaba en la carpeta no se pierde.
            temporal = destino + ".tmp"
            imagenes[0].save(temporal, "PDF", save_all=True, append_images=imagenes[1:],
                             resolution=resolucion, quality=calidad)
            os.replace(temporal, destino)
        finally:
            for img in imagenes:
                try:
                    img.close()
                except Exception:
                    pass

    borrar_salida_anterior(doc, excepto=destino)
    return destino


def borrar_salida_anterior(doc, excepto=None):
    anterior = archivo_actual(doc)
    if not anterior or not os.path.exists(anterior):
        return
    if excepto and os.path.abspath(anterior) == os.path.abspath(excepto):
        return
    try:
        os.remove(anterior)
    except OSError:
        pass


# ===========================================================================
#  MONTOS
#  Se parte del monto AFECTO (el "SUBTOTAL" del voucher o la linea "CONSUMO"
#  de la boleta), NO del total impreso: ese ultimo incluye la propina, que es
#  exenta, y por eso no se puede dividir por 1,19.
#  Las facturas de exportacion van exentas y en dolares: no se calcula IVA.
# ===========================================================================
def desglosar(afecto, exento):
    try:
        monto = float(afecto)
    except (TypeError, ValueError):
        return None, None
    if monto <= 0:
        return None, None
    if exento:
        return int(round(monto)), None
    # El neto se redondea y el IVA se saca por diferencia: asi neto + IVA da
    # siempre exactamente el afecto, sin descuadres de un peso.
    neto = int(round(monto / (1 + IVA_TASA)))
    return neto, int(round(monto)) - neto


def _numero(valor):
    """Acepta '3.600', '3600', ' 3600 ' y devuelve 3600. None si no es numero."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return valor
    texto = str(valor).strip().replace("$", "").replace(" ", "")
    if not texto:
        return None
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def validar(documentos, carpeta_destino):
    """Revisa el listado antes de escribir la planilla o mandar el correo.
    Devuelve dos listas: lo que impide continuar y lo que solo advierte."""
    errores, avisos = [], []

    for d in documentos:
        etiqueta = "%s %s" % (TIPOS.get(d["tipo"], {}).get("etiqueta", d["tipo"]), d.get("numero") or "?")
        exento = TIPOS.get(d["tipo"], {}).get("exento", False)
        neto, iva = _numero(d.get("neto")), _numero(d.get("iva"))
        total = _numero(d.get("total"))

        if not d.get("paginas"):
            errores.append("%s no tiene ninguna página escaneada." % etiqueta)
        if neto is None:
            errores.append("%s no tiene neto." % etiqueta)
        if not (d.get("formaPago") or "").strip():
            errores.append("%s no tiene forma de pago." % etiqueta)

        if neto is not None:
            suma = neto + (iva or 0)
            if total is not None and abs(suma - total) > 1:
                errores.append("%s: neto más IVA da %s y el total dice %s."
                               % (etiqueta, int(suma), int(total)))
        if exento and iva:
            avisos.append("%s es de exportación y tiene IVA cargado." % etiqueta)
        if not exento and neto is not None and not iva:
            avisos.append("%s no tiene IVA. Revisa que corresponda." % etiqueta)

    # Folios saltados dentro de cada serie: la señal mas clara de un documento
    # que se emitio y quedo sin escanear.
    for clave, nombre in (("boleta", "boletas"), ("factura", "facturas"), ("exportacion", "facturas de exportación")):
        folios = sorted(int(d["numero"]) for d in documentos
                        if d["tipo"] == clave and str(d.get("numero", "")).strip().isdigit())
        for anterior, siguiente in zip(folios, folios[1:]):
            faltantes = list(range(anterior + 1, siguiente))
            if 0 < len(faltantes) <= 20:
                avisos.append("Entre las %s %d y %d faltan: %s."
                              % (nombre, anterior, siguiente,
                                 ", ".join(str(x) for x in faltantes)))

    # Archivos sueltos en la carpeta del dia que no estan en el listado.
    if carpeta_destino and os.path.isdir(carpeta_destino):
        esperados = {os.path.basename(archivo_actual(d)).lower()
                     for d in documentos if archivo_actual(d)}
        for archivo in sorted(os.listdir(carpeta_destino)):
            if archivo.lower().endswith(".pdf") and archivo.lower() not in esperados:
                avisos.append("En la carpeta hay un PDF que no está en el listado: %s" % archivo)

    return errores, avisos


# ===========================================================================
#  SERVIDOR
# ===========================================================================
# ===========================================================================
#  ESCRITURA EN LA PLANILLA
#  Se maneja el propio Excel del computador en vez de reescribir el archivo:
#  asi los bordes, colores, fuentes y formulas quedan intactos, porque no se
#  tocan. La hoja del dia se crea DUPLICANDO la del dia anterior y borrando
#  sus datos, que es la unica forma de que el formato sea identico.
# ===========================================================================
class ErrorExcel(Exception):
    pass


def _ultima_fila_del_bloque(hoja):
    """Hasta donde llega el bloque ya formateado (las filas que muestran $0)."""
    ultima = EXCEL_FILA_DATOS
    for fila in range(EXCEL_FILA_DATOS, EXCEL_FILA_DATOS + 400):
        celda_total = hoja.Cells(fila, COL_TOTAL)
        tiene_formula = False
        try:
            tiene_formula = bool(celda_total.HasFormula)
        except Exception:
            pass
        borde = False
        try:                       # 7 = borde izquierdo; 0 = sin linea
            borde = hoja.Cells(fila, COL_TIPO).Borders(7).LineStyle != -4142
        except Exception:
            pass
        if tiene_formula or borde:
            ultima = fila
        elif fila > EXCEL_FILA_DATOS + 2:
            break
    return ultima


def _limpiar_datos(hoja, ultima):
    """Borra los valores dejando el formato y las fórmulas donde las haya."""
    for fila in range(EXCEL_FILA_DATOS, ultima + 1):
        for columna in (COL_TIPO, COL_FOLIO, COL_NETO, COL_IVA,
                        COL_PAGO, COL_AUTORIZACION, COL_OBS):
            hoja.Cells(fila, columna).ClearContents()
        celda_total = hoja.Cells(fila, COL_TOTAL)
        try:
            if not celda_total.HasFormula:
                celda_total.ClearContents()
        except Exception:
            celda_total.ClearContents()


def _extender_filas(hoja, ultima, necesarias):
    """Agrega filas copiando el formato de la última, como se hace a mano."""
    faltan = necesarias - (ultima - EXCEL_FILA_DATOS + 1)
    if faltan <= 0:
        return ultima
    origen = hoja.Range(hoja.Cells(ultima, COL_TIPO), hoja.Cells(ultima, COL_OBS))
    destino = hoja.Range(hoja.Cells(ultima + 1, COL_TIPO), hoja.Cells(ultima + faltan, COL_OBS))
    origen.Copy(destino)
    hoja.Application.CutCopyMode = False
    nueva_ultima = ultima + faltan
    _limpiar_datos(hoja, nueva_ultima)
    return nueva_ultima


def escribir_en_planilla(ruta_excel, fecha, documentos):
    """Crea (o rehace) la hoja del día y devuelve un resumen de lo hecho."""
    if not ruta_excel or not os.path.exists(ruta_excel):
        raise ErrorExcel("No encuentro la planilla de auditoría. Indícala en Configuración.")

    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()

    # Respaldo antes de tocar nada: si algo sale mal, la planilla original
    # sigue intacta en la carpeta 'respaldos'.
    carpeta_respaldos = os.path.join(CARPETA_APP, "respaldos")
    os.makedirs(carpeta_respaldos, exist_ok=True)
    marca = datetime.datetime.now().strftime("%Y-%m-%d %H%M%S")
    respaldo = os.path.join(carpeta_respaldos,
                            "%s - %s" % (marca, os.path.basename(ruta_excel)))
    shutil.copy2(ruta_excel, respaldo)

    nombre_hoja = fecha.strftime("%d.%m.%Y")
    excel = None
    libro = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        libro = excel.Workbooks.Open(os.path.abspath(ruta_excel))
        if libro.ReadOnly:
            raise ErrorExcel("La planilla está abierta en modo solo lectura. "
                             "Ciérrala en Excel y vuelve a intentar.")

        hoja = None
        for h in libro.Worksheets:
            if str(h.Name).strip() == nombre_hoja:
                hoja = h
                break
        creada = hoja is None
        if creada:
            plantilla = libro.Worksheets(libro.Worksheets.Count)
            plantilla.Copy(None, libro.Worksheets(libro.Worksheets.Count))
            hoja = libro.Worksheets(libro.Worksheets.Count)
            hoja.Name = nombre_hoja

        ultima = _ultima_fila_del_bloque(hoja)
        _limpiar_datos(hoja, ultima)
        ultima = _extender_filas(hoja, ultima, len(documentos))

        fila = EXCEL_FILA_DATOS
        for d in documentos:
            exento = TIPOS.get(d["tipo"], {}).get("exento", False)
            hoja.Cells(fila, COL_TIPO).Value = TIPO_EN_PLANILLA.get(d["tipo"], d["tipo"].upper())
            folio = d.get("numero") or ""
            hoja.Cells(fila, COL_FOLIO).Value = int(folio) if str(folio).strip().isdigit() else folio

            neto, iva = _numero(d.get("neto")), _numero(d.get("iva"))
            hoja.Cells(fila, COL_NETO).Value = neto if neto is not None else ""
            hoja.Cells(fila, COL_IVA).Value = "" if (exento or not iva) else iva
            celda_total = hoja.Cells(fila, COL_TOTAL)
            try:
                if not celda_total.HasFormula:     # si es fórmula, se respeta
                    celda_total.Value = (neto or 0) + (iva or 0)
            except Exception:
                celda_total.Value = (neto or 0) + (iva or 0)

            hoja.Cells(fila, COL_PAGO).Value = (d.get("formaPago") or "").strip()
            hoja.Cells(fila, COL_AUTORIZACION).Value = (d.get("autorizacion") or "").strip()
            hoja.Cells(fila, COL_OBS).Value = (d.get("observacion") or "").strip()
            fila += 1

        libro.Save()
        return {"hoja": nombre_hoja, "creada": creada, "filas": len(documentos),
                "respaldo": respaldo}
    except ErrorExcel:
        raise
    except Exception as e:
        raise ErrorExcel("Excel devolvió un error: %s\n\n"
                         "La planilla original quedó respaldada en:\n%s" % (e, respaldo))
    finally:
        try:
            if libro is not None:
                libro.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ===========================================================================
#  CORREO A CONTABILIDAD
#  Se arma en Outlook y se deja ABIERTO para revisar: nunca se envia solo.
#  Los documentos de una sola pagina van como JPEG (que es como los recibe
#  contabilidad hoy) y los de varias como PDF, que es lo unico que soporta
#  mas de una pagina.
# ===========================================================================
class ErrorCorreo(Exception):
    pass


def _numero_con_puntos(numero):
    texto = str(numero or "").strip()
    return "{:,}".format(int(texto)).replace(",", ".") if texto.isdigit() else texto


def preparar_adjuntos(documentos, carpeta_trabajo, ruta_excel):
    """Los adjuntos son los mismos archivos de la carpeta del dia: ya estan con
    el nombre y el formato correctos, no hace falta copiarlos a ningun lado."""
    rutas = []
    if ruta_excel and os.path.exists(ruta_excel):
        rutas.append(ruta_excel)
    for doc in documentos:
        archivo = archivo_actual(doc)
        if archivo and os.path.exists(archivo):
            rutas.append(archivo)
    return rutas


def cuerpo_correo(cfg, fecha, remitente):
    dia = fecha.strftime("%d-%m-%Y")
    saludo = (cfg.get("correo_saludo") or "Estimada Josefa").strip().rstrip(",")
    return (
        '<div style="font-family:Aptos,Calibri,\'Segoe UI\',sans-serif;font-size:11pt;color:#000000;">'
        '<p style="margin:0 0 11pt 0;">' + saludo + ',</p>'
        '<p style="margin:0 0 11pt 0;">Junto con saludar, adjunto planilla de auditoría '
        'y documentos de pagos del ' + dia + '.</p>'
        '<p style="margin:0 0 11pt 0;">Quedamos atentos</p>'
        '<p style="margin:0 0 11pt 0;">Se despide cordialmente</p>'
        '<p style="margin:0 0 14pt 0;">' + (remitente or "").strip() + '.</p>'
        '</div>'
    )


def abrir_correo(cfg, fecha, documentos, remitente, carpeta_trabajo):
    """Abre el correo en Outlook con todo adjunto, listo para revisar y enviar."""
    if not (remitente or "").strip():
        raise ErrorCorreo("Elige quién envía el correo.")

    adjuntos = preparar_adjuntos(documentos, carpeta_trabajo, cfg.get("archivo_excel"))
    if not adjuntos:
        raise ErrorCorreo("No hay nada que adjuntar.")

    asunto = "Cascadas Hotel/ Recibimientos y Pagos %s." % fecha.strftime("%d-%m-%Y")
    cuerpo = cuerpo_correo(cfg, fecha, remitente)

    import pythoncom
    pythoncom.CoInitialize()
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mensaje = outlook.CreateItem(0)          # 0 = correo nuevo
        mensaje.To = cfg.get("correo_para", "")
        mensaje.CC = cfg.get("correo_cc", "")
        mensaje.Subject = asunto
        for ruta in adjuntos:
            mensaje.Attachments.Add(os.path.abspath(ruta))
        # Se muestra primero para que Outlook ponga la firma del computador, y
        # el cuerpo se antepone a lo que quedo: asi la firma es la de verdad,
        # no una imitacion.
        mensaje.Display()
        try:
            firma = mensaje.HTMLBody or ""
        except Exception:
            firma = ""
        mensaje.HTMLBody = cuerpo + firma
        return {"via": "outlook", "adjuntos": len(adjuntos), "asunto": asunto}
    except Exception as e:
        raise ErrorCorreo("No se pudo abrir Outlook: %s" % e)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ===========================================================================
#  BUSCADOR DE DOCUMENTOS
#  Recorre toda la carpeta de escaneados, no solo lo que paso por esta
#  herramienta: los documentos de antes tambien aparecen.
# ===========================================================================
def buscar_documentos(carpeta_base, consulta, tope=60):
    consulta = re.sub(r"\D", "", str(consulta or ""))
    if len(consulta) < 3:
        raise ValueError("Escribe al menos 3 números del documento.")
    if not carpeta_base or not os.path.isdir(carpeta_base):
        raise ValueError("No encuentro la carpeta de escaneados.")

    resultados = []
    for base, carpetas, archivos in os.walk(carpeta_base):
        carpetas.sort(reverse=True)          # los años y días recientes primero
        for archivo in sorted(archivos):
            if archivo.startswith("~$"):
                continue
            solo_numeros = re.sub(r"\D", "", os.path.splitext(archivo)[0])
            if consulta not in solo_numeros:
                continue
            ruta = os.path.join(base, archivo)
            relativa = os.path.relpath(base, carpeta_base).replace("\\", "/")
            try:
                kb = round(os.path.getsize(ruta) / 1024)
            except OSError:
                kb = 0
            resultados.append({
                "nombre": archivo, "ruta": ruta, "carpeta": relativa, "kb": kb,
                "ubicacion": relativa.replace("/", " · "),
            })
            if len(resultados) >= tope:
                return resultados
    return resultados


def abrir_archivo(ruta, carpeta_base):
    """Abre un archivo del buscador. Solo se permite dentro de la carpeta de
    escaneados: la pantalla corre en un navegador y no conviene que pueda
    pedir cualquier archivo del computador."""
    ruta = os.path.abspath(ruta)
    base = os.path.abspath(carpeta_base or "")
    if not base or not (ruta == base or ruta.startswith(base + os.sep)):
        raise ValueError("Ese archivo está fuera de la carpeta de escaneados.")
    if not os.path.exists(ruta):
        raise ValueError("El archivo ya no está en el disco.")
    if ES_WINDOWS:
        os.startfile(ruta)                                  # noqa: pylint
    elif sys.platform == "darwin":
        subprocess.Popen(["open", ruta])
    else:
        subprocess.Popen(["xdg-open", ruta])
    return ruta


# ===========================================================================
#  LIMPIEZA
#  'trabajo' guarda las paginas escaneadas del dia y el avance; 'respaldos'
#  las copias de la planilla. Ambas hacen falta, pero no para siempre.
# ===========================================================================
def limpiar_antiguos(dias_trabajo=60, max_respaldos=20):
    try:
        limite = datetime.date.today() - datetime.timedelta(days=dias_trabajo)
        for nombre in os.listdir(CARPETA_TRABAJO):
            ruta = os.path.join(CARPETA_TRABAJO, nombre)
            if not os.path.isdir(ruta):
                continue
            try:
                fecha = datetime.date.fromisoformat(nombre)
            except ValueError:
                continue
            if fecha < limite:
                shutil.rmtree(ruta, ignore_errors=True)
    except FileNotFoundError:
        pass

    carpeta_respaldos = os.path.join(CARPETA_APP, "respaldos")
    try:
        archivos = sorted(
            (os.path.join(carpeta_respaldos, n) for n in os.listdir(carpeta_respaldos)),
            key=os.path.getmtime, reverse=True)
        for viejo in archivos[max_respaldos:]:
            try:
                os.remove(viejo)
            except OSError:
                pass
    except FileNotFoundError:
        pass


class Aplicacion(object):
    def __init__(self):
        self.cfg = leer_config()
        self.fecha = datetime.date.today()
        self.estado = EstadoDia(self.fecha)
        self.escaner = Escaner()
        # Primera vez: se intenta reconocer sola la carpeta de escaneados.
        if not (self.cfg.get("carpeta_base") or "").strip():
            encontrada = buscar_carpeta_escaneados()
            if encontrada:
                self.cfg["carpeta_base"] = encontrada
                guardar_config(self.cfg)
                print("  Carpeta de escaneados reconocida automaticamente.")
        if not (self.cfg.get("archivo_excel") or "").strip():
            planilla = buscar_planilla_auditoria([self.cfg.get("carpeta_base") or ""])
            if planilla:
                self.cfg["archivo_excel"] = planilla
                guardar_config(self.cfg)
                print("  Planilla de auditoria reconocida automaticamente.")

    def carpeta_destino(self, crear=True):
        base = (self.cfg.get("carpeta_base") or "").strip()
        if not base:
            return ""
        return carpeta_del_dia(base, self.fecha, crear=crear)

    # ---- consultas ----
    def guardar_valores(self, doc_id, campos):
        """Guarda los montos y datos de pago de un documento."""
        doc = self.estado.buscar(doc_id)
        if not doc:
            raise ValueError("El documento ya no está en el listado.")
        with self.estado.lock:
            for campo in ("afecto", "neto", "iva", "total", "formaPago", "autorizacion", "observacion"):
                if campo in campos:
                    doc[campo] = campos[campo]
            forma = (doc.get("formaPago") or "").strip().upper()
            if forma and forma not in self.cfg.get("formas_pago", []):
                self.cfg.setdefault("formas_pago", []).append(forma)
                guardar_config(self.cfg)
            self.estado.guardar()
        return doc

    def validaciones(self):
        errores, avisos = validar(self.estado.documentos, self.carpeta_destino(crear=False))
        return {"errores": errores, "avisos": avisos}

    def escribir_excel(self):
        errores, _avisos = validar(self.estado.documentos, self.carpeta_destino(crear=False))
        if errores:
            raise ErrorExcel("Faltan datos por completar:\n\n· " + "\n· ".join(errores))
        if not self.estado.documentos:
            raise ErrorExcel("No hay documentos que escribir.")
        return escribir_en_planilla(self.cfg.get("archivo_excel"), self.fecha, self.estado.documentos)

    def abrir_correo(self, remitente):
        errores, _avisos = validar(self.estado.documentos, self.carpeta_destino(crear=False))
        if errores:
            raise ErrorCorreo("Faltan datos por completar:\n\n· " + "\n· ".join(errores))
        nombre = (remitente or "").strip()
        if nombre and nombre not in self.cfg.get("remitentes", []):
            self.cfg.setdefault("remitentes", []).append(nombre)
            guardar_config(self.cfg)
        return abrir_correo(self.cfg, self.fecha, self.estado.documentos,
                            nombre, self.estado.carpeta_trabajo)

    def agregar_remitente(self, nombre):
        nombre = (nombre or "").strip()
        if not nombre:
            raise ValueError("Escribe un nombre.")
        if nombre not in self.cfg.get("remitentes", []):
            self.cfg.setdefault("remitentes", []).append(nombre)
            self.cfg["remitentes"].sort()
            guardar_config(self.cfg)
        return self.cfg["remitentes"]

    def eliminar_remitente(self, nombre):
        self.cfg["remitentes"] = [n for n in self.cfg.get("remitentes", []) if n != nombre]
        guardar_config(self.cfg)
        return self.cfg["remitentes"]

    def resumen(self):
        base = (self.cfg.get("carpeta_base") or "").strip()
        excel = (self.cfg.get("archivo_excel") or "").strip()
        errores, avisos = validar(self.estado.documentos, self.carpeta_destino(crear=False))
        return {
            "fecha": self.fecha.isoformat(),
            "fechaTexto": "%d de %s de %d" % (self.fecha.day, MESES[self.fecha.month - 1].lower(), self.fecha.year),
            "config": self.cfg,
            "carpetaDestino": self.carpeta_destino(crear=False),
            "baseExiste": bool(base) and os.path.isdir(base),
            "excelExiste": bool(excel) and os.path.exists(excel),
            "hojaDelDia": self.fecha.strftime("%d.%m.%Y"),
            "documentos": self.estado.documentos,
            "tipos": [{"clave": k, "etiqueta": v["etiqueta"], "exento": v["exento"]} for k, v in TIPOS.items()],
            "formasPago": self.cfg.get("formas_pago", []),
            "remitentes": self.cfg.get("remitentes", []),
            "correoPara": self.cfg.get("correo_para", ""),
            "correoCc": self.cfg.get("correo_cc", ""),
            "errores": errores,
            "avisos": avisos,
            "demo": MODO_DEMO or not ES_WINDOWS,
        }

    # ---- acciones ----
    def crear_documento(self, tipo, numero):
        if tipo not in TIPOS:
            raise ValueError("Tipo de documento no válido.")
        numero = (numero or "").strip()
        if not numero:
            raise ValueError("Falta el número del documento.")
        repetido = [d for d in self.estado.documentos
                    if d["numero"].lower() == numero.lower() and d["tipo"] == tipo]
        if repetido:
            raise ValueError("Ya hay una %s con el número %s en el listado de hoy."
                             % (TIPOS[tipo]["etiqueta"].lower(), numero))
        with self.estado.lock:
            doc = {
                "id": self.estado.siguiente_id,
                "tipo": tipo,
                "numero": numero,
                "paginas": [],
                "archivoSalida": None,
                "creado": datetime.datetime.now().strftime("%H:%M"),
            }
            self.estado.siguiente_id += 1
            self.estado.documentos.append(doc)
            self.estado.guardar()
        return doc

    def escanear_pagina(self, doc_id, etiqueta):
        doc = self.estado.buscar(doc_id)
        if not doc:
            raise ValueError("El documento ya no está en el listado.")
        indice = len(doc["paginas"]) + 1
        carpeta_doc = os.path.join(self.estado.carpeta_trabajo, "doc-%d" % doc_id)
        os.makedirs(carpeta_doc, exist_ok=True)
        destino_img = os.path.join(carpeta_doc, "pagina-%02d.jpg" % indice)

        self.escaner.escanear(destino_img, self.cfg.get("resolucion", 200), self.cfg.get("color", "gris"))

        with self.estado.lock:
            doc["paginas"].append({
                "etiqueta": (etiqueta or "Documento").strip() or "Documento",
                "archivo": destino_img,
            })
            doc["archivoSalida"] = construir_salida(doc, self.carpeta_destino(), self.cfg)
            self.estado.guardar()
        return doc

    def eliminar_pagina(self, doc_id, indice):
        doc = self.estado.buscar(doc_id)
        if not doc:
            raise ValueError("El documento ya no está en el listado.")
        if indice < 0 or indice >= len(doc["paginas"]):
            raise ValueError("Esa página no existe.")
        with self.estado.lock:
            pagina = doc["paginas"].pop(indice)
            try:
                if os.path.exists(pagina["archivo"]):
                    os.remove(pagina["archivo"])
            except OSError:
                pass
            doc["archivoSalida"] = construir_salida(doc, self.carpeta_destino(), self.cfg)
            self.estado.guardar()
        return doc

    def editar_documento(self, doc_id, tipo, numero):
        doc = self.estado.buscar(doc_id)
        if not doc:
            raise ValueError("El documento ya no está en el listado.")
        numero = (numero or "").strip()
        if not numero:
            raise ValueError("Falta el número del documento.")
        if tipo not in TIPOS:
            raise ValueError("Tipo de documento no válido.")
        choque = [d for d in self.estado.documentos
                  if d["id"] != doc_id and d["numero"].lower() == numero.lower() and d["tipo"] == tipo]
        if choque:
            raise ValueError("Ya hay una %s con el número %s." % (TIPOS[tipo]["etiqueta"].lower(), numero))
        with self.estado.lock:
            doc["tipo"] = tipo
            doc["numero"] = numero
            doc["archivoSalida"] = construir_salida(doc, self.carpeta_destino(), self.cfg)
            self.estado.guardar()
        return doc

    def eliminar_documento(self, doc_id):
        doc = self.estado.buscar(doc_id)
        if not doc:
            return
        with self.estado.lock:
            borrar_salida_anterior(doc)
            carpeta_doc = os.path.join(self.estado.carpeta_trabajo, "doc-%d" % doc_id)
            shutil.rmtree(carpeta_doc, ignore_errors=True)
            self.estado.documentos = [d for d in self.estado.documentos if d["id"] != doc_id]
            self.estado.guardar()

    def guardar_config(self, nueva):
        base = (nueva.get("carpeta_base") or "").strip()
        if not base:
            raise ValueError("Falta la carpeta de escaneados.")
        self.cfg["carpeta_base"] = base
        if "archivo_excel" in nueva:
            self.cfg["archivo_excel"] = (nueva.get("archivo_excel") or "").strip()
        self.cfg["resolucion"] = max(100, min(600, int(nueva.get("resolucion", 200))))
        self.cfg["color"] = "color" if nueva.get("color") == "color" else "gris"
        self.cfg["calidad_jpeg"] = max(40, min(95, int(nueva.get("calidad_jpeg", 65))))
        guardar_config(self.cfg)
        return self.cfg

    def abrir_carpeta(self):
        carpeta = self.carpeta_destino()
        if ES_WINDOWS:
            os.startfile(carpeta)                      # noqa: pylint
        elif sys.platform == "darwin":
            subprocess.Popen(["open", carpeta])
        else:
            subprocess.Popen(["xdg-open", carpeta])
        return carpeta


APP = None


class Manejador(http.server.BaseHTTPRequestHandler):
    server_version = "AuditoriaCascadas"

    def log_message(self, formato, *args):
        pass   # sin ruido en la consola

    # ---- utilidades ----
    def _responder(self, codigo, cuerpo, tipo="application/json; charset=utf-8"):
        datos = cuerpo if isinstance(cuerpo, bytes) else cuerpo.encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(datos)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(datos)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, obj, codigo=200):
        self._responder(codigo, json.dumps(obj, ensure_ascii=False))

    def _error(self, mensaje, codigo=400):
        self._json({"ok": False, "error": mensaje}, codigo)

    def _cuerpo(self):
        largo = int(self.headers.get("Content-Length") or 0)
        if not largo:
            return {}
        return json.loads(self.rfile.read(largo).decode("utf-8"))

    # ---- rutas ----
    def do_GET(self):
        ruta = self.path.split("?")[0]
        if ruta in ("/", "/index.html"):
            archivo = os.path.join(CARPETA_APP, "interfaz.html")
            if not os.path.exists(archivo):
                return self._responder(500, "Falta el archivo interfaz.html en la carpeta de la herramienta.",
                                       "text/plain; charset=utf-8")
            with open(archivo, "rb") as f:
                return self._responder(200, f.read(), "text/html; charset=utf-8")
        if ruta == "/api/estado":
            return self._json({"ok": True, "datos": APP.resumen()})
        if ruta == "/api/escaneres":
            try:
                return self._json({"ok": True, "equipos": APP.escaner.listar()})
            except Exception as e:
                return self._error(str(e))
        if ruta == "/api/vista":
            return self._vista_pagina()
        return self._error("Ruta no encontrada.", 404)

    def _vista_pagina(self):
        """Miniatura de una pagina escaneada, para verla en pantalla."""
        from urllib.parse import parse_qs, urlparse
        params = parse_qs(urlparse(self.path).query)
        try:
            doc_id = int(params.get("doc", ["0"])[0])
            indice = int(params.get("pag", ["0"])[0])
        except ValueError:
            return self._error("Parámetros inválidos.")
        doc = APP.estado.buscar(doc_id)
        if not doc or indice >= len(doc["paginas"]):
            return self._error("No existe esa página.", 404)
        ruta = doc["paginas"][indice]["archivo"]
        if not os.path.exists(ruta):
            return self._error("El archivo ya no está.", 404)
        with Image.open(ruta) as img:
            img.thumbnail((700, 700))
            buf = io.BytesIO()
            img.convert("RGB").save(buf, "JPEG", quality=70)
        return self._responder(200, buf.getvalue(), "image/jpeg")

    def do_POST(self):
        ruta = self.path.split("?")[0]
        try:
            cuerpo = self._cuerpo()
        except Exception:
            return self._error("No se entendió la petición.")
        try:
            if ruta == "/api/config":
                return self._json({"ok": True, "config": APP.guardar_config(cuerpo)})
            if ruta == "/api/documento":
                doc = APP.crear_documento(cuerpo.get("tipo"), cuerpo.get("numero"))
                return self._json({"ok": True, "documento": doc})
            if ruta == "/api/documento/editar":
                doc = APP.editar_documento(int(cuerpo.get("id")), cuerpo.get("tipo"), cuerpo.get("numero"))
                return self._json({"ok": True, "documento": doc})
            if ruta == "/api/documento/eliminar":
                APP.eliminar_documento(int(cuerpo.get("id")))
                return self._json({"ok": True})
            if ruta == "/api/escanear":
                doc = APP.escanear_pagina(int(cuerpo.get("id")), cuerpo.get("etiqueta"))
                return self._json({"ok": True, "documento": doc})
            if ruta == "/api/pagina/eliminar":
                doc = APP.eliminar_pagina(int(cuerpo.get("id")), int(cuerpo.get("indice")))
                return self._json({"ok": True, "documento": doc})
            if ruta == "/api/abrir-carpeta":
                return self._json({"ok": True, "carpeta": APP.abrir_carpeta()})
            if ruta == "/api/elegir-carpeta":
                elegida = elegir_carpeta_con_ventana(APP.cfg.get("carpeta_base") or "")
                if not elegida:
                    return self._json({"ok": True, "carpeta": ""})
                return self._json({"ok": True, "carpeta": os.path.normpath(elegida)})
            if ruta == "/api/buscar-documentos":
                encontrados = buscar_documentos(APP.cfg.get("carpeta_base"), cuerpo.get("consulta"))
                return self._json({"ok": True, "resultados": encontrados})
            if ruta == "/api/abrir-archivo":
                return self._json({"ok": True,
                                   "ruta": abrir_archivo(cuerpo.get("ruta"), APP.cfg.get("carpeta_base"))})
            if ruta == "/api/buscar-carpeta":
                return self._json({"ok": True, "carpeta": buscar_carpeta_escaneados()})
            if ruta == "/api/elegir-excel":
                elegido = elegir_archivo_con_ventana(
                    os.path.dirname(APP.cfg.get("archivo_excel") or "") or APP.cfg.get("carpeta_base") or "")
                return self._json({"ok": True, "archivo": os.path.normpath(elegido) if elegido else ""})
            if ruta == "/api/documento/valores":
                doc = APP.guardar_valores(int(cuerpo.get("id")), cuerpo.get("campos") or {})
                return self._json({"ok": True, "documento": doc})
            if ruta == "/api/excel/escribir":
                return self._json({"ok": True, "resultado": APP.escribir_excel()})
            if ruta == "/api/correo":
                return self._json({"ok": True, "resultado": APP.abrir_correo(cuerpo.get("remitente"))})
            if ruta == "/api/remitente/agregar":
                return self._json({"ok": True, "remitentes": APP.agregar_remitente(cuerpo.get("nombre"))})
            if ruta == "/api/remitente/eliminar":
                return self._json({"ok": True, "remitentes": APP.eliminar_remitente(cuerpo.get("nombre"))})
        except (ValueError, ErrorEscaner, ErrorExcel, ErrorCorreo) as e:
            return self._error(str(e))
        except Exception as e:
            traceback.print_exc()
            return self._error("Error inesperado: %s" % e, 500)
        return self._error("Ruta no encontrada.", 404)


class Servidor(http.server.ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def puerto_libre(desde=8760, intentos=20):
    for p in range(desde, desde + intentos):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    raise RuntimeError("No hay puertos libres para levantar la herramienta.")


def main():
    global APP
    os.makedirs(CARPETA_TRABAJO, exist_ok=True)
    limpiar_antiguos()
    APP = Aplicacion()

    puerto = puerto_libre()
    direccion = "http://127.0.0.1:%d/" % puerto
    servidor = Servidor(("127.0.0.1", puerto), Manejador)

    print("=" * 66)
    print("  CASCADAS HOTEL - AUDITORIA DIARIA")
    print("=" * 66)
    print("  Fecha:    %s" % APP.fecha.strftime("%d/%m/%Y"))
    print("  Carpeta:  %s" % APP.carpeta_destino(crear=False))
    print("  Pantalla: %s" % direccion)
    if MODO_DEMO or not ES_WINDOWS:
        print("  MODO DEMO: el escaner esta simulado.")
    print("=" * 66)
    print("\n  Deja esta ventana abierta mientras trabajas.")
    print("  Para cerrar la herramienta, cierra esta ventana.\n")

    threading.Timer(1.0, lambda: webbrowser.open(direccion)).start()
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nHerramienta cerrada.")


if __name__ == "__main__":
    main()
