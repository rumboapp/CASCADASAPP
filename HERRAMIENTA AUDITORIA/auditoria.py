# -*- coding: utf-8 -*-
"""
============================================================================
CASCADAS HOTEL - AUDITORIA DIARIA
============================================================================
Herramienta local para el cierre diario de boletas y facturas.

Corre en el computador de recepcion: un servidor pequeno en Python que hace
el trabajo con el escaner y los archivos, y una pantalla en el navegador.

Esta primera entrega cubre:
  - escanear desde el Canon (WIA), eligiendo tipo y numero de documento
  - agregar paginas al mismo documento (voucher, transferencia, pasaporte)
  - armar el PDF y dejarlo en la carpeta del dia, con el numero como nombre

El listado con valores, el Excel y el correo vienen en las siguientes
entregas. El avance del dia se guarda solo: si se apaga el computador, al
volver a abrir esta todo donde estaba.

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
    "resolucion": 200,
    "color": "gris",          # 'gris' o 'color'
    "calidad_jpeg": 65,
}


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
def nombre_archivo(doc):
    """Nombre del PDF: el numero del documento, como se hace hoy a mano."""
    numero = re.sub(r'[\\/:*?"<>|]', "-", (doc.get("numero") or "").strip())
    return (numero or ("sin numero %d" % doc["id"])) + ".pdf"


def construir_pdf(doc, carpeta_destino, cfg):
    """Rehace el PDF del documento con todas sus paginas, en orden."""
    paginas = [p["archivo"] for p in doc.get("paginas", []) if os.path.exists(p["archivo"])]
    if not paginas:
        return None

    calidad = int(cfg.get("calidad_jpeg", 65))
    resolucion = int(cfg.get("resolucion", 200))
    imagenes = []
    try:
        for ruta in paginas:
            img = Image.open(ruta)
            img.load()
            imagenes.append(img.convert("RGB" if cfg.get("color") == "color" else "L"))

        os.makedirs(carpeta_destino, exist_ok=True)
        destino = os.path.join(carpeta_destino, nombre_archivo(doc))

        # Se escribe primero un archivo temporal: si algo falla a medio camino,
        # el PDF que ya estaba en la carpeta no se pierde.
        temporal = destino + ".tmp"
        imagenes[0].save(temporal, "PDF", save_all=True, append_images=imagenes[1:],
                         resolution=resolucion, quality=calidad)
        os.replace(temporal, destino)
        return destino
    finally:
        for img in imagenes:
            try:
                img.close()
            except Exception:
                pass


def borrar_pdf_anterior(doc, carpeta_destino):
    """Si cambio el numero, se saca el PDF con el nombre viejo."""
    anterior = doc.get("archivoPdf")
    if anterior and os.path.exists(anterior):
        nuevo = os.path.join(carpeta_destino, nombre_archivo(doc))
        if os.path.abspath(anterior) != os.path.abspath(nuevo):
            try:
                os.remove(anterior)
            except OSError:
                pass


# ===========================================================================
#  SERVIDOR
# ===========================================================================
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

    def carpeta_destino(self, crear=True):
        base = (self.cfg.get("carpeta_base") or "").strip()
        if not base:
            return ""
        return carpeta_del_dia(base, self.fecha, crear=crear)

    # ---- consultas ----
    def resumen(self):
        base = (self.cfg.get("carpeta_base") or "").strip()
        return {
            "fecha": self.fecha.isoformat(),
            "fechaTexto": "%d de %s de %d" % (self.fecha.day, MESES[self.fecha.month - 1].lower(), self.fecha.year),
            "config": self.cfg,
            "carpetaDestino": self.carpeta_destino(crear=False),
            "baseExiste": bool(base) and os.path.isdir(base),
            "documentos": self.estado.documentos,
            "tipos": [{"clave": k, "etiqueta": v["etiqueta"], "exento": v["exento"]} for k, v in TIPOS.items()],
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
                "archivoPdf": None,
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
            doc["archivoPdf"] = construir_pdf(doc, self.carpeta_destino(), self.cfg)
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
            if doc["paginas"]:
                doc["archivoPdf"] = construir_pdf(doc, self.carpeta_destino(), self.cfg)
            else:
                borrar_pdf_anterior(doc, self.carpeta_destino(crear=False))
                doc["archivoPdf"] = None
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
            carpeta = self.carpeta_destino()
            anterior_nombre = nombre_archivo(doc)
            doc["tipo"] = tipo
            doc["numero"] = numero
            if nombre_archivo(doc) != anterior_nombre:
                borrar_pdf_anterior(doc, carpeta)
            if doc["paginas"]:
                doc["archivoPdf"] = construir_pdf(doc, carpeta, self.cfg)
            self.estado.guardar()
        return doc

    def eliminar_documento(self, doc_id):
        doc = self.estado.buscar(doc_id)
        if not doc:
            return
        with self.estado.lock:
            borrar_pdf_anterior(doc, self.carpeta_destino(crear=False))
            pdf = doc.get("archivoPdf")
            if pdf and os.path.exists(pdf):
                try:
                    os.remove(pdf)
                except OSError:
                    pass
            carpeta_doc = os.path.join(self.estado.carpeta_trabajo, "doc-%d" % doc_id)
            shutil.rmtree(carpeta_doc, ignore_errors=True)
            self.estado.documentos = [d for d in self.estado.documentos if d["id"] != doc_id]
            self.estado.guardar()

    def guardar_config(self, nueva):
        base = (nueva.get("carpeta_base") or "").strip()
        if not base:
            raise ValueError("Falta la carpeta de escaneados.")
        self.cfg["carpeta_base"] = base
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
            if ruta == "/api/buscar-carpeta":
                return self._json({"ok": True, "carpeta": buscar_carpeta_escaneados()})
        except (ValueError, ErrorEscaner) as e:
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
