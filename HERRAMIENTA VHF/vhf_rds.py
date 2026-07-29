# -*- coding: utf-8 -*-
"""
=============================================================================
  CASCADAS HOTEL - ABRIR EL RESUMEN DIARIO DE SITUACION (RDS) EN VISUAL HOTEL
=============================================================================
  Prueba de automatizacion: el programa maneja las ventanas de VHF como si
  fuera una persona. Abre el sistema, pasa la pantalla de parametros, entra
  con el usuario de recepcion y deja abierto el informe RDS.

  QUE HACE, PASO A PASO
    1. Abre C:\\CM\\Bin\\VHF.exe
    2. Aprieta Ok en "Parametros de Registro del Sistema FrontOffice"
    3. Entra con el usuario y la clave de recepcion
    4. Menu Consultas -> Informes
    5. En el arbol: FrontOffice -> Administrales -> Resumen Diario de
       Situacion - RDS, y aprieta Visualizar
    6. En la ventana del periodo aprieta Ok. Por defecto deja las fechas que
       propone el sistema (del 1 del mes hasta ayer); con --ayer pide solo el
       dia de ayer y con --hoy el de hoy.

  COMO SE USA
    Doble clic en "Informe RDS.bat", o desde la consola:
        py vhf_rds.py
        py vhf_rds.py --espia    (ver que ventanas y controles hay)
        py vhf_rds.py --lento    (esperar mas entre paso y paso)
        py vhf_rds.py --ayer     (el informe del dia de ayer)

  ANTES DE LA PRIMERA VEZ
    Visual Hotel es un programa de 32 bits, asi que conviene manejarlo con un
    Python de 32 bits. Si tienes el de 64 igual funciona casi todo, pero el
    listado de informes puede leerse mal. Para dejarlo redondo:
        1. Instala Python 32-bit desde python.org (Windows installer 32-bit)
        2. py -3-32 -m pip install pywinauto
    Si te quedas con el de 64 bits, basta con:
        py -m pip install pywinauto

  IMPORTANTE
    - Visual Hotel exige permisos de administrador, asi que esto tambien los
      pide: Windows va a mostrar el cartel azul y hay que darle "Si". No es
      capricho del script: Windows no deja que un programa normal maneje las
      ventanas de uno elevado.
    - El computador tiene que estar desbloqueado y con la pantalla encendida:
      esto mueve ventanas de verdad, no trabaja escondido.
    - Mientras corre, conviene no tocar el teclado ni el mouse.
    - Si algo cambia de nombre en el sistema, el script avisa exactamente en
      que paso se quedo y con --espia se ve como se llaman las cosas ahora.
=============================================================================
"""
import ctypes
import datetime
import os
import re
import sys
import time
import unicodedata

# ---------------------------------------------------------------------------
#  LO QUE SE PUEDE CAMBIAR SIN TOCAR EL RESTO
# ---------------------------------------------------------------------------
RUTA_VHF = r"C:\CM\Bin\VHF.exe"
USUARIO = "RECEPCION1"
CLAVE = "1234x"

# Camino dentro del arbol de informes. Los nombres son los que se ven en
# pantalla; las tildes y mayusculas dan lo mismo, se comparan sin ellas.
CAMINO_INFORME = ["FrontOffice", "Administrales", "Resumen Diario de Situación - RDS"]

# Titulos de las ventanas. Se busca un pedazo del titulo, no el titulo
# completo, porque el final cambia con la version ("v4.38.24m"). Las tildes
# tampoco importan: el sistema mezcla espanol y portugues ("Parametros",
# "Parámetros", "Parâmetros") y se aceptan todas.
TITULO_PARAMETROS = "Parametros de Registro"
TITULO_LOGIN = "login"
# OJO: el titulo real dice "Visual Hotal FrontOffice" (con A: es una falta de
# ortografia del propio sistema), asi que se busca solo el pedazo seguro.
TITULO_PRINCIPAL = "Visual Hot"
TITULO_INFORMES = "Visualizar Informes"
TITULO_RDS = "Resumen Diario de Situacion"   # la ventana que pide el periodo

# Mejor que el titulo: el nombre interno de la ventana, que no cambia aunque
# le arreglen la ortografia o le pongan otra version.
CLASE_LOGIN = "TfrmLogin"
CLASE_PRINCIPAL = "TfrmPrincipal"

MENU_INFORMES = ["Consultas", "Informes"]

# Camino por teclado, para cuando no se puede leer el menu (pasa con el Python
# de 64 bits). La barra de menus va en este orden:
#   0 Sistema   1 Editar   2 Reserva   3 Recepcion   4 Caja   5 Ama de Llaves
#   6 Ventas    7 Registros   8 CONSULTAS   9 CMNet   10 Ventana   11 Ayuda
# y dentro de Consultas, "Informes" es el primero de la lista.
MENU_POSICION = 8
MENU_ITEM_POSICION = 0

# Lo mismo para el listado de informes: bajo FrontOffice, "Administrales" es la
# septima carpeta (Cadastrais, Caja, Clientes, Contables, Emisiones Diversas,
# Estadisticos, Administrales).
CARPETA_POSICION = 7
# Numero del informe, que el sistema muestra al costado al elegirlo. Sirve para
# confirmar que se selecciono el correcto y no el "Modelo II".
NUMERO_INFORME = "2100.00-1"

# Que periodo pedirle al informe:
#   "dejar" -> los que el sistema propone solo (del 1 del mes hasta ayer)
#   "ayer"  -> el dia de ayer, desde y hasta
#   "hoy"   -> el dia de hoy
# Tambien se puede elegir al vuelo:  py vhf_rds.py --ayer
FECHAS_RDS = "dejar"

LENTO = "--lento" in sys.argv
ESPIA = "--espia" in sys.argv
if "--ayer" in sys.argv:
    FECHAS_RDS = "ayer"
if "--hoy" in sys.argv:
    FECHAS_RDS = "hoy"
PAUSA = 1.6 if LENTO else 0.7
ESPERA_PRINCIPAL = 180 if LENTO else 90     # el sistema tarda en cargar

ESCRITORIO = None        # se llena en main(): busca ventanas de todo Windows


# ---------------------------------------------------------------------------
#  PERMISOS DE ADMINISTRADOR
#  Visual Hotel pide elevacion para abrirse. Y Windows no deja que un programa
#  normal le mande clics ni teclas a una ventana de un programa elevado (se
#  llama UIPI). O sea que esto tiene que correr como administrador si o si:
#  no es una maña del script, es una regla del sistema.
# ---------------------------------------------------------------------------
def soy_administrador():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def reabrir_como_administrador():
    guion = os.path.abspath(sys.argv[0])
    argumentos = " ".join('"%s"' % a for a in sys.argv[1:])
    parametros = '"%s" %s' % (guion, argumentos)
    print("  Visual Hotel necesita permisos de administrador.")
    print("  Windows va a preguntar si permites la aplicación: dile que sí.\n")
    time.sleep(1.2)
    try:
        resultado = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, parametros, os.path.dirname(guion), 1)
    except Exception as e:
        print("  No pude pedir los permisos: %s" % e)
        return False
    if int(resultado) <= 32:
        print("  No se concedieron los permisos (código %s).\n" % resultado)
        print("  Hazlo a mano: botón derecho sobre «Informe RDS.bat»")
        print("  y elige «Ejecutar como administrador».\n")
        return False
    return True


# ---------------------------------------------------------------------------
#  AYUDAS
# ---------------------------------------------------------------------------
def aviso(paso, texto):
    print("  [%s] %s" % (paso, texto))
    sys.stdout.flush()


def plano(texto):
    """Compara sin tildes, sin '&' de los atajos y sin mayusculas."""
    t = unicodedata.normalize("NFD", str(texto or "")).encode("ascii", "ignore").decode()
    return t.replace("&", "").replace(".", "").strip().lower()


# Cada vocal acepta su version con tilde: el sistema mezcla espanol y
# portugues y los titulos no siempre estan escritos igual.
_IGUALES = {"a": "aáàâãÁÀÂÃ", "e": "eéèêÉÈÊ", "i": "iíìîÍÌÎ",
            "o": "oóòôõÓÒÔÕ", "u": "uúùûüÚÙÛÜ", "c": "cçCÇ", "n": "nñNÑ"}


def patron_flexible(texto):
    """'Parametros de Registro' -> regex que tambien calza con 'Parámetros'."""
    partes = []
    for ch in texto:
        bajo = ch.lower()
        if bajo in _IGUALES:
            partes.append("[%s%s]" % (_IGUALES[bajo], _IGUALES[bajo].upper()))
        elif ch == " ":
            partes.append(r"\s+")
        else:
            partes.append(re.escape(ch))
    return ".*" + "".join(partes) + ".*"


class ErrorPaso(Exception):
    """Falla controlada: se sabe en que paso fue y que se estaba buscando."""


def ventanas(titulo=None, clase=None):
    """Las ventanas visibles de Windows que calcen con lo pedido.

    Se busca primero por el nombre interno de la ventana (su "clase") porque
    es lo único que no cambia: el título de la principal, sin ir más lejos,
    dice "Visual Hotal" con una falta de ortografía. El título queda como
    respaldo, y descartando TApplication, que es una ventana fantasma que
    Delphi crea con el mismo nombre y no sirve para manejar nada."""
    from pywinauto import findwindows
    seguras, dudosas = [], []
    try:
        handles = findwindows.find_windows(visible_only=True, top_level_only=True)
    except Exception:
        return []
    patron = patron_flexible(titulo) if titulo else None
    for handle in handles:
        try:
            v = ESCRITORIO.window(handle=handle).wrapper_object()
            texto = v.window_text().strip()
            nombre_clase = v.friendly_class_name()
        except Exception:
            continue
        if clase and plano(nombre_clase) == plano(clase):
            seguras.append(v)
            continue
        if not texto:
            continue
        if plano(nombre_clase) == "tapplication":
            continue
        if patron and re.match(patron, texto):
            dudosas.append(v)
        elif patron is None and clase is None:
            dudosas.append(v)
    return seguras + dudosas


def listado_de_ventanas():
    lineas = []
    for v in ventanas():
        try:
            lineas.append("%s   [%s]" % (v.window_text().strip(), v.friendly_class_name()))
        except Exception:
            continue
    return lineas


def esperar_ventana(titulo, segundos, obligatoria=True, clase=None):
    """Busca la ventana en TODO Windows, no solo en el proceso que abrimos:
    estos sistemas antiguos se relanzan por el camino y cambian de proceso.
    Cada 10 segundos muestra qué hay en pantalla, para que si se queda
    esperando se vea al tiro por qué."""
    limite = time.time() + segundos
    segundo = 0
    while time.time() < limite:
        encontradas = ventanas(titulo, clase)
        if encontradas:
            if segundo:
                print()
            time.sleep(PAUSA)
            return encontradas[0]
        time.sleep(1)
        segundo += 1
        sys.stdout.write("      esperando «%s»… %ds   \r" % (titulo, segundo))
        sys.stdout.flush()
        if segundo % 10 == 0:
            print("\n      mientras tanto, en pantalla hay:")
            for linea in listado_de_ventanas():
                print("        · %s" % linea)
    if segundo:
        print()
    if not obligatoria:
        return None
    raise ErrorPaso("Esperé %d segundos y no apareció ninguna ventana que diga "
                    "«%s».\n      Las ventanas abiertas ahora mismo son:\n        %s"
                    % (segundos, titulo, "\n        ".join(listado_de_ventanas()) or "(ninguna)"))


def boton(ventana, *textos):
    """Busca un boton por su texto. Los botones de este sistema son TBitBtn y
    traen el '&' del atajo adentro ('&Ok'), asi que se compara en limpio."""
    buscados = [plano(t) for t in textos]
    encontrados = []
    for control in ventana.descendants():
        try:
            clase = control.friendly_class_name().lower()
            texto = control.window_text()
        except Exception:
            continue
        if "button" not in clase and "btn" not in clase:
            continue
        encontrados.append(texto)
        if plano(texto) in buscados:
            return control
    raise ErrorPaso("No encontré el botón %s en «%s».\n      Los botones que hay son: %s"
                    % (" / ".join(textos), ventana.window_text(),
                       ", ".join(repr(t) for t in encontrados) or "(ninguno)"))


def campos_de_texto(ventana):
    """Los cuadros donde se escribe, ordenados como se ven: de arriba a abajo."""
    campos = []
    for control in ventana.descendants():
        try:
            if "edit" not in control.friendly_class_name().lower():
                continue
            rect = control.rectangle()
            campos.append((rect.top, rect.left, control))
        except Exception:
            continue
    campos.sort(key=lambda c: (c[0], c[1]))
    return [c for _, _, c in campos]


def escribir(campo, texto, verificar=True):
    """Deja el cuadro con exactamente lo que se le pide. Se escribe de una vez
    (no letra por letra): es instantáneo y no se pierde nada por el camino.

    `verificar=False` es para los cuadros de contraseña: ahí no se puede leer
    lo que quedó escrito (muestran ●●●●●), así que comprobar sería tirar el
    texto a la basura y volver a teclearlo por gusto."""
    campo.set_focus()
    time.sleep(0.1)
    try:
        campo.set_edit_text(texto)
        time.sleep(0.15)
        if not verificar or plano(campo.window_text()) == plano(texto):
            return
    except Exception:
        pass
    # Si el cuadro no acepta que le pongan el texto de golpe (pasa con los de
    # contraseña), se teclea, pero rápido.
    campo.type_keys("^a{DEL}", pause=0.02)
    campo.type_keys(texto, with_spaces=True, pause=0.02)
    time.sleep(0.15)


def sigue_abierta(ventana):
    try:
        return bool(ctypes.windll.user32.IsWindow(ventana.handle)) and ventana.is_visible()
    except Exception:
        return False


def apretar_hasta_que_cierre(ventana, control, campo_para_enter=None, segundos=10):
    """Aprieta un botón y comprueba que la ventana efectivamente se cerró. Si
    no pasa nada, insiste de otras maneras: en estos sistemas antiguos a veces
    el clic no entra (la ventana estaba sin foco, el equipo se durmió) y hay
    que mandarle el clic como mensaje o directamente un Enter."""
    intentos = [("clic", lambda: control.click_input()),
                ("clic por mensaje", lambda: control.click())]
    if campo_para_enter is not None:
        intentos.append(("Enter", lambda: campo_para_enter.type_keys("{ENTER}")))

    for nombre, accion in intentos:
        try:
            accion()
        except Exception as e:
            aviso("...", "el %s no se pudo mandar (%s)" % (nombre, e))
            continue
        for _ in range(segundos):
            time.sleep(1)
            if not sigue_abierta(ventana):
                return True
        aviso("...", "la ventana no se cerró con el %s: pruebo otra forma" % nombre)
    return False


# ---------------------------------------------------------------------------
#  EL MENU DE ARRIBA
# ---------------------------------------------------------------------------
def es_problema_de_bits(e):
    """WinError 299 sale cuando se intenta leer la memoria de un programa de
    32 bits desde un Python de 64. No es un error nuestro: es la mezcla."""
    return "299" in str(e) or "ReadProcessMemory" in str(e)


def elegir_menu_con_teclado(ventana):
    """Sin leer nada: F10 enciende la barra de menús, las flechas caminan
    hasta Consultas y Enter abre Informes. Es lo mismo que haría una persona
    que no usa el mouse, y no necesita espiar la memoria del programa."""
    ventana.set_focus()
    time.sleep(0.4)
    ventana.type_keys("{F10}")
    time.sleep(0.5)
    ventana.type_keys("{RIGHT}" * MENU_POSICION, pause=0.12)
    time.sleep(0.4)
    ventana.type_keys("{DOWN}")            # abre el menú en su primer item
    time.sleep(0.4)
    if MENU_ITEM_POSICION:
        ventana.type_keys("{DOWN}" * MENU_ITEM_POSICION, pause=0.12)
        time.sleep(0.3)
    ventana.type_keys("{ENTER}")


def elegir_menu(ventana, camino):
    """Primero por el camino de siempre; si falla por lo que sea (tildes,
    atajos raros, el choque de 64 contra 32 bits, o el control que se demora
    en responder), se abre con el teclado, que no depende de leer nada."""
    try:
        ventana.menu_select("->".join(camino))
        return
    except Exception as e:
        motivo = "no puedo leer el menú (64 vs 32 bits)" if es_problema_de_bits(e) \
            else "el camino directo no funcionó (%s)" % e
        aviso("4/6", "%s: lo abro con el teclado" % motivo)

    try:
        elegir_menu_con_teclado(ventana)
        return
    except Exception:
        pass          # sigue con el metodo leyendo item por item, mas lento

    try:
        menu = ventana.menu()
    except Exception as e:
        raise ErrorPaso("No pude abrir el menú de ninguna forma: %s" % e)
    if menu is None:
        raise ErrorPaso("Esta ventana no tiene menú donde buscar «%s»." % camino[0])
    actual = menu.items()
    elegido = None
    for nombre in camino:
        elegido = None
        for item in actual:
            if plano(item.text()) == plano(nombre):
                elegido = item
                break
        if elegido is None:
            disponibles = ", ".join(repr(i.text()) for i in actual)
            raise ErrorPaso("En el menú no encontré «%s».\n      Hay: %s"
                            % (nombre, disponibles))
        if nombre != camino[-1]:
            elegido.select()
            time.sleep(PAUSA)
            actual = elegido.sub_menu().items()
    elegido.select()


# ---------------------------------------------------------------------------
#  EL ARBOL DE INFORMES
# ---------------------------------------------------------------------------
def arbol_de(ventana):
    for control in ventana.descendants():
        try:
            if "tree" in control.friendly_class_name().lower():
                return control
        except Exception:
            continue
    raise ErrorPaso("No encontré el listado de informes dentro de «%s»."
                    % ventana.window_text())


def hijos(nodo):
    try:
        return list(nodo.children())
    except Exception:
        return []


def buscar_nodo(nodos, nombre):
    """Primero busca el nombre exacto; si no, uno que empiece o contenga.
    El exacto va primero a propósito: en el listado conviven «… - RDS» y
    «… - RDS Modelo II», y hay que abrir el primero."""
    objetivo = plano(nombre)
    for nodo in nodos:
        if plano(nodo.text()) == objetivo:
            return nodo
    for nodo in nodos:
        texto = plano(nodo.text())
        if texto.startswith(objetivo) or objetivo in texto:
            return nodo
    return None


def numero_a_la_vista(ventana):
    """El sistema muestra «Informe Nº 2100.00-1» al costado cuando el informe
    elegido es el correcto. Leer una etiqueta suelta sí funciona aunque el
    Python sea de 64 bits, así que sirve para confirmar la selección."""
    for control in ventana.descendants():
        try:
            if NUMERO_INFORME in (control.window_text() or ""):
                return True
        except Exception:
            continue
    return False


def bajar_por_el_arbol_con_teclado(arbol):
    """Sin leer el listado: Inicio se va al primer renglón (FrontOffice), la
    flecha derecha abre la carpeta, las flechas abajo caminan hasta
    Administrales, y después se escribe el nombre del informe, que es como el
    listado busca solo mientras uno teclea."""
    arbol.set_focus()
    time.sleep(0.4)
    arbol.type_keys("{HOME}")
    time.sleep(0.3)
    arbol.type_keys("{RIGHT}")                       # abre FrontOffice
    time.sleep(0.6)
    arbol.type_keys("{DOWN}" * CARPETA_POSICION, pause=0.12)
    time.sleep(0.4)
    arbol.type_keys("{RIGHT}")                       # abre Administrales
    time.sleep(0.9)
    # Escribiendo el principio del nombre, el listado salta solo. Se llega
    # hasta la "S" de Situación, que es donde se separa del "de Operación".
    arbol.type_keys("resumen diario de s", with_spaces=True, pause=0.09)
    time.sleep(0.6)


def hijos_con_espera(nodo, intentos=6):
    """Los hijos de una rama recien abierta a veces tardan en aparecer (el
    arbol los va cargando de a poco). Se insiste varias veces antes de
    darse por vencido."""
    for _ in range(intentos):
        nodos = hijos(nodo)
        if nodos:
            return nodos
        time.sleep(PAUSA)
    return []


def bajar_por_el_arbol(arbol, camino):
    """Va abriendo carpeta por carpeta hasta llegar al informe. Primero se
    intenta con get_item, que es una funcion de pywinauto hecha justo para
    esto: recibe el camino completo y expande cada nivel ella sola, lo que
    es mas confiable que ir expandiendo y releyendo a mano."""
    try:
        item = arbol.get_item(camino, exact=False)
        item.click_input()
        time.sleep(PAUSA)
        return item
    except Exception:
        pass          # sigue con el metodo manual, mas lento pero mas claro

    nodos = arbol.roots()
    if not nodos:
        raise ErrorPaso("El listado de informes se ve vacío.\n"
                        "      Suele pasar por manejar un programa de 32 bits con un\n"
                        "      Python de 64. Instala el Python de 32 bits (ver arriba).")
    nodo = None
    for i, nombre in enumerate(camino):
        nodo = buscar_nodo(nodos, nombre)
        if nodo is None:
            disponibles = ", ".join(repr(n.text()) for n in nodos[:25]) or "(ninguno)"
            raise ErrorPaso("En el listado no encontré «%s».\n"
                            "      A ese nivel hay: %s" % (nombre, disponibles))
        aviso("arbol", "abriendo «%s»" % nodo.text().strip())
        if i < len(camino) - 1:
            try:
                nodo.expand()
            except Exception:
                nodo.click_input(double=True)
            time.sleep(PAUSA)
            nodos = hijos_con_espera(nodo)
            if not nodos:
                raise ErrorPaso("Abrí «%s» pero no le vi ningún elemento adentro,\n"
                                "      ni esperando varios segundos." % nombre.strip())
    nodo.select()
    nodo.click_input()
    time.sleep(PAUSA)
    return nodo


# ---------------------------------------------------------------------------
#  LA VENTANA QUE PIDE EL PERIODO
#  Al apretar Visualizar no sale el informe todavia: sale esta ventana con la
#  fecha inicial, la fecha final y unas casillas. El sistema propone del 1 del
#  mes hasta ayer, que suele ser lo que se quiere, asi que por defecto no se
#  toca nada y solo se aprieta Ok.
# ---------------------------------------------------------------------------
def escribir_fecha(campo, fecha):
    """Los cuadros de fecha llevan mascara dd/mm/aaaa. Se prueba con las
    barras y, si no las acepta, con los numeros pelados."""
    for texto in (fecha.strftime("%d/%m/%Y"), fecha.strftime("%d%m%Y")):
        escribir(campo, texto)
        quedo = re.sub(r"\D", "", campo.window_text())
        if quedo == fecha.strftime("%d%m%Y"):
            return True
    return False


def pedir_periodo(ventana):
    campos = campos_de_texto(ventana)
    fechas = [c for c in campos if re.match(r"^\s*\d{2}\D\d{2}\D\d{4}\s*$", c.window_text() or "")]
    if len(fechas) < 2:
        fechas = campos[:2]

    if FECHAS_RDS in ("ayer", "hoy") and len(fechas) >= 2:
        dia = datetime.date.today()
        if FECHAS_RDS == "ayer":
            dia -= datetime.timedelta(days=1)
        aviso("6/6", "Pidiendo el informe del %s" % dia.strftime("%d/%m/%Y"))
        if not (escribir_fecha(fechas[0], dia) and escribir_fecha(fechas[1], dia)):
            aviso("6/6", "No pude cambiar las fechas: sigo con las que trae.")
    else:
        desde = (fechas[0].window_text() or "?").strip() if fechas else "?"
        hasta = (fechas[1].window_text() or "?").strip() if len(fechas) > 1 else "?"
        aviso("6/6", "Período que propone el sistema: del %s al %s" % (desde, hasta))

    if not apretar_hasta_que_cierre(ventana, boton(ventana, "Ok", "Aceptar"), segundos=25):
        aviso("6/6", "La ventana del período sigue abierta; revisa la pantalla.")
        return False
    return True


# ---------------------------------------------------------------------------
#  MODO ESPIA: para cuando algo no calza y hay que ver como se llama ahora
# ---------------------------------------------------------------------------
def espiar():
    print("\n" + "=" * 74)
    print("  VENTANAS ABIERTAS AHORA")
    print("=" * 74)
    for linea in ventanas_visibles():
        print("  · %s" % linea)
    try:
        activa = ESCRITORIO.window(active_only=True, visible_only=True).wrapper_object()
        print("\n" + "=" * 74)
        print("  CONTROLES DE LA VENTANA ACTIVA: %s" % activa.window_text())
        print("=" * 74)
        activa.print_control_identifiers(depth=3)
        for control in activa.descendants():
            if "tree" in control.friendly_class_name().lower():
                print("\n  CONTENIDO DEL LISTADO (hasta 3 niveles):")
                for raiz in control.roots():
                    print("   · %s" % raiz.text())
                    for hijo in hijos_con_espera(raiz, intentos=2):
                        print("      - %s" % hijo.text())
                        for nieto in hijos_con_espera(hijo, intentos=2):
                            print("         · %s" % nieto.text())
                break
    except Exception as e:
        print("  No pude leer la ventana activa: %s" % e)


# ---------------------------------------------------------------------------
#  EL TRABAJO
# ---------------------------------------------------------------------------
def main():
    global ESCRITORIO

    if os.name != "nt":
        print("Esto solo corre en Windows: maneja ventanas de Windows.")
        return 1

    try:
        from pywinauto.application import Application
        from pywinauto import Desktop, timings
    except ImportError:
        print("\n  Falta instalar la pieza que mueve las ventanas.")
        print("  Abre la consola y escribe:\n")
        print("      py -m pip install pywinauto\n")
        return 1

    timings.Timings.window_find_timeout = 20
    timings.Timings.after_click_wait = 0.4
    ESCRITORIO = Desktop(backend="win32")

    print("=" * 74)
    print("  ABRIENDO EL RESUMEN DIARIO DE SITUACIÓN (RDS)")
    print("=" * 74)

    if not soy_administrador():
        if reabrir_como_administrador():
            print("  Sigue en la ventana nueva que se acaba de abrir.")
            return 0
        return 1

    if sys.maxsize > 2 ** 32:
        print("  Aviso: estás usando el Python de 64 bits y Visual Hotel es de 32.")
        print("  Todo funciona menos, a veces, la lectura del listado de informes.")
        print("  Si falla ahí, instala el Python de 32 bits (ver arriba del archivo).\n")

    print("  No toques el teclado ni el mouse mientras trabaja.\n")

    if not os.path.exists(RUTA_VHF):
        print("  No encuentro el programa en:\n    %s" % RUTA_VHF)
        print("  Corrige la ruta arriba en este archivo (RUTA_VHF).")
        return 1

    try:
        # ---- 1. abrir el sistema ----
        ya_abierto = esperar_ventana(TITULO_PRINCIPAL, 2, obligatoria=False,
                                     clase=CLASE_PRINCIPAL)
        if ya_abierto is not None:
            aviso("1/6", "Visual Hotel ya estaba abierto: sigo con el que hay.")
            principal = ya_abierto
        else:
            aviso("1/6", "Abriendo %s" % RUTA_VHF)
            try:
                Application(backend="win32").start(
                    RUTA_VHF, work_dir=os.path.dirname(RUTA_VHF))
            except Exception as e:
                if "740" in str(e) or "elevaci" in str(e).lower():
                    raise ErrorPaso(
                        "Windows no dejó abrir Visual Hotel: pide administrador.\n"
                        "      Cierra esto y abre «Informe RDS.bat» con el botón\n"
                        "      derecho → «Ejecutar como administrador».")
                raise
            time.sleep(PAUSA * 2)
            principal = None

        if ESPIA:
            time.sleep(3)
            espiar()
            return 0

        if principal is None:
            # ---- 2. la pantalla de parámetros ----
            # No siempre aparece (depende de una casilla del propio sistema),
            # así que se espera a la primera de las dos que salga: si el acceso
            # ya está en pantalla, no hay nada que apretar y se sigue.
            parametros = None
            limite = time.time() + 30
            while time.time() < limite:
                if ventanas(TITULO_LOGIN, CLASE_LOGIN):
                    break
                encontradas = ventanas(TITULO_PARAMETROS)
                if encontradas:
                    parametros = encontradas[0]
                    break
                time.sleep(0.5)

            if parametros is not None:
                aviso("2/6", "Pantalla de parámetros: apretando Ok")
                parametros.set_focus()
                apretar_hasta_que_cierre(parametros, boton(parametros, "Ok", "Aceptar"))
            else:
                aviso("2/6", "No hay pantalla de parámetros que apretar: sigo de largo.")

            # ---- 3. entrar con el usuario ----
            acceso = esperar_ventana(TITULO_LOGIN, 60, clase=CLASE_LOGIN)
            acceso.set_focus()
            campos = campos_de_texto(acceso)
            if len(campos) < 2:
                raise ErrorPaso("La ventana de acceso no tiene los dos cuadros "
                                "(usuario y clave) donde los esperaba.")
            aviso("3/6", "Entrando como %s" % USUARIO)
            escribir(campos[0], USUARIO)
            escribir(campos[1], CLAVE, verificar=False)
            if not apretar_hasta_que_cierre(acceso, boton(acceso, "Ok", "Aceptar"), campos[1]):
                raise ErrorPaso(
                    "Escribí el usuario y la clave, pero la ventana de acceso no se cerró\n"
                    "      ni con el clic, ni mandando el clic por mensaje, ni con Enter.\n"
                    "      Lo más probable es que el sistema esté reclamando algo (clave\n"
                    "      equivocada, usuario ocupado). Mira la pantalla: ahí hay:\n        %s"
                    % "\n        ".join(listado_de_ventanas()))

            aviso("3/6", "Adentro. Esperando a que cargue el sistema (puede tardar)")
            principal = esperar_ventana(TITULO_PRINCIPAL, ESPERA_PRINCIPAL,
                                        clase=CLASE_PRINCIPAL)

        # ---- 4. menú Consultas -> Informes ----
        principal.set_focus()
        time.sleep(PAUSA)
        aviso("4/6", "Menú %s" % " → ".join(MENU_INFORMES))
        elegir_menu(principal, MENU_INFORMES)
        time.sleep(PAUSA * 2)

        # ---- 5. el informe dentro del árbol ----
        informes = esperar_ventana(TITULO_INFORMES, 60)
        informes.set_focus()
        aviso("5/6", "Buscando el informe en el listado")
        arbol = arbol_de(informes)
        try:
            bajar_por_el_arbol(arbol, CAMINO_INFORME)
        except ErrorPaso as e:
            aviso("5/6", "El camino leyendo el listado no resultó (%s): "
                         "lo recorro con el teclado" % e)
            bajar_por_el_arbol_con_teclado(arbol)

        # El sistema muestra el número del informe elegido: si es el nuestro,
        # la selección quedó bien, se haya hecho leyendo o a ciegas.
        if numero_a_la_vista(informes):
            aviso("5/6", "Informe %s seleccionado" % NUMERO_INFORME)
        else:
            raise ErrorPaso(
                "Llegué al listado pero no logré dejar elegido el informe %s.\n"
                "      Esto se arregla del todo instalando el Python de 32 bits:\n"
                "          py install 3.13-32\n"
                "          py -V:3.13-32 -m pip install pywinauto\n"
                "      (o bajando el «Windows installer (32-bit)» de python.org)"
                % NUMERO_INFORME)

        boton(informes, "Visualizar").click_input()
        time.sleep(PAUSA * 2)

        # ---- 6. el período que pide el informe ----
        periodo = esperar_ventana(TITULO_RDS, 40, obligatoria=False)
        if periodo is None:
            aviso("6/6", "No apareció la ventana del período: el informe debería "
                         "estar en pantalla.")
        else:
            periodo.set_focus()
            pedir_periodo(periodo)

        print("\n" + "=" * 74)
        print("  LISTO: el informe quedó abierto en pantalla.")
        print("=" * 74)
        return 0

    except ErrorPaso as e:
        print("\n" + "!" * 74)
        print("  ME QUEDÉ EN ESTE PUNTO")
        print("!" * 74)
        print("  %s\n" % e)
        print("  Para ver cómo se llaman las ventanas y los botones ahora:")
        print("      py vhf_rds.py --espia\n")
        return 1
    except Exception as e:
        print("\n" + "!" * 74)
        print("  ALGO NO SALIÓ COMO SE ESPERABA")
        print("!" * 74)
        print("  %s: %s\n" % (type(e).__name__, e))
        if es_problema_de_bits(e):
            print("  Esto es el choque de 64 contra 32 bits. La solución de fondo:")
            print("      py install 3.13-32")
            print("      py -V:3.13-32 -m pip install pywinauto")
            print("  (o el «Windows installer (32-bit)» de python.org)\n")
        try:
            espiar()
        except Exception:
            pass
        print("\n  Manda esta pantalla completa y lo ajusto.")
        return 1


if __name__ == "__main__":
    codigo = main()
    print()
    try:
        input("  Aprieta Enter para cerrar esta ventana. ")
    except EOFError:
        pass
    sys.exit(codigo)
