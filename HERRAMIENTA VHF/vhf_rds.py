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

  COMO SE USA
    Doble clic en "Informe RDS.bat", o desde la consola:
        py vhf_rds.py
        py vhf_rds.py --espia    (ver que ventanas y controles hay)
        py vhf_rds.py --lento    (esperar mas entre paso y paso)

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
TITULO_PRINCIPAL = "Visual Hotel FrontOffice"
TITULO_INFORMES = "Visualizar Informes"

MENU_INFORMES = ["Consultas", "Informes"]

LENTO = "--lento" in sys.argv
ESPIA = "--espia" in sys.argv
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


def ventanas_visibles():
    from pywinauto import findwindows
    fuera = []
    for handle in findwindows.find_windows(visible_only=True):
        try:
            v = ESCRITORIO.window(handle=handle).wrapper_object()
            titulo = v.window_text().strip()
            if titulo:
                fuera.append("%s   [%s]" % (titulo, v.friendly_class_name()))
        except Exception:
            continue
    return fuera


def esperar_ventana(titulo, segundos, obligatoria=True):
    """Busca la ventana en TODO Windows, no solo en el proceso que abrimos:
    estos sistemas antiguos a veces se relanzan y cambian de proceso por el
    camino. Va marcando puntos para que se note que sigue trabajando."""
    patron = patron_flexible(titulo)
    limite = time.time() + segundos
    puntos = 0
    while time.time() < limite:
        try:
            v = ESCRITORIO.window(title_re=patron)
            if v.exists(timeout=0.4) and v.is_visible():
                if puntos:
                    print()
                time.sleep(PAUSA)
                return v.wrapper_object()
        except Exception:
            pass
        time.sleep(1)
        puntos += 1
        if puntos % 3 == 0:
            sys.stdout.write("   esperando «%s»… %ds\r" % (titulo, puntos))
            sys.stdout.flush()
    if puntos:
        print()
    if not obligatoria:
        return None
    raise ErrorPaso("Esperé %d segundos y no apareció ninguna ventana que diga "
                    "«%s».\n      Las ventanas abiertas ahora mismo son:\n        %s"
                    % (segundos, titulo, "\n        ".join(ventanas_visibles()) or "(ninguna)"))


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


def escribir(campo, texto):
    """Deja el cuadro con exactamente lo que se le pide. Se escribe de una vez
    (no letra por letra): es instantáneo y no se pierde nada por el camino."""
    campo.set_focus()
    time.sleep(0.1)
    try:
        campo.set_edit_text(texto)
        time.sleep(0.15)
        if plano(campo.window_text()) == plano(texto):
            return
    except Exception:
        pass
    # Si el cuadro no acepta que le pongan el texto de golpe (pasa con los de
    # contraseña), se teclea, pero rápido.
    campo.type_keys("^a{DEL}", pause=0.02)
    campo.type_keys(texto, with_spaces=True, pause=0.02)
    time.sleep(0.15)


# ---------------------------------------------------------------------------
#  EL MENU DE ARRIBA
# ---------------------------------------------------------------------------
def elegir_menu(ventana, camino):
    """Primero por el camino de siempre; si el menu tiene tildes o atajos que
    no calzan, se busca item por item comparando en limpio."""
    try:
        ventana.menu_select("->".join(camino))
        return
    except Exception:
        pass
    menu = ventana.menu()
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


def bajar_por_el_arbol(arbol, camino):
    """Va abriendo carpeta por carpeta hasta llegar al informe."""
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
            nodos = hijos(nodo)
            if not nodos:                      # a veces tarda en cargar la rama
                time.sleep(PAUSA * 2)
                nodos = hijos(nodo)
    nodo.select()
    nodo.click_input()
    time.sleep(PAUSA)
    return nodo


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
        activa = ESCRITORIO.window(active_only=True).wrapper_object()
        print("\n" + "=" * 74)
        print("  CONTROLES DE LA VENTANA ACTIVA: %s" % activa.window_text())
        print("=" * 74)
        activa.print_control_identifiers(depth=3)
        for control in activa.descendants():
            if "tree" in control.friendly_class_name().lower():
                print("\n  CONTENIDO DEL LISTADO:")
                for raiz in control.roots():
                    print("   · %s" % raiz.text())
                    for hijo in hijos(raiz):
                        print("      - %s" % hijo.text())
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
        ya_abierto = esperar_ventana(TITULO_PRINCIPAL, 2, obligatoria=False)
        if ya_abierto is not None:
            aviso("1/5", "Visual Hotel ya estaba abierto: sigo con el que hay.")
            principal = ya_abierto
        else:
            aviso("1/5", "Abriendo %s" % RUTA_VHF)
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
            # ---- 2. la pantalla de parámetros (no siempre aparece) ----
            parametros = esperar_ventana(TITULO_PARAMETROS, 25, obligatoria=False)
            if parametros is not None:
                aviso("2/5", "Pantalla de parámetros: apretando Ok")
                parametros.set_focus()
                boton(parametros, "Ok", "Aceptar").click_input()
                time.sleep(PAUSA)
            else:
                aviso("2/5", "No apareció la pantalla de parámetros: sigo de largo.")

            # ---- 3. entrar con el usuario ----
            acceso = esperar_ventana(TITULO_LOGIN, 60)
            acceso.set_focus()
            campos = campos_de_texto(acceso)
            if len(campos) < 2:
                raise ErrorPaso("La ventana de acceso no tiene los dos cuadros "
                                "(usuario y clave) donde los esperaba.")
            aviso("3/5", "Entrando como %s" % USUARIO)
            escribir(campos[0], USUARIO)
            escribir(campos[1], CLAVE)
            boton(acceso, "Ok", "Aceptar").click_input()

            aviso("3/5", "Esperando a que cargue el sistema (puede tardar)")
            principal = esperar_ventana(TITULO_PRINCIPAL, ESPERA_PRINCIPAL)

        # ---- 4. menú Consultas -> Informes ----
        principal.set_focus()
        time.sleep(PAUSA)
        aviso("4/5", "Menú %s" % " → ".join(MENU_INFORMES))
        elegir_menu(principal, MENU_INFORMES)
        time.sleep(PAUSA * 2)

        # ---- 5. el informe dentro del árbol ----
        informes = esperar_ventana(TITULO_INFORMES, 60)
        informes.set_focus()
        aviso("5/5", "Buscando el informe en el listado")
        bajar_por_el_arbol(arbol_de(informes), CAMINO_INFORME)
        boton(informes, "Visualizar").click_input()
        time.sleep(PAUSA * 2)

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
