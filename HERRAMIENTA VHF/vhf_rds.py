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

  ANTES DE LA PRIMERA VEZ, una sola instalacion:
        py -m pip install pywinauto

  IMPORTANTE
    - El computador tiene que estar desbloqueado y con la pantalla encendida:
      esto mueve ventanas de verdad, no trabaja escondido.
    - Mientras corre, conviene no tocar el teclado ni el mouse.
    - Si algo cambia de nombre en el sistema, el script avisa exactamente en
      que paso se quedo y con --espia se ve como se llaman las cosas ahora.
=============================================================================
"""
import os
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

# Titulos de las ventanas. Se buscan por un pedazo del titulo, no completo,
# porque la version del sistema cambia el final ("v4.38.24m").
TITULO_PARAMETROS = "Parametros de Registro"
TITULO_LOGIN = "login"
TITULO_PRINCIPAL = "Visual Hotel FrontOffice"
TITULO_INFORMES = "Visualizar Informes"

MENU_INFORMES = ["Consultas", "Informes"]

LENTO = "--lento" in sys.argv
ESPIA = "--espia" in sys.argv
PAUSA = 1.6 if LENTO else 0.7
ESPERA_LARGA = 120 if LENTO else 60


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


class ErrorPaso(Exception):
    """Falla controlada: se sabe en que paso fue y que se estaba buscando."""


def esperar(ventana, estado="ready", segundos=None):
    ventana.wait(estado, timeout=segundos or ESPERA_LARGA)
    time.sleep(PAUSA)
    return ventana


def boton(ventana, *textos):
    """Busca un boton por su texto. Los botones de este sistema son TBitBtn y
    traen el '&' del atajo adentro ('&Ok'), asi que se compara en limpio."""
    buscados = [plano(t) for t in textos]
    for control in ventana.descendants():
        try:
            clase = control.friendly_class_name()
            texto = control.window_text()
        except Exception:
            continue
        if "button" not in clase.lower() and "btn" not in clase.lower():
            continue
        if plano(texto) in buscados:
            return control
    raise ErrorPaso("No encontré el botón %s en la ventana «%s»."
                    % (" / ".join(textos), ventana.window_text()))


def campos_de_texto(ventana):
    """Los cuadros donde se escribe, ordenados como se ven: de arriba a abajo."""
    campos = []
    for control in ventana.descendants():
        try:
            clase = control.friendly_class_name().lower()
            if "edit" not in clase:
                continue
            rect = control.rectangle()
            campos.append((rect.top, rect.left, control))
        except Exception:
            continue
    campos.sort()
    return [c for _, _, c in campos]


def escribir(campo, texto):
    """Deja el cuadro con exactamente lo que se le pide."""
    campo.set_focus()
    try:
        campo.set_edit_text("")
    except Exception:
        campo.type_keys("^a{DEL}")
    time.sleep(0.15)
    campo.type_keys(texto, with_spaces=True, pause=0.04)
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
    """Primero busca el nombre exacto; si no, uno que empiece o contenga."""
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
def espiar(app):
    from pywinauto import findwindows
    print("\n" + "=" * 74)
    print("  VENTANAS ABIERTAS AHORA")
    print("=" * 74)
    for handle in findwindows.find_windows():
        try:
            v = app.window(handle=handle).wrapper_object()
            if v.window_text().strip():
                print("  · %-45s [%s]" % (v.window_text()[:45], v.friendly_class_name()))
        except Exception:
            continue
    try:
        activa = app.top_window()
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
    if os.name != "nt":
        print("Esto solo corre en Windows: maneja ventanas de Windows.")
        return 1

    try:
        from pywinauto.application import Application
        from pywinauto import timings
    except ImportError:
        print("\n  Falta instalar la pieza que mueve las ventanas.")
        print("  Abre la consola y escribe:\n")
        print("      py -m pip install pywinauto\n")
        return 1

    timings.Timings.window_find_timeout = ESPERA_LARGA
    timings.Timings.after_click_wait = 0.4

    print("=" * 74)
    print("  ABRIENDO EL RESUMEN DIARIO DE SITUACIÓN (RDS)")
    print("=" * 74)
    print("  No toques el teclado ni el mouse mientras trabaja.\n")

    if not os.path.exists(RUTA_VHF):
        print("  No encuentro el programa en:\n    %s" % RUTA_VHF)
        print("  Corrige la ruta arriba en este archivo (RUTA_VHF).")
        return 1

    app = None
    try:
        # ---- 1. abrir el sistema (o engancharse al que ya está abierto) ----
        try:
            app = Application(backend="win32").connect(path=RUTA_VHF, timeout=3)
            aviso("1/5", "Visual Hotel ya estaba abierto: me engancho a él.")
        except Exception:
            aviso("1/5", "Abriendo %s" % RUTA_VHF)
            app = Application(backend="win32").start(
                '"%s"' % RUTA_VHF, work_dir=os.path.dirname(RUTA_VHF))
            time.sleep(PAUSA * 2)

        if ESPIA:
            time.sleep(2)
            espiar(app)
            return 0

        # ---- 2. la pantalla de parámetros ----
        try:
            parametros = app.window(title_re=".*%s.*" % TITULO_PARAMETROS)
            esperar(parametros, segundos=25)
            aviso("2/5", "Pantalla de parámetros: apretando Ok")
            parametros.set_focus()
            boton(parametros, "Ok", "Aceptar").click_input()
            time.sleep(PAUSA)
        except ErrorPaso:
            raise
        except Exception:
            # En algunos computadores está marcada la opción de no mostrarla.
            aviso("2/5", "No apareció la pantalla de parámetros: sigo de largo.")

        # ---- 3. entrar con el usuario ----
        acceso = app.window(title_re="(?i).*%s.*" % TITULO_LOGIN)
        esperar(acceso)
        acceso.set_focus()
        campos = campos_de_texto(acceso)
        if len(campos) < 2:
            raise ErrorPaso("La ventana de acceso no tiene los dos cuadros "
                            "(usuario y clave) donde los esperaba.")
        aviso("3/5", "Entrando como %s" % USUARIO)
        escribir(campos[0], USUARIO)
        escribir(campos[1], CLAVE)
        boton(acceso, "Ok", "Aceptar").click_input()
        time.sleep(PAUSA * 2)

        # ---- 4. menú Consultas -> Informes ----
        principal = app.window(title_re=".*%s.*" % TITULO_PRINCIPAL)
        esperar(principal)
        principal.set_focus()
        aviso("4/5", "Menú %s" % " → ".join(MENU_INFORMES))
        elegir_menu(principal, MENU_INFORMES)
        time.sleep(PAUSA * 2)

        # ---- 5. el informe dentro del árbol ----
        informes = app.window(title_re=".*%s.*" % TITULO_INFORMES)
        esperar(informes)
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
        if app is not None:
            try:
                espiar(app)
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
