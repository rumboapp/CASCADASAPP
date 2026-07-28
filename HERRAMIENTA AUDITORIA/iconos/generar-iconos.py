# -*- coding: utf-8 -*-
"""Alternativas de icono para la herramienta de auditoria.
Todo se dibuja a 4x y se reduce, para que los bordes queden limpios."""
from PIL import Image, ImageDraw, ImageFilter

ESC = 4
GRAFITO_CL, GRAFITO_OS = (0x5A, 0x58, 0x63), (0x2E, 0x2D, 0x35)
CREMA = (0xFA, 0xF7, 0xF2)
DORADO = (0xD3, 0xA5, 0x69)
TINTA = (0x4A, 0x48, 0x52)
LOGO = Image.open("/home/user/CASCADASAPP/logo-hotel-las-cascadas8_bl-2.png").convert("RGBA").crop((41, 31, 282, 273))


def lienzo(tam):
    """Cuadrado redondeado con degradado, en resolucion de trabajo."""
    T = tam * ESC
    g = Image.new("RGB", (T, T))
    d = ImageDraw.Draw(g)
    for y in range(T):
        t = y / max(1, T - 1)
        d.line([(0, y), (T, y)], fill=tuple(
            int(a + (b - a) * t) for a, b in zip(GRAFITO_CL, GRAFITO_OS)))
    mascara = Image.new("L", (T, T), 0)
    ImageDraw.Draw(mascara).rounded_rectangle([0, 0, T - 1, T - 1], radius=round(T * 0.22), fill=255)
    base = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    base.paste(g, (0, 0), mascara)
    return base


def rematar(img, tam):
    """Filo dorado y reduccion al tamano final."""
    T = tam * ESC
    borde = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    ImageDraw.Draw(borde).rounded_rectangle(
        [0, 0, T - 1, T - 1], radius=round(T * 0.22),
        outline=DORADO + (110,), width=max(ESC, round(T * 0.012)))
    if tam >= 32:
        img = Image.alpha_composite(img, borde)
    return img.resize((tam, tam), Image.LANCZOS)


def boleta(d, T, x0, x1, y0, y1, dientes=7, grosor=1.0):
    """Boleta termica: rectangulo con el borde inferior dentado."""
    d.rounded_rectangle([x0, y0, x1, y1 - (x1 - x0) * 0.06], radius=T * 0.012, fill=CREMA)
    paso = (x1 - x0) / dientes
    puntos = [(x0, y1 - (x1 - x0) * 0.06)]
    for i in range(dientes):
        puntos.append((x0 + paso * (i + .5), y1))
        puntos.append((x0 + paso * (i + 1), y1 - (x1 - x0) * 0.06))
    d.polygon(puntos, fill=CREMA)
    # Renglones de la boleta
    alto = y1 - y0
    for i, (rel, largo) in enumerate([(.20, .74), (.34, .60), (.46, .78), (.58, .52), (.70, .70)]):
        gr = max(ESC, round(T * 0.016 * grosor))
        d.line([(x0 + (x1 - x0) * .13, y0 + alto * rel),
                (x0 + (x1 - x0) * .13 + (x1 - x0) * .74 * largo, y0 + alto * rel)],
               fill=TINTA, width=gr)


def haz(img, T, y, x0, x1, intensidad=1.0):
    """Haz de luz del escaner: resplandor amplio, nucleo grueso y filo claro.
    A tamano chico un hilo fino desaparece, asi que el nucleo va generoso."""
    brillo = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    ImageDraw.Draw(brillo).rectangle([x0 - T * .04, y - T * .075, x1 + T * .04, y + T * .075],
                                     fill=DORADO + (int(150 * intensidad),))
    brillo = brillo.filter(ImageFilter.GaussianBlur(T * 0.045))
    img = Image.alpha_composite(img, brillo)

    nucleo = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    dn = ImageDraw.Draw(nucleo)
    dn.rounded_rectangle([x0, y - T * .028, x1, y + T * .028], radius=T * .028,
                         fill=DORADO + (255,))
    dn.rounded_rectangle([x0 + T * .01, y - T * .012, x1 - T * .01, y + T * .004],
                         radius=T * .008, fill=(0xF6, 0xE2, 0xC4, 255))
    return Image.alpha_composite(img, nucleo)


# ─────────────────────────── A · boleta con haz ───────────────────────────
def icono_a(tam):
    T = tam * ESC
    img = lienzo(tam)
    capa = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    boleta(ImageDraw.Draw(capa), T, T * .30, T * .70, T * .17, T * .83)
    img = Image.alpha_composite(img, capa)
    img = haz(img, T, T * .53, T * .155, T * .845)
    return rematar(img, tam)


# ────────────────── B · boleta con el emblema del hotel ──────────────────
def icono_b(tam):
    T = tam * ESC
    img = lienzo(tam)
    capa = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    d = ImageDraw.Draw(capa)
    x0, x1, y0, y1 = T * .28, T * .72, T * .15, T * .85
    d.rounded_rectangle([x0, y0, x1, y1 - (x1 - x0) * .06], radius=T * .012, fill=CREMA)
    paso = (x1 - x0) / 7
    pts = [(x0, y1 - (x1 - x0) * .06)]
    for i in range(7):
        pts.append((x0 + paso * (i + .5), y1))
        pts.append((x0 + paso * (i + 1), y1 - (x1 - x0) * .06))
    d.polygon(pts, fill=CREMA)
    img = Image.alpha_composite(img, capa)
    # El emblema, en grafito, como va impreso en la boleta
    lado = round((x1 - x0) * .58)
    emb = LOGO.resize((lado, lado), Image.LANCZOS)
    tinta = Image.new("RGBA", emb.size, TINTA + (255,))
    tinta.putalpha(emb.split()[3])
    img.paste(tinta, (round((T - lado) / 2), round(y0 + (y1 - y0) * .09)), tinta)
    cap2 = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    d2 = ImageDraw.Draw(cap2)
    for rel, largo in [(.62, .8), (.72, .62), (.82, .74)]:
        d2.line([(x0 + (x1 - x0) * .15, y0 + (y1 - y0) * rel),
                 (x0 + (x1 - x0) * .15 + (x1 - x0) * .70 * largo, y0 + (y1 - y0) * rel)],
                fill=TINTA, width=max(ESC, round(T * .017)))
    img = Image.alpha_composite(img, cap2)
    return rematar(img, tam)


# ───────────────── C · boleta saliendo del escaner ─────────────────
def icono_c(tam):
    T = tam * ESC
    img = lienzo(tam)
    capa = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    boleta(ImageDraw.Draw(capa), T, T * .31, T * .69, T * .13, T * .70, dientes=5)
    img = Image.alpha_composite(img, capa)
    img = haz(img, T, T * .688, T * .17, T * .83, intensidad=1.5)
    # Cuerpo del escaner en dorado, para que se separe de la boleta color crema
    cuerpo = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    dc = ImageDraw.Draw(cuerpo)
    dc.rounded_rectangle([T * .14, T * .685, T * .86, T * .875], radius=T * .06,
                         fill=DORADO + (255,))
    dc.rounded_rectangle([T * .14, T * .685, T * .86, T * .755], radius=T * .06,
                         fill=(0xE4, 0xBB, 0x83, 255))
    dc.rounded_rectangle([T * .25, T * .705, T * .75, T * .735], radius=T * .015,
                         fill=GRAFITO_OS + (255,))
    img = Image.alpha_composite(img, cuerpo)
    return rematar(img, tam)


VARIANTES = [("A · boleta escaneada", icono_a),
             ("B · boleta con el logo", icono_b),
             ("C · saliendo del escáner", icono_c)]

if __name__ == "__main__":
    from PIL import ImageFont
    muestra = [160, 64, 48, 32, 16]
    fila_alto = 210
    ancho = 340 + sum(t + 30 for t in muestra)
    lamina = Image.new("RGB", (ancho, fila_alto * len(VARIANTES) + 30), (247, 244, 239))
    dl = ImageDraw.Draw(lamina)
    for i, (nombre, fn) in enumerate(VARIANTES):
        y = 30 + i * fila_alto
        dl.text((30, y + 78), nombre, fill=(60, 58, 66))
        x = 320
        for t in muestra:
            ic = fn(t)
            lamina.paste(ic, (x, y + (160 - t) // 2), ic)
            x += t + 30
    lamina.save("iconos-opciones.png")
    print("lamina lista")
