# -*- coding: utf-8 -*-
"""
Scraper linguistico de corpus.

Dada una (o varias) pagina web y una palabra objetivo, busca TODAS las veces
que esa palabra aparece en el texto (este o no en negrita en la web) y extrae
por cada aparicion:

    - palabra   : la palabra encontrada (tal como aparece en el texto)
    - oracion   : la oracion completa donde aparece, con la palabra EN NEGRITA
                  dentro del Excel para poder identificarla facilmente
    - fecha      : fecha de publicacion detectada en la pagina (si existe)
    - autor      : autor detectado en la pagina (si existe)
    - fuente     : nombre del sitio / medio
    - link       : URL de la pagina

El resultado se guarda en un archivo Excel (.xlsx) para analizarlo despues.
La palabra buscada se resalta en negrita dentro de la celda de la oracion.

Uso basico (interactivo):
    python scraper.py

Uso por linea de comandos:
    python scraper.py --url "https://ejemplo.com/articulo" --palabra "lenguaje"
    python scraper.py --url "https://a.com" "https://b.com" --palabra "che" --salida corpus.xlsx
    python scraper.py --archivo urls.txt --palabra "vos"

Rastrear una portada/seccion (descubre y scrapea sus articulos):
    python scraper.py --url "https://www.latercera.com/politica/" --palabra "punto" --crawl

Cobertura amplia por sitemap con rango de fechas (recomendado para corpus):
    python scraper.py --sitemap "https://www.latercera.com/arc/outboundfeeds/sitemap-index/?outputType=xml" ^
        --palabra "punto" --desde 2025-01-01 --hasta 2026-08-15 --max-articulos 800 --salida corpus.xlsx

Prueba rapida (QA) cortando en 30 apariciones:
    python scraper.py --url "https://el-sitio.com/nota" --palabra "punto" --limite 30

Columnas del Excel: palabra | oracion (palabra en negrita) | categoria (seccion) |
                    fecha | autor | fuente | link

Notas:
    - La busqueda ignora mayusculas/minusculas y tildes por defecto (util en espanol),
      y detecta la palabra completa (no cuenta "punto" dentro de "puntos").
    - La negrita es del RESULTADO en Excel, no un filtro sobre la web.
"""

import argparse
import os
import re
import sys
import time
import unicodedata
from datetime import date, datetime
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuracion
# ---------------------------------------------------------------------------

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
}

# Etiquetas de bloque que contienen texto "de oracion" (hojas de contenido).
# Evitamos div/section/article genericos para no duplicar bloques grandes.
BLOQUES_TEXTO = ["p", "li", "blockquote", "figcaption", "caption",
                 "dd", "dt", "td", "th",
                 "h1", "h2", "h3", "h4", "h5", "h6"]

TIEMPO_ESPERA = 20  # segundos de timeout por peticion

# Etiquetas cuyo contenido NO es texto de lectura (se descartan).
BASURA = {"script", "style", "noscript", "template", "svg", "nav", "footer", "header", "aside"}


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------

def normalizar(texto: str) -> str:
    """Pasa a minusculas y quita tildes/diacriticos para comparar."""
    if not texto:
        return ""
    texto = texto.lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto


def limpiar_espacios(texto: str) -> str:
    """Colapsa espacios/saltos de linea multiples en uno solo."""
    return re.sub(r"\s+", " ", texto or "").strip()


def limpiar_puntuacion(texto: str) -> str:
    """Devuelve la oracion mas 'literal': quita espacios sobrantes antes de signos
    de puntuacion y dentro de parentesis/comillas."""
    texto = re.sub(r"\s+([,.;:!?%…])", r"\1", texto)   # sin espacio antes del signo
    texto = re.sub(r"\(\s+", "(", texto)
    texto = re.sub(r"\s+\)", ")", texto)
    texto = re.sub(r"«\s+", "«", texto)
    texto = re.sub(r"\s+»", "»", texto)
    return texto.strip()


# Abreviaturas frecuentes: evitan cortar la oracion en el punto que las sigue.
ABREVIATURAS = {
    "sr", "sra", "srta", "dr", "dra", "ee", "uu", "etc", "vs", "av", "avda",
    "gob", "art", "pag", "pag", "cap", "fig", "min", "max", "aprox", "ex",
    "no", "num", "tel", "ss", "op", "cit", "ed", "vol", "col", "d", "a", "c",
}


def dividir_en_oraciones(texto: str):
    """Divide un parrafo en oraciones respetando abreviaturas y numeros decimales."""
    texto = limpiar_espacios(texto)
    oraciones = []
    inicio = 0
    # Corta tras . ! ? ... solo si sigue espacio y luego mayuscula / signo de apertura.
    for m in re.finditer(r"[\.\!\?…]+\s+(?=[«\"“¿¡(]?[A-ZÁÉÍÓÚÑ])", texto):
        prev = texto[:m.start()]
        ult = re.search(r"(\w+)[\.\!\?…]*$", prev + texto[m.start():m.start()+1])
        token = ""
        mt = re.search(r"(\w+)$", prev)
        if mt:
            token = mt.group(1).lower()
        # No cortar tras abreviatura o inicial de una sola letra.
        if token in ABREVIATURAS or len(token) == 1:
            continue
        fin = m.start() + len(re.match(r"[\.\!\?…]+", texto[m.start():]).group(0))
        oraciones.append(texto[inicio:fin].strip())
        inicio = m.end()
    resto = texto[inicio:].strip()
    if resto:
        oraciones.append(resto)
    return [limpiar_puntuacion(o) for o in oraciones if o.strip()]


def encontrar_apariciones(texto: str, palabra_norm: str):
    """Devuelve las posiciones (inicio, fin) en `texto` donde aparece la palabra
    como palabra completa, ignorando mayusculas y tildes.

    Mapea las posiciones de la version normalizada de vuelta al texto original,
    para poder resaltar exactamente el fragmento correcto.
    """
    if not texto or not palabra_norm:
        return []

    # Normalizamos caracter por caracter guardando el indice original de cada
    # caracter resultante (quitar tildes puede cambiar la longitud).
    norm_chars = []
    idx_map = []
    for i, ch in enumerate(texto):
        dec = unicodedata.normalize("NFD", ch.lower())
        dec = "".join(c for c in dec if unicodedata.category(c) != "Mn")
        for c in dec:
            norm_chars.append(c)
            idx_map.append(i)

    norm_str = "".join(norm_chars)
    patron = r"(?<!\w)" + re.escape(palabra_norm) + r"(?!\w)"

    spans = []
    for m in re.finditer(patron, norm_str):
        if m.end() == 0:
            continue
        ini = idx_map[m.start()]
        fin = idx_map[m.end() - 1] + 1
        spans.append((ini, fin))
    return spans


# ---------------------------------------------------------------------------
# Extraccion de metadatos (fecha, autor, fuente)
# ---------------------------------------------------------------------------

def _meta(soup, **atributos):
    """Devuelve el content del primer <meta> que coincida con los atributos dados."""
    tag = soup.find("meta", attrs=atributos)
    if tag and tag.get("content"):
        return limpiar_espacios(tag["content"])
    return None


def extraer_fecha(soup) -> str:
    candidatos = [
        _meta(soup, property="article:published_time"),
        _meta(soup, property="article:modified_time"),
        _meta(soup, itemprop="datePublished"),
        _meta(soup, attrs={"name": "date"}),
        _meta(soup, attrs={"name": "publish-date"}),
        _meta(soup, attrs={"name": "pubdate"}),
        _meta(soup, property="og:updated_time"),
    ]
    for c in candidatos:
        if c:
            return _formatear_fecha(c)

    # Etiqueta <time datetime="...">
    time_tag = soup.find("time")
    if time_tag:
        if time_tag.get("datetime"):
            return _formatear_fecha(time_tag["datetime"])
        if time_tag.get_text(strip=True):
            return limpiar_espacios(time_tag.get_text())

    return ""


def _formatear_fecha(valor: str) -> str:
    """Intenta normalizar una fecha ISO a YYYY-MM-DD; si no puede, devuelve el original."""
    valor = valor.strip()
    m = re.match(r"(\d{4}-\d{2}-\d{2})", valor)
    if m:
        return m.group(1)
    for fmt in ("%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(valor[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return valor


def extraer_autor(soup) -> str:
    candidatos = [
        _meta(soup, attrs={"name": "author"}),
        _meta(soup, property="article:author"),
        _meta(soup, itemprop="author"),
        _meta(soup, attrs={"name": "sailthru.author"}),
        _meta(soup, attrs={"name": "twitter:creator"}),
    ]
    for c in candidatos:
        if c:
            return c

    # Patrones comunes en el HTML.
    for selector in [
        {"itemprop": "author"},
        {"rel": "author"},
        {"class": re.compile(r"author|autor|byline|firma", re.I)},
    ]:
        tag = soup.find(attrs=selector)
        if tag:
            texto = limpiar_espacios(tag.get_text())
            if texto and len(texto) < 120:
                return texto

    return ""


def extraer_fuente(soup, url: str) -> str:
    fuente = _meta(soup, property="og:site_name")
    if fuente:
        return fuente
    if soup.title and soup.title.get_text(strip=True):
        # A veces el titulo es "Nota | Medio"; tomamos la ultima parte.
        titulo = limpiar_espacios(soup.title.get_text())
        for sep in ("|", "-", "—", "·", "»"):
            if sep in titulo:
                return titulo.split(sep)[-1].strip()
    # Fallback: el dominio.
    return urlparse(url).netloc.replace("www.", "")


def extraer_categoria(soup, url: str) -> str:
    """Seccion/categoria de la nota. Prefiere los metadatos; si no, usa el primer
    segmento de la ruta de la URL (ej. /el-deportivo/noticia/... -> 'el-deportivo')."""
    for c in [
        _meta(soup, property="article:section"),
        _meta(soup, attrs={"name": "section"}),
        _meta(soup, itemprop="articleSection"),
    ]:
        if c:
            return c

    segmentos = [s for s in urlparse(url).path.split("/") if s]
    IGNORAR = {"noticia", "articulo", "nota", "story", "news", "amp"}
    for seg in segmentos:
        if seg.lower() not in IGNORAR:
            return seg
    return ""


# ---------------------------------------------------------------------------
# Descubrimiento de articulos (crawling desde una portada / seccion)
# ---------------------------------------------------------------------------

# Rutas que NO son notas (indices de autor, tags, secciones, etc.).
RUTA_EXCLUIDA = re.compile(
    r"/(autor|autores|tag|tags|etiqueta|seccion|secciones|categoria|"
    r"buscar|search|suscri|newsletter|podcast|video|videos|multimedia|en-vivo)/",
    re.I,
)


def es_link_articulo(url: str, base_netloc: str) -> bool:
    """Heuristica para decidir si un link es una nota (y no un indice/seccion)."""
    p = urlparse(url)
    if base_netloc not in p.netloc:
        return False
    ruta = p.path
    if RUTA_EXCLUIDA.search(ruta):
        return False
    # Senal fuerte: muchos medios usan /noticia/ o /articulo/ en la ruta.
    if re.search(r"/(noticia|articulo|nota|story|news)/", ruta, re.I):
        return True
    # Senal general: el ultimo segmento es un slug largo con guiones.
    ultimo = ruta.rstrip("/").split("/")[-1]
    if ultimo.count("-") >= 3 and len(ultimo) > 20:
        return True
    return False


def descubrir_articulos(url_indice: str, session: requests.Session, maximo: int = 60):
    """Descarga una portada/seccion y devuelve una lista de URLs de articulos."""
    try:
        resp = session.get(url_indice, headers=HEADERS, timeout=TIEMPO_ESPERA)
        resp.raise_for_status()
    except requests.RequestException as e:
        print(f"  [ERROR] No se pudo descargar el indice {url_indice}: {e}")
        return []

    if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
        resp.encoding = resp.apparent_encoding

    soup = BeautifulSoup(resp.text, "lxml")
    base_netloc = urlparse(url_indice).netloc.replace("www.", "")

    encontrados = []
    vistos = set()
    for a in soup.find_all("a", href=True):
        href = urljoin(url_indice, a["href"]).split("?")[0].split("#")[0]
        if href in vistos:
            continue
        if es_link_articulo(href, base_netloc):
            vistos.add(href)
            encontrados.append(href)
        if len(encontrados) >= maximo:
            break

    print(f"  -> {len(encontrados)} articulo(s) descubierto(s) en la portada/seccion.")
    return encontrados


# ---------------------------------------------------------------------------
# Descubrimiento por sitemap (cobertura amplia + filtro por fecha)
# ---------------------------------------------------------------------------

def _parse_fecha_iso(valor: str):
    """Convierte 'AAAA-MM-DDT...' o 'AAAA-MM-DD' a un objeto date (o None)."""
    if not valor:
        return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", valor.strip())
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def _leer_xml(url: str, session: requests.Session):
    """Descarga y parsea un XML de sitemap. Devuelve (tipo, entradas).

    tipo == 'index'  -> entradas = [url_subsitemap, ...]
    tipo == 'urlset' -> entradas = [(loc, fecha_date_o_None), ...]
    """
    try:
        resp = session.get(url, headers=HEADERS, timeout=TIEMPO_ESPERA)
        resp.raise_for_status()
    except requests.RequestException as e:
        print(f"  [ERROR] No se pudo descargar el sitemap {url}: {e}")
        return None, []

    soup = BeautifulSoup(resp.content, "xml")

    if soup.find("sitemapindex"):
        subs = [s.get_text(strip=True) for s in soup.find_all("loc")]
        return "index", subs

    entradas = []
    for u in soup.find_all("url"):
        loc = u.find("loc")
        if not loc:
            continue
        lastmod = u.find("lastmod")
        f = _parse_fecha_iso(lastmod.get_text()) if lastmod else None
        entradas.append((loc.get_text(strip=True), f))
    return "urlset", entradas


def descubrir_por_sitemap(sitemap_url: str, session: requests.Session,
                          desde: date = None, hasta: date = None,
                          maximo: int = 500):
    """Recorre un sitemap (indice o urlset) y devuelve URLs de articulos dentro del
    rango de fechas [desde, hasta]. Aprovecha que los sitemaps van de mas nuevo a
    mas viejo para cortar temprano cuando ya se paso de 'desde'."""
    base_netloc = urlparse(sitemap_url).netloc.replace("www.", "")
    tipo, entradas = _leer_xml(sitemap_url, session)

    # Lista de sub-sitemaps a recorrer (si es indice) o el propio urlset.
    if tipo == "index":
        subs = entradas
        print(f"  -> indice con {len(subs)} sub-sitemap(s).")
    elif tipo == "urlset":
        subs = [sitemap_url]  # lo re-leemos abajo de forma uniforme
    else:
        return []

    articulos = []
    vistos = set()

    for i, sub in enumerate(subs, 1):
        if tipo == "index":
            _, urls = _leer_xml(sub, session)
        else:
            urls = entradas  # ya lo tenemos

        if not urls:
            continue

        fechas_validas = [f for _, f in urls if f]
        mas_nueva = max(fechas_validas) if fechas_validas else None

        agregados_sub = 0
        paso_de_desde = False
        for loc, f in urls:
            if RUTA_EXCLUIDA.search(urlparse(loc).path):
                continue
            if base_netloc not in urlparse(loc).netloc:
                continue
            if hasta and f and f > hasta:
                continue
            if desde and f and f < desde:
                paso_de_desde = True
                continue
            if loc in vistos:
                continue
            vistos.add(loc)
            articulos.append(loc)
            agregados_sub += 1
            if len(articulos) >= maximo:
                print(f"  -> {len(articulos)} articulo(s) reunidos (tope alcanzado).")
                return articulos

        if tipo == "index":
            print(f"     [{i}/{len(subs)}] {sub.split('from=')[-1] if 'from=' in sub else sub} "
                  f"-> {agregados_sub} en rango (acumulado {len(articulos)})")

        # Corte temprano: si toda esta tanda ya es mas vieja que 'desde',
        # las siguientes (mas viejas aun) tampoco serviran.
        if desde and mas_nueva and mas_nueva < desde:
            print("  -> se alcanzo el limite de antiguedad; corto la busqueda.")
            break
        if tipo == "urlset":
            break

    print(f"  -> {len(articulos)} articulo(s) reunidos desde el sitemap.")
    return articulos


# ---------------------------------------------------------------------------
# Descubrimiento por Wayback Machine (Internet Archive) -> cobertura de ANIOS
# ---------------------------------------------------------------------------
#
# El sitemap de un medio solo lista lo reciente (en La Tercera, ~2 meses / 10.000
# notas). Para armar un corpus con anios de profundidad usamos la CDX API del
# Internet Archive, que tiene URLs archivadas del sitio desde hace mucho (La
# Tercera esta desde 1998). Devolvemos, por cada articulo historico:
#     - link     : la URL original (para citar en el corpus)
#     - descarga : la URL del snapshot archivado (de ahi se baja el HTML)
#     - fecha    : fecha del snapshot (fallback si la nota no trae metadato)

CDX_API = "http://web.archive.org/cdx/search/cdx"

# Secciones tipicas de La Tercera (base para el prefijo del CDX). Se pueden
# sobreescribir por CLI con --secciones. Cada seccion usa URLs del tipo
# latercera.com/<seccion>/noticia/<slug>, que es lo que rastreamos.
SECCIONES_LATERCERA = [
    "nacional", "politica", "negocios", "pulso", "el-deportivo", "mundo",
    "tendencias", "culto", "la-tercera-pm", "que-pasa", "practico", "paula",
    "mouse", "finde", "opinion", "santiago", "reconstitucion",
]

# Por que via CDX por PREFIJO + PAGINACION (y no por rango de fechas):
#   - Las consultas CDX con filtro de fecha o 'collapse' sobre todo el dominio
#     hacen un escaneo en vivo y se caen con 504/timeout en dominios grandes.
#   - En cambio, 'url=<dominio>/<seccion>&matchType=prefix&page=N' usa el indice
#     precomputado: responde en ~1-2s por pagina y cada pagina trae ~2.500 URLs
#     unicas de articulos que abarcan varios anios (2018+ en La Tercera).
#   - Filtramos por fecha y deduplicamos del lado del cliente.


def _norm_url_clave(url: str) -> str:
    """Clave para deduplicar URLs (sin esquema, www, puerto, query ni barra final)."""
    p = urlparse(url)
    host = p.netloc.replace("www.", "").split(":")[0]
    return (host + p.path.rstrip("/")).lower()


def _snapshot_url(timestamp: str, original: str) -> str:
    """URL del snapshot 'crudo' (sufijo id_ = sin la barra de Wayback)."""
    return f"https://web.archive.org/web/{timestamp}id_/{original}"


# Senales de URL "basura": el archivo a veces guarda URLs que aparecian dentro
# de JSON/JavaScript incrustado en la pagina, con comillas/comas codificadas.
_BASURA_URL = ('%22', '%2c', '%2C', '%7b', '%7d', '%3e', '%3c', '%5c',
               '"', ',', '{', '}', '\\', '<', '>')


def _articulo_wayback_valido(original: str) -> bool:
    """Filtra URLs de CDX: exige una nota /noticia/<slug> real y descarta la
    basura embebida en JSON/JS (comillas, comas, llaves codificadas, etc.)."""
    if any(b in original for b in _BASURA_URL):
        return False
    path = urlparse(original).path
    if "/noticia/" not in path:
        return False
    slug = path.split("/noticia/", 1)[1].strip("/")
    # Debe haber un slug con guiones y largo razonable (no el indice vacio).
    primer = slug.split("/")[0] if slug else ""
    return len(primer) >= 12 and "-" in primer


# Timeout (conexion, lectura). El de lectura corta si el servidor no manda datos,
# evitando cuelgues indefinidos cuando el Internet Archive esta saturado.
CDX_TIMEOUT = (15, 45)
CDX_REINTENTOS = 3


def _cdx_get(params: dict, session: requests.Session):
    """GET a la CDX API con reintentos y backoff. Devuelve el Response o None."""
    for intento in range(1, CDX_REINTENTOS + 1):
        try:
            r = session.get(CDX_API, params=params, headers=HEADERS, timeout=CDX_TIMEOUT)
            if r.status_code == 200 and "Gateway Time-out" not in r.text[:200]:
                return r
            # 429/503/504 -> el archivo esta saturado o limitando; esperamos.
        except requests.RequestException:
            pass
        if intento < CDX_REINTENTOS:
            espera = 5 * intento
            print(f"       (CDX lento/limitando; reintento {intento} en {espera}s)")
            time.sleep(espera)
    return None


def _cdx_seccion_fechas(prefijo: str, session: requests.Session,
                        desde: date = None, hasta: date = None,
                        limite: int = 200000):
    """UNA sola consulta a CDX por seccion, filtrando por fecha del lado del
    SERVIDOR (params from/to). Devuelve las filas [original, timestamp] ya
    acotadas al rango.

    Por que asi y no paginando el indice:
      - showNumPages/page IGNORAN el filtro de fecha (devuelven el indice
        completo, ~cientos de paginas), lo que obliga a barrer todo y dispara
        el throttling del Archive.
      - En cambio 'from'/'to' sobre el prefijo de seccion se resuelven server-
        side y traen solo los snapshots del rango en un unico request liviano.
    """
    params = {
        "url": prefijo, "matchType": "prefix",
        "output": "json", "fl": "original,timestamp", "collapse": "urlkey",
        "limit": str(limite),
    }
    if desde:
        params["from"] = desde.strftime("%Y%m%d")
    if hasta:
        params["to"] = hasta.strftime("%Y%m%d")
    r = _cdx_get(params, session)
    if not r:
        return []
    try:
        data = r.json()
    except ValueError:
        return []
    if data and data[0] and data[0][0] == "original":
        data = data[1:]
    return data


def descubrir_por_wayback(dominio: str, session: requests.Session,
                          desde: date = None, hasta: date = None,
                          maximo: int = 500, pausa: float = 1.0,
                          secciones=None):
    """Reune URLs historicas de articulos de `dominio` via CDX API (Internet
    Archive). Hace UNA consulta por seccion con filtro de fecha del lado del
    servidor (from/to), evitando barrer todo el indice y el throttling. Filtra
    a notas /noticia/ validas, deduplica y, si hay cuota por seccion, toma una
    muestra REPARTIDA en el tiempo para cubrir todo el rango de fechas.

    `maximo`: tope total de articulos (0 = sin tope, trae TODO el rango).
    Devuelve una lista de dicts: {"link", "descarga", "fecha"}.
    """
    base = dominio.replace("www.", "")
    if secciones is None:
        secciones = SECCIONES_LATERCERA if "latercera" in base else []
    if not secciones:
        print("  [AVISO] No hay secciones para este dominio. Usa --secciones sec1 sec2 ...")
        return []

    resultados = []
    vistos = set()
    # Cuota por seccion para repartir el corpus entre secciones (mas variedad).
    cuota = max(1, maximo // len(secciones)) if maximo else 0

    for seccion in secciones:
        if maximo and len(resultados) >= maximo:
            break
        prefijo = f"{base}/{seccion}"
        filas = _cdx_seccion_fechas(prefijo, session, desde=desde, hasta=hasta)
        print(f"  [{seccion}] {len(filas)} snapshot(s) en rango; filtrando...")

        # 1) Filtrar a notas validas + deduplicar dentro de la seccion.
        candidatas = []
        vistos_sec = set()
        for fila in filas:
            if not fila or len(fila) < 2:
                continue
            original, timestamp = fila[0], fila[1]
            if not _articulo_wayback_valido(original):
                continue
            if RUTA_EXCLUIDA.search(urlparse(original).path):
                continue
            f = _parse_fecha_iso(f"{timestamp[:4]}-{timestamp[4:6]}-{timestamp[6:8]}")
            if desde and f and f < desde:
                continue
            if hasta and f and f > hasta:
                continue
            clave = _norm_url_clave(original)
            if clave in vistos or clave in vistos_sec:
                continue
            vistos_sec.add(clave)
            candidatas.append({
                "link": original,
                "descarga": _snapshot_url(timestamp, original),
                "fecha": f,
            })

        # 2) Ordenar por fecha para que un eventual muestreo cubra todo el rango.
        candidatas.sort(key=lambda a: (a["fecha"] is None, a["fecha"] or date.min))

        # 3) Cuanto tomar de esta seccion: la cuota (si hay) o todo, acotado por
        #    el presupuesto global restante.
        limite_sec = cuota if cuota else len(candidatas)
        if maximo:
            limite_sec = min(limite_sec, maximo - len(resultados))

        # 4) Si sobran candidatas, muestrear REPARTIDO en el tiempo (paso fijo).
        if 0 < limite_sec < len(candidatas):
            paso = len(candidatas) / limite_sec
            candidatas = [candidatas[int(i * paso)] for i in range(limite_sec)]

        agregados_sec = 0
        for a in candidatas:
            if maximo and len(resultados) >= maximo:
                break
            vistos.add(_norm_url_clave(a["link"]))
            resultados.append(a)
            agregados_sec += 1

        print(f"     -> {agregados_sec} de '{seccion}' (acumulado {len(resultados)}).")
        time.sleep(pausa)

    print(f"  -> {len(resultados)} articulo(s) historicos reunidos via Wayback.")
    return resultados


# ---------------------------------------------------------------------------
# Nucleo del scraping
# ---------------------------------------------------------------------------

def procesar_pagina(url: str, palabra: str, session: requests.Session, limite: int = 0,
                    url_descarga: str = None, fecha_fallback=None):
    """Descarga y analiza una pagina; devuelve lista de dicts (una fila por oracion
    que contiene la palabra).

    - `url` es el link que se guarda en el corpus (la URL original a citar).
    - `url_descarga` (opcional) es de donde se baja realmente el HTML; util para
      snapshots de Wayback (se cita el original pero se lee el archivado).
    - `fecha_fallback` (date) se usa si la pagina no trae metadato de fecha.
    - Si `limite` > 0, deja de recolectar apenas alcanza esa cantidad de filas.
    """
    palabra_norm = normalizar(palabra)
    filas = []

    # Orden de intento: primero la pagina EN VIVO (rapida y suele seguir online),
    # y si falla, el snapshot archivado (mas lento, pero rescata notas caidas).
    intentos = [(url, TIEMPO_ESPERA)]
    if url_descarga:
        intentos.append((url_descarga, TIEMPO_ESPERA * 3))

    resp = None
    for origen, timeout in intentos:
        try:
            r = session.get(origen, headers=HEADERS, timeout=timeout)
            r.raise_for_status()
            resp = r
            break
        except requests.RequestException as e:
            ultimo_error = e
    if resp is None:
        print(f"  [ERROR] No se pudo descargar {url}: {ultimo_error}")
        return filas

    # Deja que requests deduzca la codificacion correcta.
    if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
        resp.encoding = resp.apparent_encoding

    soup = BeautifulSoup(resp.text, "lxml")

    # Quita elementos que no son texto de lectura.
    for tag in soup(list(BASURA)):
        tag.decompose()

    # Metadatos comunes a toda la pagina.
    fecha = extraer_fecha(soup)
    if not fecha and fecha_fallback:
        fecha = fecha_fallback.strftime("%Y-%m-%d")
    autor = extraer_autor(soup)
    fuente = extraer_fuente(soup, url)
    categoria = extraer_categoria(soup, url)

    vistos = set()  # evita oraciones duplicadas

    for bloque in soup.find_all(BLOQUES_TEXTO):
        texto = limpiar_espacios(bloque.get_text(" "))
        if not texto:
            continue

        for oracion in dividir_en_oraciones(texto):
            spans = encontrar_apariciones(oracion, palabra_norm)
            if not spans:
                continue

            clave = normalizar(oracion)
            if clave in vistos:
                continue
            vistos.add(clave)

            # La "palabra" es el fragmento tal como aparece en el texto original.
            ini, fin = spans[0]
            palabra_encontrada = oracion[ini:fin]

            filas.append({
                "palabra": palabra_encontrada,
                "oracion": oracion,
                "spans": spans,        # posiciones para resaltar en el Excel
                "categoria": categoria,
                "fecha": fecha,
                "autor": autor,
                "fuente": fuente,
                "link": url,
            })

            if limite and len(filas) >= limite:
                print(f"  -> {len(filas)} aparicion(es) encontradas (limite alcanzado).")
                return filas

    print(f"  -> {len(filas)} aparicion(es) encontradas.")
    return filas


# ---------------------------------------------------------------------------
# Guardado en Excel (con la palabra resaltada en negrita)
# ---------------------------------------------------------------------------

def _oracion_rich(oracion: str, spans):
    """Construye una celda con rich text: la palabra buscada en negrita y el
    resto en formato normal."""
    negrita = InlineFont(b=True)
    partes = []
    cursor = 0
    for ini, fin in spans:
        if ini > cursor:
            partes.append(oracion[cursor:ini])          # texto normal antes
        partes.append(TextBlock(negrita, oracion[ini:fin]))  # palabra en negrita
        cursor = fin
    if cursor < len(oracion):
        partes.append(oracion[cursor:])                  # resto normal
    return CellRichText(*partes)


def guardar_excel(filas, salida: str):
    encabezados = ["palabra", "oracion", "categoria", "fecha", "autor", "fuente", "link"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Corpus"

    # Encabezados en negrita.
    ws.append(encabezados)
    for celda in ws[1]:
        celda.font = Font(bold=True)

    for f in filas:
        fila = ws.max_row + 1
        ws.cell(row=fila, column=1, value=f["palabra"])
        ws.cell(row=fila, column=2, value=_oracion_rich(f["oracion"], f["spans"]))
        ws.cell(row=fila, column=3, value=f.get("categoria", ""))
        ws.cell(row=fila, column=4, value=f["fecha"])
        ws.cell(row=fila, column=5, value=f["autor"])
        ws.cell(row=fila, column=6, value=f["fuente"])
        ws.cell(row=fila, column=7, value=f["link"])

    # Ancho de columnas y ajuste de texto en la oracion.
    anchos = [18, 90, 16, 14, 26, 22, 50]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    wrap = Alignment(wrap_text=True, vertical="top")
    for fila in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        for celda in fila:
            celda.alignment = wrap

    ws.freeze_panes = "A2"

    # Si el archivo esta abierto (bloqueado por Excel), guardamos con otro nombre.
    try:
        wb.save(salida)
    except PermissionError:
        raiz, ext = os.path.splitext(salida)
        alterno = f"{raiz}_{datetime.now().strftime('%H%M%S')}{ext}"
        wb.save(alterno)
        print(f"\n[AVISO] '{salida}' estaba abierto/bloqueado. Guarde en '{alterno}'.")
        salida = alterno

    print(f"\nListo: {len(filas)} fila(s) guardadas en '{salida}'.")


# ---------------------------------------------------------------------------
# Entrada / CLI
# ---------------------------------------------------------------------------

def leer_urls_de_archivo(ruta: str):
    urls = []
    with open(ruta, "r", encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip()
            if linea and not linea.startswith("#"):
                urls.append(linea)
    return urls


def main():
    parser = argparse.ArgumentParser(
        description="Busca una palabra en paginas web y guarda cada oracion en Excel, con la palabra en negrita."
    )
    parser.add_argument("--url", nargs="+", help="Una o mas URLs a analizar.")
    parser.add_argument("--archivo", help="Ruta a un .txt con una URL por linea.")
    parser.add_argument("--palabra", help="Palabra a buscar.")
    parser.add_argument("--salida", default="corpus.xlsx", help="Archivo Excel de salida (por defecto: corpus.xlsx).")
    parser.add_argument("--espera", type=float, default=1.0, help="Segundos de espera entre paginas (por defecto: 1).")
    parser.add_argument("--limite", type=int, default=0,
                        help="Corta el proceso al llegar a N apariciones en total (0 = sin limite). Util para pruebas QA.")
    parser.add_argument("--crawl", action="store_true",
                        help="Trata las URLs como portadas/secciones: descubre y scrapea sus articulos.")
    parser.add_argument("--sitemap", nargs="+",
                        help="URL(s) de sitemap (indice o urlset) para cobertura amplia por fechas.")
    parser.add_argument("--wayback", nargs="+",
                        help="Dominio(s) (ej. latercera.com) para reunir articulos HISTORICOS de anios "
                             "via Internet Archive. Usa --desde/--hasta para acotar los anios.")
    parser.add_argument("--secciones", nargs="+",
                        help="(solo --wayback) Secciones a rastrear, ej: nacional politica pulso. "
                             "Si se omite en La Tercera, usa una lista por defecto.")
    parser.add_argument("--desde", help="Fecha minima AAAA-MM-DD (solo articulos desde esta fecha).")
    parser.add_argument("--hasta", help="Fecha maxima AAAA-MM-DD (solo articulos hasta esta fecha).")
    parser.add_argument("--max-articulos", type=int, default=60,
                        help="Maximo de articulos a reunir (por portada en --crawl, o en total en --sitemap/--wayback).")
    parser.add_argument("--guardar-cada", type=int, default=25,
                        help="Guarda el Excel de forma incremental cada N articulos procesados (0 = solo al final).")
    args = parser.parse_args()

    # Modo interactivo si faltan datos.
    urls = []
    if args.url:
        urls.extend(args.url)
    if args.archivo:
        urls.extend(leer_urls_de_archivo(args.archivo))
    if not urls and not args.sitemap and not args.wayback:
        entrada = input("Pega la(s) URL(s) separadas por espacio: ").strip()
        urls = entrada.split()

    palabra = args.palabra
    if not palabra:
        palabra = input("Palabra a buscar: ").strip()

    if (not urls and not args.sitemap and not args.wayback) or not palabra:
        print("Necesito al menos una URL (o --sitemap / --wayback) y una palabra. Saliendo.")
        sys.exit(1)

    desde = _parse_fecha_iso(args.desde) if args.desde else None
    hasta = _parse_fecha_iso(args.hasta) if args.hasta else None
    if args.desde and not desde:
        print(f"Fecha --desde invalida: {args.desde} (usa AAAA-MM-DD). Saliendo.")
        sys.exit(1)
    if args.hasta and not hasta:
        print(f"Fecha --hasta invalida: {args.hasta} (usa AAAA-MM-DD). Saliendo.")
        sys.exit(1)

    if args.limite:
        print(f"\nModo QA: se cortara al llegar a {args.limite} aparicion(es).")

    session = requests.Session()

    # Cada "trabajo" es un dict: {"link", "descarga"(opcional), "fecha"(opcional)}.
    # descarga = de donde se baja el HTML (snapshot de Wayback); si falta, se usa link.
    trabajos = []
    vistos = set()

    # 1) Modo wayback: articulos historicos (anios) via Internet Archive.
    if args.wayback:
        rango = f" (anios {desde.year if desde else '...'} a {hasta.year if hasta else date.today().year})"
        print(f"\nModo wayback: reuniendo articulos historicos de {len(args.wayback)} dominio(s){rango}...\n")
        for dom in args.wayback:
            dom = dom.replace("https://", "").replace("http://", "").strip("/")
            print(f"[wayback] {dom}")
            for art in descubrir_por_wayback(dom, session, desde=desde, hasta=hasta,
                                             maximo=args.max_articulos, pausa=args.espera,
                                             secciones=args.secciones):
                clave = _norm_url_clave(art["link"])
                if clave not in vistos:
                    vistos.add(clave)
                    trabajos.append(art)
            if len(trabajos) >= args.max_articulos:
                break
        if not trabajos:
            print("\nNo se reunieron articulos historicos. Revisa el dominio o el rango de anios.")
            return

    # 2) Modo sitemap: cobertura amplia (reciente) con filtro por fecha.
    elif args.sitemap:
        rango = ""
        if desde or hasta:
            rango = f" (rango {desde or '...'} a {hasta or '...'})"
        print(f"\nModo sitemap: reuniendo articulos de {len(args.sitemap)} sitemap(s){rango}...\n")
        for sm in args.sitemap:
            print(f"[sitemap] {sm}")
            for art in descubrir_por_sitemap(sm, session, desde=desde, hasta=hasta,
                                             maximo=args.max_articulos):
                if art not in vistos:
                    vistos.add(art)
                    trabajos.append({"link": art})
            if len(trabajos) >= args.max_articulos:
                break
        if not trabajos:
            print("\nNo se reunieron articulos del sitemap. Revisa la URL o el rango de fechas.")
            return

    # 3) Modo crawl: expandimos cada portada/seccion a sus articulos.
    elif args.crawl:
        print(f"\nModo crawl: descubriendo articulos en {len(urls)} portada(s)/seccion(es)...\n")
        for indice in urls:
            print(f"[indice] {indice}")
            for art in descubrir_articulos(indice, session, maximo=args.max_articulos):
                if art not in vistos:
                    vistos.add(art)
                    trabajos.append({"link": art})
            time.sleep(args.espera)
        if not trabajos:
            print("\nNo se descubrieron articulos. Revisa la URL o proba sin --crawl.")
            return

    # 4) Modo directo: las URLs son las paginas a analizar.
    else:
        trabajos = [{"link": u} for u in urls]

    print(f"\nBuscando la palabra '{palabra}' en {len(trabajos)} pagina(s)...\n")

    todas = []
    for i, t in enumerate(trabajos, 1):
        url = t["link"]
        print(f"[{i}/{len(trabajos)}] {url}")
        restante = (args.limite - len(todas)) if args.limite else 0
        todas.extend(procesar_pagina(
            url, palabra, session, limite=restante,
            url_descarga=t.get("descarga"), fecha_fallback=t.get("fecha"),
        ))

        if args.limite and len(todas) >= args.limite:
            todas = todas[:args.limite]
            print(f"\nLimite de {args.limite} alcanzado. Detengo el proceso.")
            break

        # Guardado incremental: util en corridas largas (wayback) por si se corta.
        if args.guardar_cada and i % args.guardar_cada == 0 and todas:
            guardar_excel(todas, args.salida)

        if i < len(trabajos):
            time.sleep(args.espera)

    if not todas:
        print("\nNo se encontraron apariciones. No se genero el Excel.")
        return

    guardar_excel(todas, args.salida)


if __name__ == "__main__":
    main()
