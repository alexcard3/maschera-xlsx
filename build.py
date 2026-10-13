"""Costruisce Maschera: un solo file HTML da aprire nel browser, in dist/Maschera.html.

Mette dentro la pagina la libreria SheetJS (sha256 controllato) e src/anonimizza_core.js, e aggiunge una
Content-Security-Policy che vieta le connessioni di rete fatte dal codice (fetch, XMLHttpRequest, WebSocket:
connect-src non e' concesso), il caricamento di script, stili, immagini e font da fuori, e l'invio di moduli. Non copre
la navigazione (un link cliccato, un window.open): la pagina ha un solo link, fisso, alla pagina del progetto nella nota
«versione in sviluppo», che si apre solo se l'utente lo clicca e non porta dati; non scrive mai HTML preso dal file
(solo testo), quindi un foglio di calcolo non puo' iniettarne.

Con --tema si applica un aspetto proprio (logo, colori, font, testi, nota «versione in sviluppo» si' o no): vedi
docs/tema-esempio.json.

Uso:  python build.py [--tema tema.json] [--uscita dist/Maschera.html]
"""
import argparse
import base64
import hashlib
import html
import json
import re
from pathlib import Path

QUI = Path(__file__).resolve().parent
SHA_SHEETJS = "cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41"   # xlsx 0.20.3, cdn.sheetjs.com
CSP = ('<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; script-src \'unsafe-inline\'; '
       'style-src \'unsafe-inline\'; img-src data: blob:; form-action \'none\'; base-uri \'none\'">')
FONT_BASE = '"Segoe UI", system-ui, -apple-system, sans-serif'
SEGNAPOSTO = ("<!--CSP-->", '<script src="../vendor/xlsx.full.min.js"></script>',
              '<script src="anonimizza_core.js"></script>', "/*tema:stile*/", "<!--tema:logo-->",
              "<!--tema:sovratitolo-->", "<!--versione-->")
RE_NOTA = re.compile(r"<!--tema:nota-->(.*?)<!--/tema:nota-->", re.S)
RE_AVVISO = re.compile(r"[ \t]*<!--tema:avviso-->\n(.*?)[ \t]*<!--/tema:avviso-->\n", re.S)
CHIAVI_TEMA = {"logo", "logo_alt", "sovratitolo", "nota", "font", "colori", "css", "avviso_sviluppo"}
TIPI_LOGO = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".svg": "image/svg+xml"}
RE_VARIABILE = re.compile(r"^--[a-z0-9-]+$")
RE_COLORE = re.compile(r"^#([0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$")
RE_FONT = re.compile(r"""^[\w\s"',.-]+$""")
# Avvisi di licenza dentro il file distribuito: nessun tema li puo' togliere (MIT e Apache 2.0, par. 4)
LICENZE = """<!--
  Maschera %s - https://github.com/alexcard3/maschera-xlsx
  Copyright (c) 2026 alexcard3 - MIT License: https://github.com/alexcard3/maschera-xlsx/blob/main/LICENSE
  Includes SheetJS Community Edition 0.20.3 - (C) 2013-present SheetJS - Apache License 2.0:
  https://www.apache.org/licenses/LICENSE-2.0
-->"""


def leggi(percorso):
    with open(percorso, encoding="utf-8", newline="") as fh:
        return fh.read()


def variabili(colori, quale):
    righe = []
    for nome, valore in (colori.get(quale) or {}).items():
        assert RE_VARIABILE.match(nome), "tema: nome di variabile non valido: " + nome
        assert isinstance(valore, str) and RE_COLORE.match(valore), "tema: colore non valido per %s: %r" % (nome, valore)
        righe.append("%s: %s;" % (nome, valore))
    return " ".join(righe)


def applica_tema(pagina, tema, cartella):
    """Sostituisce i segnaposto «tema:» con logo, sovratitolo, nota, colori, font e stile del tema (tutti facoltativi).
    I testi entrano come testo (html.escape), mai come HTML; colori e font si controllano prima di entrare nello stile."""
    sconosciute = set(tema) - CHIAVI_TEMA
    assert not sconosciute, "tema: chiavi sconosciute: " + ", ".join(sorted(sconosciute))
    for chiave in CHIAVI_TEMA - {"colori", "avviso_sviluppo"}:
        assert chiave not in tema or isinstance(tema[chiave], str), "tema: «%s» deve essere un testo" % chiave
    assert isinstance(tema.get("avviso_sviluppo", True), bool), "tema: «avviso_sviluppo» deve essere true o false"
    logo = ""
    if tema.get("logo"):
        file_logo = (cartella / tema["logo"]).resolve()
        tipo = TIPI_LOGO.get(file_logo.suffix.lower())
        assert tipo, "tema: il logo deve essere PNG, JPG o SVG"
        assert file_logo.exists(), "tema: logo non trovato: " + str(file_logo)
        dati = base64.b64encode(file_logo.read_bytes()).decode("ascii")
        logo = ('<img class="logo" src="data:%s;base64,%s" alt="%s"><span class="divisore"></span>'
                % (tipo, dati, html.escape(tema.get("logo_alt", ""))))
    sovra = '<div class="sovra">%s</div>' % html.escape(tema["sovratitolo"]) if tema.get("sovratitolo") else ""

    colori = tema.get("colori") or {}
    assert isinstance(colori, dict) and set(colori) <= {"chiaro", "scuro"}, "tema: «colori» ammette solo «chiaro» e «scuro»"
    assert all(isinstance(colori[k], dict) for k in colori), "tema: «chiaro» e «scuro» devono essere elenchi di colori"
    font = tema.get("font")
    assert not font or (RE_FONT.match(font) and "/*" not in font), "tema: font non valido (solo nomi, virgole e virgolette)"
    stile = []
    chiaro = (variabili(colori, "chiaro") + (" --font: %s, %s;" % (font, FONT_BASE) if font else "")).strip()
    if chiaro:
        stile.append(":root { %s }" % chiaro)
    scuro = variabili(colori, "scuro")
    if scuro:
        stile.append('@media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) { %s } }' % scuro)
        stile.append(':root[data-theme="dark"] { %s }' % scuro)
    if tema.get("css"):
        assert not re.search(r"</|<!--", tema["css"]), "tema: il css non puo' contenere «</» o «<!--»"
        stile.append(tema["css"])

    pagina = pagina.replace("<!--tema:logo-->", logo).replace("<!--tema:sovratitolo-->", sovra)
    pagina = pagina.replace("/*tema:stile*/", ("/* tema */\n" + "\n".join(stile)) if stile else "")
    if tema.get("nota"):
        pagina = RE_NOTA.sub(lambda m: html.escape(tema["nota"]), pagina)
    if tema.get("avviso_sviluppo") is False:
        pagina = RE_AVVISO.sub("", pagina)
    return pagina


def togli_segnaposto_tema(pagina):
    pagina = pagina.replace("<!--tema:logo-->", "").replace("<!--tema:sovratitolo-->", "").replace("/*tema:stile*/", "")
    pagina = RE_AVVISO.sub(lambda m: m.group(1), pagina)
    return RE_NOTA.sub(lambda m: m.group(1), pagina)


def inserisci_codice(pagina, sheetjs, core, versione):
    """Mette dentro la pagina la CSP, la libreria e la logica. Va fatto per ultimo: le sostituzioni del tema non devono
    poter toccare il codice delle librerie."""
    lib = re.sub(r"</(script)", r"<\\/\1", sheetjs.decode("utf-8"), flags=re.I)
    assert not re.search(r"</script", core, re.I), "anonimizza_core.js contiene </script"
    pagina = (pagina.replace("<!--CSP-->", CSP + "\n" + LICENZE % versione)
              .replace('<script src="../vendor/xlsx.full.min.js"></script>', "<script>\n" + lib + "\n</script>")
              .replace('<script src="anonimizza_core.js"></script>', "<script>\n" + core + "\n</script>"))
    assert len(re.findall(r"<script\b", pagina, re.I)) == 3 and len(re.findall(r"</script", pagina, re.I)) == 3, "script spezzati"
    return pagina


def costruisci(tema=None, uscita=None):
    sheetjs = (QUI / "vendor" / "xlsx.full.min.js").read_bytes()
    assert hashlib.sha256(sheetjs).hexdigest() == SHA_SHEETJS, "vendor/xlsx.full.min.js diverso da quello controllato"
    versione = leggi(QUI / "VERSION").strip()
    assert re.fullmatch(r"\d+\.\d+\.\d+(-[0-9A-Za-z.]+)?", versione), "VERSION non valido: " + versione
    pagina = leggi(QUI / "src" / "maschera.html")
    for s in SEGNAPOSTO:
        assert pagina.count(s) == 1, "segnaposto mancante o doppio: " + s
    assert len(RE_NOTA.findall(pagina)) == 1, "segnaposto mancante o doppio: tema:nota"
    assert len(RE_AVVISO.findall(pagina)) == 1, "segnaposto mancante o doppio: tema:avviso"

    if tema:
        percorso = Path(tema).resolve()
        with open(percorso, encoding="utf-8") as fh:
            pagina = applica_tema(pagina, json.load(fh), percorso.parent)
    pagina = togli_segnaposto_tema(pagina).replace("<!--versione-->", html.escape(versione))
    pagina = inserisci_codice(pagina, sheetjs, leggi(QUI / "src" / "anonimizza_core.js"), versione)

    uscita = Path(uscita) if uscita else QUI / "dist" / "Maschera.html"
    uscita.parent.mkdir(parents=True, exist_ok=True)
    dati = pagina.encode("utf-8")
    with open(uscita, "wb") as fh:
        fh.write(dati)
    print("%s  %d byte  sha256 %s" % (uscita, len(dati), hashlib.sha256(dati).hexdigest()))
    return uscita


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Costruisce Maschera in un solo file HTML / Builds Maschera into one HTML file.")
    ap.add_argument("--tema", help="file JSON con l'aspetto (vedi docs/tema-esempio.json)")
    ap.add_argument("--uscita", help="file da scrivere (predefinito: dist/Maschera.html)")
    a = ap.parse_args()
    costruisci(a.tema, a.uscita)
