"""Manda le celle del banco a rizzo-pii in locale e registra le entita' trovate, in tre modi:
- cella:        il testo della cella da solo;
- intestazione: «Intestazione: testo» (il modo naturale per Maschera, che conosce le intestazioni);
- riga:         tutta la riga, «Intestazione: valore; ...», come nella prova del 12/10/2026.
Per ogni entita' tiene solo la parte che cade dentro il testo della cella.
Uso: python rizzo.py <cartella con celle.json> [http://127.0.0.1:5005]
"""
import json
import sys
import time
import urllib.request
from pathlib import Path

ETICHETTE = {"FULLNAME", "ORG"}


def chiama(url, testo):
    dati = json.dumps({"text": testo, "include_mapping": True}).encode("utf-8")
    req = urllib.request.Request(url + "/analyze", data=dati, headers={"Content-Type": "application/json"})
    t0 = time.perf_counter()
    with urllib.request.urlopen(req, timeout=300) as r:
        out = json.load(r)
    return out, time.perf_counter() - t0


def entita(out):
    """Segmenti -> [{label, start, end, testo}] con le posizioni nel testo inviato."""
    pos, ris = 0, []
    for s in out["segments"]:
        t = s.get("t", "")
        if "label" in s:
            ris.append({"label": s["label"], "start": pos, "end": pos + len(t), "testo": t})
        pos += len(t)
    return ris


def dentro(ents, inizio, testo):
    """Le parti delle entita' che cadono nel testo della cella, che nel testo inviato comincia a «inizio»."""
    fine, ris = inizio + len(testo), []
    for e in ents:
        a, b = max(e["start"], inizio), min(e["end"], fine)
        if a < b:
            ris.append({"label": e["label"], "testo": testo[a - inizio:b - inizio].strip(), "intera": e["end"] - e["start"] == b - a})
    return ris


def main(cartella, url):
    cartella = Path(cartella)
    with open(cartella / "celle.json", encoding="utf-8") as fh:
        celle = json.load(fh)["celle"]
    with urllib.request.urlopen(url + "/health", timeout=10) as r:
        salute = json.load(r)
    risultati = {"modello": {k: salute.get(k) for k in ("app_version", "model", "model_version", "device")}, "varianti": {}}

    for variante in ("cella", "intestazione", "riga"):
        voci, tempi, cache = [], [], {}
        for c in celle:
            if variante == "intestazione" and c["intestazione"]:
                prefisso = c["intestazione"] + ": "
                inviato, inizio = prefisso + c["testo"], len(prefisso)
            elif variante == "riga" and c["testo_riga"]:
                inviato = c["testo_riga"]
                inizio = inviato.index("%s: %s" % (c["intestazione"], c["testo"])) + len(c["intestazione"]) + 2
            else:
                inviato, inizio = c["testo"], 0
            if inviato not in cache:
                out, sec = chiama(url, inviato)
                cache[inviato] = entita(out)
                tempi.append(sec)
            ents = [e for e in dentro(cache[inviato], inizio, c["testo"]) if e["label"] in ETICHETTE and e["testo"]]
            voci.append({"foglio": c["foglio"], "cella": c["cella"], "entita": ents})
        risultati["varianti"][variante] = {"celle": voci, "chiamate": len(tempi), "secondi_totali": round(sum(tempi), 2),
                                           "secondi_max": round(max(tempi), 3)}
        print("%-12s %3d chiamate, %.1f s" % (variante, len(tempi), sum(tempi)))
    with open(cartella / "rizzo.json", "w", encoding="utf-8") as fh:
        json.dump(risultati, fh, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "http://127.0.0.1:5005")
