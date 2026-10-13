// Valuta le proposte di nomi sul banco di prova: Maschera da sola, rizzo-pii nei tre modi, e le unioni.
// Ogni metodo produce un insieme di valori proposti; si accettano tutti e si applicano come fa Maschera: la cella
// uguale a un valore si sostituisce tutta, e i valori di almeno 4 caratteri si cercano anche dentro i testi, con la
// stessa ricerca di Maschera (AnonCore.ricerca). Poi si guarda, per ogni occorrenza di un nome vero, se e' coperta.
// Uso: node valuta.js <cartella con banco.xlsx, celle.json, rizzo.json>
"use strict";
const fs = require("fs");
const path = require("path");
const XLSX = require("../../vendor/xlsx.full.min.js");
require("../../src/anonimizza_core.js");
const A = globalThis.AnonCore;
A.init(XLSX);

const dir = process.argv[2];
const { celle, nomi } = JSON.parse(fs.readFileSync(path.join(dir, "celle.json"), "utf8"));
const rizzo = JSON.parse(fs.readFileSync(path.join(dir, "rizzo.json"), "utf8"));
const MIN_OVUNQUE = 4;   // come MIN_LUNGHEZZA_OVUNQUE in anonimizza_core.js

// --- la verita': occorrenze di ogni forma, tale e quale; si tolgono quelle contenute in una forma piu' lunga
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const occorrenze = [];
celle.forEach((c, ic) => {
  const qui = [];
  Object.entries(nomi).forEach(([gruppo, forme]) => forme.forEach((f) => {
    const re = new RegExp("(?<![\\p{L}\\p{N}])" + esc(f) + "(?![\\p{L}\\p{N}])", "gu");
    let m;
    while ((m = re.exec(c.testo))) qui.push({ ic, gruppo, forma: f, a: m.index, b: m.index + f.length });
  }));
  qui.filter((o) => !qui.some((p) => p !== o && p.a <= o.a && p.b >= o.b && p.b - p.a > o.b - o.a))
    .forEach((o) => { o.intera = c.testo.trim() === o.forma; occorrenze.push(o); });
});

// --- dove un insieme di valori accettati copre i testi: [ic] -> array di [a, b)
function copertura(valori) {
  const perCella = celle.map(() => []);
  const chiavi = new Set(valori.map((v) => A.norm(v)));
  // come anonimizza_core.js: nel testo libero solo valori di almeno 4 caratteri e non fatti di sole cifre
  const lunghi = valori.filter((v) => v.trim().length >= MIN_OVUNQUE && !/^[\d\s.,/-]+$/.test(v.trim()));
  const ric = lunghi.length ? A.ricerca(lunghi.map((t) => ({ testo: t, codice: "X" }))) : null;
  celle.forEach((c, ic) => {
    if (chiavi.has(A.norm(c.testo))) { perCella[ic].push([0, c.testo.length]); return; }
    if (!ric) return;
    ric.re.lastIndex = 0;
    let m;
    // i confini di Maschera sono lookaround: la corrispondenza e' esattamente il valore
    while ((m = ric.re.exec(c.testo))) {
      perCella[ic].push([m.index, m.index + m[0].length]);
      if (m[0].length === 0) ric.re.lastIndex++;
    }
  });
  return perCella;
}

function stato(o, cop) {
  const r = cop[o.ic];
  let coperti = 0;
  for (let i = o.a; i < o.b; i++) if (r.some(([a, b]) => i >= a && i < b)) coperti++;
  const lettere = o.b - o.a;
  return coperti === lettere ? "coperta" : coperti > 0 ? "parziale" : "scoperta";
}

function valuta(nome, valori) {
  valori = [...new Map(valori.map((v) => [A.norm(v), v.trim()])).values()].filter(Boolean);
  const cop = copertura(valori);
  const perGruppo = {};
  const esito = { coperta: 0, parziale: 0, scoperta: 0 };
  const inFrase = { coperta: 0, parziale: 0, scoperta: 0 };
  occorrenze.forEach((o) => {
    const s = stato(o, cop);
    esito[s]++;
    if (!o.intera) inFrase[s]++;
    const g = (perGruppo[o.gruppo] = perGruppo[o.gruppo] || { coperta: 0, parziale: 0, scoperta: 0 });
    g[s]++;
  });
  // falsi suggerimenti: valori che, applicati da soli, non toccano nessun nome vero
  const falsi = valori.filter((v) => {
    const c1 = copertura([v]);
    return !occorrenze.some((o) => c1[o.ic].some(([a, b]) => a < o.b && b > o.a));
  });
  const celleSenzaNomi = celle.filter((c, ic) => cop[ic].length && !occorrenze.some((o) => o.ic === ic)).length;
  return { metodo: nome, proposti: valori.length, esito, inFrase, perGruppo, falsi, celleSenzaNomiToccate: celleSenzaNomi };
}

// --- i metodi
const lettura = A.leggi(fs.readFileSync(path.join(dir, "banco.xlsx")));
const maschera = A.testiDistinti(lettura).filter((t) => t.proposto).map((t) => t.testo);
const daRizzo = (v) => rizzo.varianti[v].celle.flatMap((c) => c.entita.map((e) => e.testo));
const metodi = [
  ["Maschera", maschera],
  ["rizzo-pii, cella", daRizzo("cella")],
  ["rizzo-pii, intestazione", daRizzo("intestazione")],
  ["rizzo-pii, riga", daRizzo("riga")],
  ["Maschera + rizzo-pii intestazione", maschera.concat(daRizzo("intestazione"))],
  ["Maschera + rizzo-pii cella", maschera.concat(daRizzo("cella"))],
];
// ESPLORATIVO, aggiunto dopo aver visto i risultati (non e' nella predizione): le proposte di Maschera fatte di una
// sola parola restano solo se rizzo-pii (intestazione) le riconosce anche lui; le altre restano tutte.
const confermati = new Set(daRizzo("intestazione").concat(daRizzo("cella")).map((t) => A.norm(t)));
const filtrata = maschera.filter((t) => /\s/.test(t.trim()) || confermati.has(A.norm(t)));
metodi.push(["(esplorativo) Maschera filtrata + rizzo-pii intestazione", filtrata.concat(daRizzo("intestazione"))]);
const risultati = metodi.map(([n, v]) => valuta(n, v));

const tot = occorrenze.length;
const pct = (n, d) => (d ? Math.round((100 * n) / d) : 0) + "%";
console.log("occorrenze di nomi: " + tot + " (" + occorrenze.filter((o) => !o.intera).length + " dentro una frase), celle di testo: " + celle.length);
console.log("modello: " + JSON.stringify(rizzo.modello));
Object.entries(rizzo.varianti).forEach(([v, x]) => console.log("tempo " + v + ": " + x.chiamate + " chiamate, " + x.secondi_totali + " s, massimo " + x.secondi_max + " s"));
console.log("");
console.log("| metodo | proposti | coperte | parziali | scoperte | in frase coperte | falsi suggerimenti | celle senza nomi toccate |");
console.log("|---|---|---|---|---|---|---|---|");
risultati.forEach((r) => {
  const f = occorrenze.filter((o) => !o.intera).length;
  console.log("| " + [r.metodo, r.proposti, r.esito.coperta + " (" + pct(r.esito.coperta, tot) + ")", r.esito.parziale, r.esito.scoperta,
    r.inFrase.coperta + "/" + f, r.falsi.length, r.celleSenzaNomiToccate].join(" | ") + " |");
});
console.log("");
const gruppi = Object.keys(nomi);
console.log("| gruppo (occorrenze) | " + risultati.map((r) => r.metodo).join(" | ") + " |");
console.log("|---|" + risultati.map(() => "---").join("|") + "|");
gruppi.forEach((g) => {
  const n = occorrenze.filter((o) => o.gruppo === g).length;
  console.log("| " + g + " (" + n + ") | " + risultati.map((r) => {
    const x = r.perGruppo[g] || { coperta: 0, parziale: 0 };
    return x.coperta + (x.parziale ? " +" + x.parziale + "p" : "");
  }).join(" | ") + " |");
});
console.log("");
risultati.forEach((r) => console.log("falsi suggerimenti, " + r.metodo + ": " + (r.falsi.join(" · ") || "nessuno")));
fs.writeFileSync(path.join(dir, "valutazione.json"), JSON.stringify({ occorrenze, risultati }, null, 1));
