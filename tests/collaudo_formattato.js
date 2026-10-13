// Collaudo dell'uscita che tiene la formattazione, su un file creato da Excel con dati INVENTATI
// (genera_prova_excel.ps1). Scrive prova_formattata_anonima.xlsx per il controllo in Excel (apri_in_excel.ps1).
// Uso: node collaudo_formattato.js <cartella>
"use strict";
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
require("../src/anonimizza_core.js");
const A = globalThis.AnonCore;
A.init(XLSX);

const dir = process.argv[2];
const S = JSON.parse(fs.readFileSync(path.join(dir, "sensibili.json"), "utf8"));
let falliti = 0;
function controlla(nome, ok, dettaglio) {
  console.log((ok ? "PASS " : "FAIL ") + nome + (dettaglio ? " — " + dettaglio : ""));
  if (!ok) falliti++;
}
const dec = new TextDecoder("utf-8");
function parti(b) {
  const zip = XLSX.CFB.read(b, { type: "array" }), out = {};
  zip.FullPaths.forEach((p, i) => {
    const f = zip.FileIndex[i];
    if (f.type === 2 && f.content && !/Sh33tJ5/.test(p)) out[p.replace(/^Root Entry\//, "")] = f.content;
  });
  return out;
}

const buf = new Uint8Array(fs.readFileSync(path.join(dir, "prova_formattata.xlsx")));
const lettura = A.leggi(buf);
const proposti = A.testiDistinti(lettura).filter((t) => t.proposto).map((t) => t.testo);
// come farebbe l'utente: le proposte e gli autori dei commenti (la pagina li marca da sola), piu' la forma breve e
// il nome dentro una frase segnalati dall'avviso, aggiunti a mano
const valori = proposti.concat(lettura.autori).map((t) => ({ testo: t, categoria: "PERSONA" }))
  .concat([{ testo: "Edilnord", categoria: "AZIENDA" }, { testo: "Laura Verdi", categoria: "PERSONA" }]);
const cfg = { fogli: {}, termini: "", valori: valori, ovunque: true, rinominaFogli: false,
  regole: { iban: true, cf: true, email: true, piva: true, tel: true } };
const esito = A.anonimizza(lettura, cfg, null);

// testi fuori dalle celle: la pagina li mostra prima di scaricare
const dove = [...new Set(lettura.extra.map((x) => x.dove.replace(/ .*/, "")))].sort().join(",");
controlla("testi fuori dalle celle trovati: commento, intestazione, casella di testo, grafico, convalida",
  ["casella", "commento", "grafico", "intestazione", "messaggio"].every((d) => dove.includes(d)), dove);
controlla("autore del commento proposto", lettura.autori.includes("Sara Moretti"), lettura.autori.join(","));
const casella = esito.fuori.find((x) => /casella/.test(x.dove));
controlla("casella di testo: nome sostituito anche nell'anteprima", !!casella && !/Davide Conti/.test(casella.dopo), casella && casella.dopo);

const prima = JSON.stringify(esito.report.regole);
const res = A.anonimizzaXlsx(buf, esito);
controlla("uscita formattata prodotta", !!res.buffer, res.motivo || "");
controlla("contatori dell'anteprima non gonfiati", JSON.stringify(esito.report.regole) === prima);
controlla("controllo finale dello strumento: pulito", res.controllo.ok, JSON.stringify(res.controllo.problemi));

// controllo indipendente e grezzo: ogni valore inventato, in ogni parte di testo del pacchetto (con le & delle
// intestazioni scritte doppie e gli accenti composti)
const sensibili = S.sensibili.concat(["Marco Galli", "Laura Verdi", "Edilnord", "Ferrero", "Tecnofarm", "Sara Moretti",
  "Davide Conti", "Elena Rinaldi", S.piva_numero.replace(/^0+/, "")]);
const pOut = parti(res.buffer), pIn = parti(buf);
const trovati = [];
Object.keys(pOut).forEach((p) => {
  if (!/\.(xml|rels|vml)$/.test(p)) return;
  const t = dec.decode(pOut[p]).replace(/&amp;/g, "&").replace(/&&/g, "&").normalize("NFC").toLowerCase();
  sensibili.forEach((v) => { if (t.includes(v.normalize("NFC").toLowerCase())) trovati.push(p + ": " + v); });
});
controlla("controllo indipendente: nessun valore inventato in nessuna parte", trovati.length === 0, trovati.join(" | "));

// proprieta' del file e percorso della cartella: c'erano nell'originale, nell'uscita no
const core = (p) => dec.decode(p["docProps/core.xml"]);
const autore = (s, tag) => (new RegExp("<" + tag + "[^>]*>([^<]*)<").exec(s) || [])[1] || "";
controlla("autore e ultimo a modificare: presenti nell'originale, vuoti nell'uscita",
  autore(core(pIn), "dc:creator") && !autore(core(pOut), "dc:creator") && !autore(core(pOut), "cp:lastModifiedBy"),
  "originale: " + (autore(core(pIn), "dc:creator") ? "pieno" : "vuoto"));
const wbIn = dec.decode(pIn["xl/workbook.xml"]), wbOut = dec.decode(pOut["xl/workbook.xml"]);
controlla("percorso della cartella di lavoro: presente nell'originale, tolto nell'uscita", /absPath/.test(wbIn) && !/absPath/.test(wbOut));
const conUtente = Object.keys(pOut).filter((p) => /\.(xml|rels|vml)$/.test(p) && /[A-Za-z]:\\+(Users|Utenti)\\+(?!utente\\)/i.test(dec.decode(pOut[p])));
controlla("nessun percorso con il nome dell'utente", conUtente.length === 0, conUtente.join(","));

// il controllo finale deve mordere: sul file originale deve trovare i nomi
const morso = A.controllaPacchetto(buf, esito.mappa);
controlla("il controllo finale scatta sul file originale", !morso.ok && morso.problemi.length >= 5,
  morso.problemi.map((p) => p.parte + " x" + p.volte).join(", "));

// parti del pacchetto: stesse dell'originale, a parte le stampanti (tolte di proposito)
const elenco = (p) => Object.keys(p).filter((x) => !/printerSettings/.test(x)).sort().join("|");
controlla("stesse parti del pacchetto dell'originale (stampanti a parte)", elenco(pIn) === elenco(pOut),
  Object.keys(pIn).filter((x) => !(x in pOut)).join(","));

// numeri, date e formati: ogni cella numerica dell'originale e' uguale nell'uscita, salvo le P.IVA salvate come
// numero sotto un'etichetta «P.IVA», che diventano codici. Un formato con un nome dentro («"rif. Davide Conti"»)
// cambia solo nel nome: rimettendo i nomi veri torna quello originale.
const righe = esito.mappa.righe();
const stessoFormato = (zo, zu) => (zo || "") === (zu || "") || (zo !== zu && A.ripristinaTesto(zu || "", righe).testo === zo);
const o = XLSX.read(buf, { type: "array", cellNF: true, cellFormula: true });
const u = XLSX.read(res.buffer, { type: "array", cellNF: true, cellFormula: true });
let numeri = 0;
const diversi = [], piva = [];
o.SheetNames.forEach((nome, i) => {
  const a = o.Sheets[nome], b = u.Sheets[u.SheetNames[i]];
  Object.keys(a).forEach((k) => {
    if (k[0] === "!" || a[k].t !== "n") return;
    numeri++;
    const x = b[k];
    if (S.pive.concat([S.piva_numero]).some((p) => +p === a[k].v)) { piva.push(nome + "!" + k + "=" + (x && x.v)); return; }
    if (!x || x.t !== "n" || x.v !== a[k].v || !stessoFormato(a[k].z, x.z) || !!x.f !== !!a[k].f)
      diversi.push(nome + "!" + k + " " + a[k].v + "/" + (x && x.v) + " " + (a[k].z || "") + "/" + (x && x.z || ""));
  });
});
controlla("numeri, date e formati numerici identici (" + numeri + " celle)", numeri > 10 && diversi.length === 0, diversi.slice(0, 5).join(" | "));
controlla("P.IVA salvate come numero diventate codici", piva.length === 2 && piva.every((x) => /PIVA_\d{3}$/.test(x)), piva.join(" | "));
const imp = u.Sheets[u.SheetNames[1]]["B5"];
controlla("importo che sembra una P.IVA resta numero", imp && imp.t === "n" && imp.v === S.importo_come_piva);

// ripristino che tiene la formattazione: il file anonimo torna con i nomi veri, fogli, formule e P.IVA numeriche
const rip = A.ripristinaXlsx(res.buffer, esito.mappa.righe());
const b = XLSX.read(rip.buffer, { type: "array" }), oo = XLSX.read(buf, { type: "array" });
const celle = ["A1", "A4", "B5", "C6", "D7", "A9"];
const uguali = celle.every((k) => String(oo.Sheets["Clienti"][k].v) === String(b.Sheets["Clienti"][k].v));
controlla("ripristino formattato: valori, nome del foglio e formule tornano quelli veri",
  rip.formattazione && uguali && b.SheetNames[1] === "Note Edilnord" && rip.sconosciuti.length === 0 &&
  b.Sheets["Riepilogo"]["A1"].f === "'Note Edilnord'!A1",
  b.SheetNames.join(",") + " | " + (b.Sheets["Riepilogo"]["A1"].f || ""));
const b4 = b.Sheets["Note Edilnord"]["B4"];
controlla("ripristino formattato: P.IVA salvata solo come numero torna numero", b4 && b4.t === "n" && b4.v === +S.piva_numero, b4 && b4.t + " " + b4.v);
fs.writeFileSync(path.join(dir, "prova_formattata_ripristinata.xlsx"), rip.buffer);

fs.writeFileSync(path.join(dir, "prova_formattata_anonima.xlsx"), res.buffer);
fs.writeFileSync(path.join(dir, "prova_formattata_corrispondenze.json"), JSON.stringify(esito.mappa.righe()));
console.log("nomi dei fogli:", esito.nomiFogli.join(", "), "| avvisi:", res.avvisi.join(" ") || "nessuno");
console.log("\n" + (falliti ? falliti + " controlli FALLITI" : "tutti i controlli superati"));
process.exit(falliti ? 1 : 0);
