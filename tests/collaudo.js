// Collaudo dell'anonimizzatore su dati INVENTATI (genera_prova.py).
// Uso: node collaudo.js <cartella con prova_bilancio.xlsx e sensibili.json>
// Esce con codice 1 se un controllo fallisce.
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

// --- controlli dei codici
const T = A._test;
controlla("CF pubblico valido", T.cfValido("RSSMRA85T10A562S"));
controlla("CF con controllo sbagliato scartato", !T.cfValido("RSSMRA85T10A562T"));
S.cf.forEach((c) => controlla("CF inventato valido", T.cfValido(c)));
S.pive.forEach((p) => controlla("P.IVA inventata valida " + p.slice(-2), T.pivaValida(p)));
controlla("P.IVA sbagliata scartata", !T.pivaValida("12345678901"));
controlla("IBAN seguito da una parola", T.ibanPrefissoValido(S.iban + " RATA") === S.iban.length);
controlla("IBAN con spazi", T.ibanPrefissoValido("IT60 X054 2811 1010 0000 0123 456") === 33);
// falsi IBAN (prova sui file veri del 13/10): testo normale, paese sconosciuto, inizio di un codice piu' lungo
controlla("IBAN attaccato ad altre cifre: scartato", T.ibanPrefissoValido(S.iban + "789") === 0);
controlla("IBAN con un paese sconosciuto: scartato", T.ibanPrefissoValido("ZZ60X0542811101000000123456") === 0);
controlla("IBAN di lunghezza sbagliata per il paese: scartato", T.ibanPrefissoValido(S.iban.slice(0, 22)) === 0);
(function () {
  // con la regola vecchia (ogni lunghezza da 15 a 34, minuscole ammesse) circa una frase su cinque di questa forma
  // passava per un IBAN: ora nessuna
  let falsi = 0;
  for (let i = 0; i < 2000; i++) {
    const frase = "da " + (10 + (i % 90)) + " Ma Ro cl " + i + " ab Cd ef Gh ij kl mn";
    if (A.anonimizza({ fogli: [{ nome: "F", righe: [[frase]], origine: { r: 0, c: 0 }, unioni: [] }], extra: [] },
      { fogli: {}, termini: "", valori: [], ovunque: true, regole: { iban: true } }).report.regole.iban) falsi++;
  }
  controlla("2000 frasi normali: nessuna presa per IBAN", falsi === 0, falsi + " falsi");
})();

// --- intestazioni di stampa: i codici di Excel (&L, &12, &"Arial,Bold", &&) restano, i nomi dentro no
const TI = (s) => T.trasformaIntestazione(s, (x) => x.replace(/Mario Rossi/g, "PERSONA_001").replace(/Rossi & Figli/g, "AZIENDA_001"));
controlla("intestazione: nome dopo il codice di posizione", TI("&LMario Rossi&CPagina &P") === "&LPERSONA_001&CPagina &P", TI("&LMario Rossi&CPagina &P"));
controlla("intestazione: & scritta doppia", TI('&"Arial,Bold"&12Rossi && Figli') === '&"Arial,Bold"&12AZIENDA_001', TI('&"Arial,Bold"&12Rossi && Figli'));
controlla("intestazione: la & del nome ripristinato torna doppia",
  T.trasformaIntestazione("&RAZIENDA_001", (x) => x.replace("AZIENDA_001", "Rossi & Figli")) === "&RRossi && Figli");
controlla("intestazione in parole", T.testoIntestazione("&LMario Rossi&R&D").trim().startsWith("Mario Rossi"));
// --- percorsi con il nome dell'utente
controlla("percorso Windows: nome utente tolto", T.togliUtente("C:\\Users\\mrossi\\Documents\\") === "C:\\Users\\utente\\Documents\\",
  T.togliUtente("C:\\Users\\mrossi\\Documents\\"));
controlla("percorso OneDrive: nome tolto", T.togliUtente("https://ditta-my.sharepoint.com/personal/mario_rossi_ditta_it/Documents/x.xlsx")
  === "https://ditta-my.sharepoint.com/personal/utente/Documents/x.xlsx");
// --- date: niente spostamento al giorno prima per il fuso orario
controlla("data in testo senza fuso", T.valoreInTesto(new Date(2025, 0, 15)) === "2025-01-15", T.valoreInTesto(new Date(2025, 0, 15)));
// --- area dichiarata enorme con due celle: niente milione di righe
(function () {
  const wb = XLSX.utils.book_new(), ws = XLSX.utils.aoa_to_sheet([["Mario Rossi", 1]]);
  XLSX.utils.book_append_sheet(wb, ws, "Foglio1");
  // l'area si dichiara direttamente nell'XML, come fa Excel: SheetJS la scriverebbe riga per riga
  const zip = XLSX.CFB.read(new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" })), { type: "array" });
  const i = zip.FullPaths.findIndex((p) => /sheet1\.xml$/.test(p));
  const x = new TextDecoder().decode(zip.FileIndex[i].content).replace(/<dimension ref="[^"]*"/, '<dimension ref="A1:T1048576"');
  zip.FileIndex[i].content = new TextEncoder().encode(x);
  zip.FileIndex[i].size = zip.FileIndex[i].content.length;
  const grande = XLSX.CFB.write(zip, { fileType: "zip", type: "array" });
  const t0 = Date.now(), l = A.leggi(grande);
  controlla("area dichiarata enorme: lette solo le righe vere", l.fogli[0].righe.length === 1 && Date.now() - t0 < 3000,
    l.fogli[0].righe.length + " righe, " + (Date.now() - t0) + " ms");
})();

// --- configurazione come la sceglierebbe un utente
const CONFIG = {
  fogli: {
    Clienti: { rigaIntest: 3, colonne: [
      { indice: 0, prefisso: "CLIENTE" }, { indice: 1, prefisso: "REFERENTE" }, { indice: 2, prefisso: "EMAIL" },
      { indice: 3, prefisso: "TELEFONO" }, { indice: 4, prefisso: "PIVA" }] },
    Movimenti: { rigaIntest: 1, colonne: [{ indice: 4, prefisso: "CLIENTE" }] },
    "Note Edilnord": { rigaIntest: 1, colonne: [] },
  },
  termini: "Laura Verdi => PERSONA\nGianni Fabbri => PERSONA\nMario Bianchi => PERSONA\nCM-2291 => COMMESSA\nEdilnord => CLIENTE\n" +
    "Nicolò Ferrè => PERSONA",
  regole: { iban: true, cf: true, email: true, piva: true, tel: true },
  ovunque: true,
  rinominaFogli: false,
};

function tutteLeCelle(wb) {
  const out = [];
  wb.SheetNames.forEach((n) => {
    out.push({ dove: "nome foglio", v: n });
    const ws = wb.Sheets[n];
    Object.keys(ws).forEach((k) => {
      if (k[0] === "!") return;
      out.push({ dove: n + "!" + k, v: ws[k].v, c: ws[k] });
    });
  });
  return out;
}

function prova(nomeFile, righePrecedenti) {
  console.log("\n== " + nomeFile + (righePrecedenti ? " (con corrispondenze precedenti)" : ""));
  const buf = fs.readFileSync(path.join(dir, nomeFile));
  const lettura = A.leggi(buf);
  const cols = A.colonne(lettura.fogli[0], 3);
  controlla("colonne suggerite: ragione sociale, referente, email, telefono, P.IVA",
    ["A", "B", "C", "D", "E"].every((l) => cols.find((c) => c.lettera === l).suggerita) &&
    !cols.find((c) => c.lettera === "F").suggerita, cols.filter((c) => c.suggerita).map((c) => c.lettera).join(","));
  const res = A.anonimizza(lettura, CONFIG, righePrecedenti);
  const outBuf = A.scrivi(res.wb);
  const out = XLSX.read(outBuf, { type: "array", cellDates: true });

  // 1. nessun valore sensibile in chiaro, in nessuna cella, nei nomi dei fogli, nei commenti, nelle proprieta'
  const celle = tutteLeCelle(out);
  const residui = [];
  celle.forEach((x) => {
    const s = String(x.v).normalize("NFC").toLowerCase();   // NFC: anche gli accenti scritti separati
    S.sensibili.forEach((v) => { if (s.includes(v.normalize("NFC").toLowerCase())) residui.push(x.dove + ": " + v); });
    if (x.c && (x.c.c || x.c.l)) residui.push(x.dove + ": commento o collegamento rimasto");
  });
  const props = JSON.stringify(out.Props || {});
  if (/Verdi/i.test(props)) residui.push("proprieta' del file: autore");
  controlla("nessun dato sensibile in chiaro", residui.length === 0, residui.join(" | "));

  // 2. i numeri non cambiano
  const wsC = out.Sheets["Clienti"], wsM = out.Sheets["Movimenti"];
  const fatt = S.numeri["Clienti!F"].every((n, i) => wsC["F" + (4 + i)].v === n);
  const imp = S.numeri["Movimenti!C"].every((n, i) => wsM["C" + (2 + i)].v === n);
  controlla("numeri intatti (fatturato e importi, anche quello che sembra una P.IVA)", fatt && imp &&
    S.numeri["Movimenti!C"].includes(S.importo_come_piva));
  controlla("date intatte", wsM["A2"].v instanceof Date && wsM["A2"].v.toISOString().slice(0, 10) === "2025-01-15");
  const desc7 = wsM["B7"].v;
  controlla("niente falsi positivi su date e numeri di fattura nel testo",
    desc7.includes("01/04/2025") && desc7.includes("2025/0012"), desc7);
  // openpyxl non calcola le formule: qui G2 non ha un valore salvato. Si controlla solo che nessuna formula resti;
  // la formula con valore salvato (scritta da Excel) diventa valore nel collaudo formattato.
  controlla("nessuna formula nel file a soli valori", !wsM["G2"] || !wsM["G2"].f);
  controlla("foglio nascosto resta nascosto", out.Workbook.Sheets[2].Hidden === 1);

  // 3. stesso valore, stesso codice, in tutto il file
  const codEdil = wsC["A4"].v;
  controlla("stesso cliente stesso codice tra colonna e controparte", wsM["E2"].v === codEdil, codEdil + " / " + wsM["E2"].v);
  controlla("cliente nel testo libero sostituito con lo stesso codice", wsM["B2"].v.includes(codEdil), wsM["B2"].v);

  // 4. ripristino
  const corrBuf = A.fileCorrispondenze(res.mappa);
  const righe = A.leggiCorrispondenze(corrBuf);
  const rip = A.ripristinaExcel(outBuf, righe);
  const back = XLSX.read(rip.buffer, { type: "array" });
  const orig = XLSX.read(buf, { type: "array" });
  const ok = ["A4", "B5", "C6", "D7", "A9"].every((k) => String(back.Sheets["Clienti"][k].v) === String(orig.Sheets["Clienti"][k].v));
  controlla("ripristino dell'Excel", ok && rip.sconosciuti.length === 0, "sconosciuti: " + rip.sconosciuti.join(","));
  controlla("ripristino: nomi dei fogli", JSON.stringify(back.SheetNames) === JSON.stringify(orig.SheetNames), back.SheetNames.join(","));
  // B4 c'e' solo come numero; B3 invece ha lo stesso valore scritto come testo nella colonna P.IVA, quindi un solo
  // codice per due forme: torna testo, con le stesse cifre.
  const b4 = back.Sheets["Note Edilnord"] && back.Sheets["Note Edilnord"]["B4"];
  controlla("ripristino: P.IVA salvata solo come numero torna numero", !!b4 && b4.t === "n" && b4.v === orig.Sheets["Note Edilnord"]["B4"].v,
    b4 && (b4.t + " " + b4.v));
  const rt = A.ripristinaTesto("Il cliente " + codEdil + " ha il codice inesistente CLIENTE_999.", righe);
  controlla("ripristino del testo", rt.trovati === 1 && rt.testo.includes("Edilnord Srl") && rt.sconosciuti[0] === "CLIENTE_999", rt.testo);

  // 5. file modificato da Claude: righe al contrario, una colonna nuova, un foglio di sintesi
  const mod = XLSX.read(outBuf, { type: "array", cellDates: true });
  const aoa = XLSX.utils.sheet_to_json(mod.Sheets["Clienti"], { header: 1, defval: null });
  const corpo = aoa.slice(3).reverse().map((r) => r.concat(["Rischio basso per " + r[0]]));
  const nuovoWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(nuovoWb, XLSX.utils.aoa_to_sheet([aoa[2].concat(["Commento"])].concat(corpo)), "Analisi");
  XLSX.utils.book_append_sheet(nuovoWb, XLSX.utils.aoa_to_sheet([["Sintesi: " + codEdil + " vale più di tutti gli altri insieme."]]), "Sintesi");
  const ripMod = A.ripristinaExcel(XLSX.write(nuovoWb, { type: "array", bookType: "xlsx" }), righe);
  const bm = XLSX.read(ripMod.buffer, { type: "array" });
  const an = bm.Sheets["Analisi"];
  const ultimo = orig.Sheets["Clienti"]["A9"].v;
  controlla("file modificato: nomi rimessi anche con righe spostate, colonne e fogli nuovi",
    an["A2"].v === ultimo && an["H2"].v === "Rischio basso per " + ultimo &&
    bm.Sheets["Sintesi"]["A1"].v.startsWith("Sintesi: Edilnord Srl"), an["A2"].v + " | " + an["H2"].v);

  console.log("report: " + JSON.stringify(res.report));
  console.log("primi testi rimasti (" + res.resti.length + "): " + res.resti.slice(0, 8).map((r) => (r.nome ? "*" : "") + r.testo).join(" | "));
  return { righe, codEdil };
}

// --- proposte automatiche sul file a tabella: aziende e persone si', intestazioni e frasi no
(function () {
  const lettura = A.leggi(fs.readFileSync(path.join(dir, "prova_bilancio.xlsx")));
  const proposti = A.testiDistinti(lettura).filter((t) => t.proposto).map((t) => t.testo).sort();
  const attesi = ["Alfa Logistica Srl", "Bianchi & Figli Snc", "Davide Conti", "Edilnord Srl", "Elena Rinaldi",
    "Giovanna Ferrero", "Laura Verdi", "Marco Galli", "Mondo Verde Coop", "Paolo Bianchi", "Sara Moretti",
    "Studio Associato Rinaldi", "Tecnofarm S.p.A."];
  controlla("proposte sul bilancio: solo aziende e persone", JSON.stringify(proposti) === JSON.stringify(attesi),
    "in piu': " + proposti.filter((p) => !attesi.includes(p)).join(", ") + " | mancano: " + attesi.filter((a) => !proposti.includes(a)).join(", "));
  // con le sole proposte, la forma breve «Edilnord» (nome del foglio) e «Bianchi» (Mario Bianchi) devono emergere
  const cfg = { fogli: {}, termini: "", valori: proposti, prefissoValori: "NOME", ovunque: true, rinominaFogli: false,
    regole: { iban: true, cf: true, email: true, piva: true, tel: true } };
  const fr = A.anonimizza(lettura, cfg, null).frammenti;
  const parole = [...new Set(fr.map((f) => f.parola))].sort();
  controlla("frammenti: forme brevi dei nomi marcati segnalate", parole.includes("Edilnord") && parole.includes("Bianchi") &&
    fr.some((f) => f.dove === "nome del foglio"), fr.map((f) => f.parola + " in «" + f.testo + "»").join(" | "));
})();

const primo = prova("prova_bilancio.xlsx");
// .xls vecchio formato: lo stesso file convertito
const wbXls = XLSX.read(fs.readFileSync(path.join(dir, "prova_bilancio.xlsx")), { type: "buffer", cellDates: true, cellNF: true });
fs.writeFileSync(path.join(dir, "prova_bilancio.xls"), XLSX.write(wbXls, { type: "buffer", bookType: "biff8" }));
const secondo = prova("prova_bilancio.xls", primo.righe);
controlla("secondo file con corrispondenze caricate: stessi codici", secondo.codEdil === primo.codEdil);

// --- prospetto a blocchi: nomi sparsi, nessuna colonna di nomi
(function () {
  console.log("\n== prova_prospetto.xlsx (nomi sparsi)");
  const P = JSON.parse(fs.readFileSync(path.join(dir, "prospetto.json"), "utf8"));
  const lettura = A.leggi(fs.readFileSync(path.join(dir, "prova_prospetto.xlsx")));
  const proposti = A.testiDistinti(lettura).filter((t) => t.proposto).map((t) => t.testo).sort();
  controlla("proposti esattamente i cognomi", JSON.stringify(proposti) === JSON.stringify(P.cognomi), proposti.join(", "));
  const cfg = { fogli: {}, termini: "", valori: proposti, prefissoValori: "NOME", ovunque: true, rinominaFogli: false,
    regole: { iban: true, cf: true, email: true, piva: true, tel: true } };
  const res = A.anonimizza(lettura, cfg, null);
  const out = XLSX.read(A.scrivi(res.wb), { type: "array" });
  const ws = out.Sheets["Consumi"];
  const testi = Object.keys(ws).filter((k) => k[0] !== "!").map((k) => String(ws[k].v));
  const rimasti = P.cognomi.filter((c) => testi.some((t) => t.toLowerCase().includes(c.toLowerCase())));
  controlla("nessun cognome rimasto", rimasti.length === 0, rimasti.join(", "));
  const etichetteOk = P.etichette.every((e) => testi.includes(e));
  controlla("etichette (mesi, voci, unita') intatte", etichetteOk);
  const numeriOk = P.numeri.every(([r, c, v]) => ws[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })].v === v);
  controlla("numeri e totali intatti", numeriOk);
  const nomi = testi.filter((t) => /^NOME_\d{3}$/.test(t));
  controlla("stesso cognome due volte, stesso codice", new Set(nomi).size === P.cognomi.length && nomi.length === 12,
    nomi.length + " celle, " + new Set(nomi).size + " codici");

  // categorie: due inquilini e il resto persone; poi il file del mese dopo con le corrispondenze caricate
  const marcati = proposti.map((t, i) => ({ testo: t, categoria: i < 2 ? "Inquilino" : "PERSONA" }));
  const res2 = A.anonimizza(lettura, Object.assign({}, cfg, { valori: marcati }), null);
  const righe2 = res2.mappa.righe();
  const cat = righe2.map((r) => r[0].replace(/_\d+$/, ""));
  controlla("codici con la categoria scelta", cat.filter((c) => c === "INQUILINO").length === 2 &&
    cat.filter((c) => c === "PERSONA").length === proposti.length - 2, [...new Set(cat)].join(","));
  const prop = A.proposteDaCorrispondenze(lettura, righe2);
  const ok = prop.length === proposti.length &&
    prop.every((p) => marcati.find((m) => m.testo === p.testo).categoria.toUpperCase() === p.categoria);
  controlla("corrispondenze del file precedente: nomi riproposti con la loro categoria", ok, prop.length + " proposti");
  const res3 = A.anonimizza(lettura, Object.assign({}, cfg, { valori: prop }), righe2);
  controlla("e con gli stessi codici", JSON.stringify(res3.mappa.righe()) === JSON.stringify(righe2) && res3.report.nuoviCodici === 0);
})();

console.log("\n" + (falliti ? falliti + " controlli FALLITI" : "tutti i controlli superati"));
process.exit(falliti ? 1 : 0);
