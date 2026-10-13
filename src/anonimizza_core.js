/* anonimizza_core.js — logica di Maschera: anonimizza e ripristina i fogli di calcolo (prima versione 12/10/2026).
 *
 * Nessuna chiamata di rete: lavora su ArrayBuffer e oggetti in memoria.
 * Usa SheetJS (XLSX) passato da fuori: nel browser e' la variabile globale, in Node si passa con init().
 *
 * Idea: l'utente sceglie le colonne che contengono nomi (clienti, fornitori, persone...). Ogni valore di quelle
 * colonne diventa un codice stabile (CLIENTE_001), e lo stesso valore prende sempre lo stesso codice. Gli stessi
 * valori, piu' un elenco di termini scritto a mano, si cercano anche nel testo libero delle altre celle. Regole fisse
 * per CF, P.IVA, IBAN, email e telefono. I numeri non si toccano mai fuori dalle colonne scelte, salvo le P.IVA
 * salvate come numero sotto un'intestazione o accanto a un'etichetta «P.IVA».
 */
(function (root) {
  "use strict";
  var XLSX = root.XLSX;

  var CONFINE_PRIMA = "(?<![\\p{L}\\p{N}])";
  var CONFINE_DOPO = "(?![\\p{L}\\p{N}])";
  var MIN_LUNGHEZZA_OVUNQUE = 4;   // valori piu' corti non si cercano nel testo libero: troppi falsi positivi

  // ------------------------------------------------------------------ utilita'
  function norm(s) {
    return String(s).normalize("NFKC").trim().replace(/\s+/g, " ").toLowerCase();
  }

  function prefissoDa(testo, ripiego) {
    var p = String(testo || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
      .toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 20).replace(/_+$/, "");
    if (!p) p = ripiego;
    if (/^\d/.test(p)) p = "C_" + p;
    return p;
  }

  function lettera(c) { return XLSX.utils.encode_col(c); }

  function valoreInTesto(v) {
    // data locale: toISOString passa per UTC e in Italia sposterebbe le mezzanotti al giorno prima
    if (v instanceof Date) {
      if (isNaN(v)) return "";
      var p = function (x) { return String(x).padStart(2, "0"); };
      return v.getFullYear() + "-" + p(v.getMonth() + 1) + "-" + p(v.getDate());
    }
    return String(v);
  }

  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

  // Forme da cercare per un testo marcato. Il testo del file si porta in NFC prima della ricerca (cosi' «Nicolò» con
  // l'accento separato, comune nei testi incollati da un Mac, viene trovato). Il testo marcato si cerca come e'
  // scritto (in NFC, con le sue maiuscole: la «İ» turca non ha una minuscola semplice), in minuscolo, e nella forma
  // NFKC di norm() (che scioglie legature come «ﬀ» e trasforma «´» in accento): quale delle tre compare nel file
  // dipende da come e' stato scritto.
  function formeRicerca(testo) {
    var nfc = String(testo).normalize("NFC").trim().replace(/\s+/g, " "), forme = [];
    [nfc, nfc.toLowerCase(), norm(testo)].forEach(function (f) { if (f && forme.indexOf(f) < 0) forme.push(f); });
    return forme;
  }

  // Un nome di foglio si scrive senza apici nelle formule se fatto solo di lettere, cifre, _ e punto.
  var FOGLIO_SEMPLICE = /^[\p{L}_][\p{L}\p{N}_.]*$/u;
  function foglioCitato(nome) { return FOGLIO_SEMPLICE.test(nome) ? nome : "'" + nome.replace(/'/g, "''") + "'"; }

  // Parole che, in un'intestazione di colonna o in un'etichetta accanto, dicono che il numero e' una P.IVA.
  var ETICHETTA_PIVA = /p\.?\s?iva|partita\s+iva|\bvat\b|\bcf\b|c\.f\.|codice\s+fiscale/i;

  // ------------------------------------------------------------------ corrispondenze
  function Mappa() {
    this.perChiave = new Map();   // norm(valore) -> codice
    this.perCodice = new Map();   // codice -> {valore, origine}
    this.contatori = new Map();   // prefisso -> ultimo numero
    this.nuovi = 0;
  }
  // chiave: forma usata per riconoscere lo stesso valore scritto in modi diversi (es. telefono senza spazi);
  // nel file delle corrispondenze resta il valore com'era scritto la prima volta.
  Mappa.prototype.codice = function (valore, prefisso, origine, chiave) {
    var k = norm(chiave != null ? chiave : valore);
    if (this.perChiave.has(k)) return this.perChiave.get(k);
    var n = (this.contatori.get(prefisso) || 0) + 1;
    this.contatori.set(prefisso, n);
    var c = prefisso + "_" + String(n).padStart(3, "0");
    this.perChiave.set(k, c);
    this.perCodice.set(c, { valore: String(valore).trim(), origine: origine });
    this.nuovi++;
    return c;
  };
  Mappa.prototype.carica = function (righe) {
    // righe: [[codice, valore, origine], ...] senza intestazione
    var self = this;
    righe.forEach(function (r) {
      if (!r || r[0] == null || r[1] == null) return;
      var c = String(r[0]).trim(), m = /^(.*)_(\d+)$/.exec(c);
      if (!m) return;
      self.perChiave.set(norm(r[1]), c);
      if (/^(IBAN|PIVA|TEL)$/.test(m[1])) self.perChiave.set(norm(String(r[1]).replace(/[\s.\/-]/g, "")), c);
      self.perCodice.set(c, { valore: String(r[1]), origine: r[2] == null ? "" : String(r[2]) });
      self.contatori.set(m[1], Math.max(self.contatori.get(m[1]) || 0, +m[2]));
    });
  };
  Mappa.prototype.righe = function () {
    var out = [];
    this.perCodice.forEach(function (v, c) { out.push([c, v.valore, v.origine]); });
    out.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
    return out;
  };

  // ------------------------------------------------------------------ controlli dei codici
  function pivaValida(d) {
    if (!/^\d{11}$/.test(d) || /^0+$/.test(d)) return false;
    var s = 0;
    for (var i = 0; i < 10; i++) {
      var x = +d[i];
      if (i % 2 === 1) { x *= 2; if (x > 9) x -= 9; }
      s += x;
    }
    return (10 - (s % 10)) % 10 === +d[10];
  }

  // Da numero si perdono gli zeri iniziali: «01234567897» diventa 1234567897. Si accettano da 7 cifre in su (resta il
  // controllo della cifra finale, e serve comunque l'etichetta «P.IVA»).
  function pivaNumerica(v) {
    return typeof v === "number" && Number.isInteger(v) && v >= 1e6 && v < 1e11 && pivaValida(String(v).padStart(11, "0"));
  }

  var CF_DISPARI =[1, 0, 5, 7, 9, 13, 15, 17, 19, 21, 2, 4, 18, 20, 11, 3, 6, 8, 12, 14, 16, 10, 22, 25, 24, 23];
  function cfValido(cf) {
    cf = cf.toUpperCase();
    if (!/^[A-Z0-9]{16}$/.test(cf)) return false;
    var s = 0;
    for (var i = 0; i < 15; i++) {
      var ch = cf[i], v = /\d/.test(ch) ? +ch : ch.charCodeAt(0) - 65;
      s += (i % 2 === 0) ? CF_DISPARI[v] : v;
    }
    return String.fromCharCode(65 + (s % 26)) === cf[15];
  }

  // Lunghezza dell'IBAN per paese (registro SWIFT). Un IBAN si riconosce solo con un paese di questo elenco e la sua
  // lunghezza esatta: provando tutte le lunghezze, in un testo qualunque una sequenza passa il controllo per caso
  // (1 su 97 a tentativo), e parole normali diventavano «IBAN_001» (prova sui file veri, 13/10).
  var IBAN_LUNGHEZZE = { IT: 27, SM: 27, VA: 22, DE: 22, FR: 27, ES: 24, CH: 21, GB: 22, AT: 20, BE: 16, NL: 18,
    PT: 25, LU: 20, IE: 22, MC: 27, LI: 21, PL: 28, CZ: 24, SK: 24, SI: 19, HR: 21, RO: 24, GR: 27, DK: 18,
    AD: 24, AE: 23, AL: 28, BA: 20, BG: 22, CY: 28, EE: 20, FI: 18, FO: 18, GI: 23, GL: 18, HU: 28, IS: 26, LT: 20,
    LV: 21, ME: 22, MK: 19, MT: 31, NO: 15, RS: 22, SE: 24, TR: 26, UA: 29, SA: 24, IL: 23, MD: 24, GE: 22, AZ: 28,
    BH: 22, KW: 30, QA: 29, JO: 30, LB: 28, EG: 29, TN: 24, MA: 28, BR: 29, XK: 20 };
  function ibanMod97(s) {
    var r = s.slice(4) + s.slice(0, 4), resto = 0;
    for (var i = 0; i < r.length; i++) {
      var ch = r[i], n = /\d/.test(ch) ? ch : String(ch.charCodeAt(0) - 55);
      for (var j = 0; j < n.length; j++) resto = (resto * 10 + +n[j]) % 97;
    }
    return resto === 1;
  }
  // Ritorna quanti caratteri del testo trovato formano un IBAN valido (gli spazi contano), o 0. Paese conosciuto,
  // lunghezza esatta del paese, controllo mod-97, e subito dopo niente lettere o cifre attaccate (un IBAN non e'
  // l'inizio di un codice piu' lungo).
  function ibanPrefissoValido(testo) {
    var compatto = testo.replace(/ /g, "");
    var L = IBAN_LUNGHEZZE[compatto.slice(0, 2)];
    if (!L || compatto.length < L || !/^[A-Z]{2}\d{2}[A-Z0-9]+$/.test(compatto.slice(0, L)) || !ibanMod97(compatto.slice(0, L))) return 0;
    var visti = 0, i = 0;
    for (; i < testo.length && visti < L; i++) if (testo[i] !== " ") visti++;
    if (/[A-Za-z0-9]/.test(testo.charAt(i))) return 0;
    return i;
  }

  // Ogni regola: re (con flag g), prefisso, valida(testoTrovato) -> caratteri da sostituire (0 = scarta).
  var REGOLE = {
    // anche con uno spazio tra paese e cifre di controllo («IT 60X…»): l'IBAN trovato si confronta senza spazi
    iban: { prefisso: "IBAN", re: /(?<![A-Za-z0-9])[A-Z]{2} ?\d{2}(?: ?[A-Z0-9]){11,32}(?: ?[A-Za-z0-9])?/g,
      valida: function (t) { return ibanPrefissoValido(t); } },
    cf: { prefisso: "CF",
      re: /(?<![A-Za-z0-9])[A-Za-z]{6}[0-9LMNPQRSTUVlmnpqrstuv]{2}[A-Za-z][0-9LMNPQRSTUVlmnpqrstuv]{2}[A-Za-z][0-9LMNPQRSTUVlmnpqrstuv]{3}[A-Za-z](?![A-Za-z0-9])/g,
      valida: function (t) { return cfValido(t) ? t.length : 0; } },
    email: { prefisso: "EMAIL", re: /[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+/g,
      valida: function (t) { return t.length; } },
    piva: { prefisso: "PIVA", re: /(?<![A-Za-z0-9])(?:IT ?)?\d{11}(?!\d)/g,
      valida: function (t) { return pivaValida(t.replace(/^IT ?/, "")) ? t.length : 0; } },
    tel: { prefisso: "TEL",
      re: /(?<![\p{L}\p{N}+])(?:\+39[ .-]?)?(?:3\d{2}[ .-]?\d{3}[ .-]?\d{3,4}|0\d{1,3}[ ./-]?\d{5,8})(?![\p{L}\p{N}])/gu,
      valida: function (t) { var n = t.replace(/^\+39/, "").replace(/\D/g, "").length; return n >= 9 && n <= 11 ? t.length : 0; } }
  };
  var ORDINE_REGOLE = ["iban", "cf", "email", "piva", "tel"];

  // ------------------------------------------------------------------ sostituzione a segmenti
  // Un testo e' una lista di segmenti {t, fisso}. I segmenti fissi sono codici gia' messi: non si toccano piu'.
  function applicaRe(segmenti, re, scegli) {
    var out = [];
    segmenti.forEach(function (seg) {
      if (seg.fisso) { out.push(seg); return; }
      var s = seg.t, pos = 0, m;
      re.lastIndex = 0;
      while ((m = re.exec(s)) !== null) {
        var r = scegli(m[0]);
        if (!r) { re.lastIndex = m.index + 1; continue; }
        if (m.index > pos) out.push({ t: s.slice(pos, m.index) });
        out.push({ t: r.codice, fisso: true });
        pos = m.index + r.lunghezza;
        re.lastIndex = pos;
      }
      if (pos < s.length) out.push({ t: s.slice(pos) });
    });
    return out;
  }

  function costruisciRicerca(voci) {
    // voci: [{testo, codice}] -> {re, perChiave}
    var perChiave = new Map(), testi = [];
    voci.forEach(function (v) {
      var k = norm(v.testo);
      if (!k || perChiave.has(k)) return;
      perChiave.set(k, v.codice);
      testi.push(v.testo);
    });
    if (!testi.length) return null;
    var alternative = [];
    testi.forEach(function (t) { formeRicerca(t).forEach(function (f) { if (alternative.indexOf(f) < 0) alternative.push(f); }); });
    alternative.sort(function (a, b) { return b.length - a.length; });
    alternative = alternative.map(function (t) { return escapeRe(t).replace(/ /g, "\\s+"); });
    return { re: new RegExp(CONFINE_PRIMA + "(?:" + alternative.join("|") + ")" + CONFINE_DOPO, "giu"), perChiave: perChiave };
  }

  // ------------------------------------------------------------------ lettura
  function leggi(buffer) {
    var wb = XLSX.read(buffer, { type: "array", cellDates: true, cellFormula: true, cellStyles: false });
    var nascosti = ((wb.Workbook || {}).Sheets || []);
    var extra = [], autori = [];
    function aggiungiExtra(testo, dove) {
      testo = String(testo == null ? "" : testo).trim();
      if (testo && /\p{L}/u.test(testo)) extra.push({ testo: testo, dove: dove });
    }
    var fogli = wb.SheetNames.map(function (nome, i) {
      var ws = wb.Sheets[nome];
      var info = { nome: nome, nascosto: !!(nascosti[i] && nascosti[i].Hidden), origine: { r: 0, c: 0 },
        righe: [], commenti: 0, formule: 0, collegamenti: 0, unioni: ws["!merges"] || [], colonne: ws["!cols"] };
      // L'area dichiarata dal file («A1:T1048576») puo' essere enorme anche con due celle piene: si usa quella
      // delle celle che ci sono davvero, se no il browser costruisce un milione di righe vuote.
      var r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
      Object.keys(ws).forEach(function (k) {
        if (k[0] === "!") return;
        var c = ws[k], a = XLSX.utils.decode_cell(k);
        if (a.r < r0) r0 = a.r;
        if (a.c < c0) c0 = a.c;
        if (a.r > r1) r1 = a.r;
        if (a.c > c1) c1 = a.c;
        if (c.c && c.c.length) {
          info.commenti++;
          c.c.forEach(function (x) {
            if (x.a && autori.indexOf(String(x.a).trim()) < 0 && !/^tc=/.test(x.a)) autori.push(String(x.a).trim());
            aggiungiExtra(x.t, "commento in «" + nome + "» " + k);
          });
        }
        if (c.f) info.formule++;
        if (c.l) info.collegamenti++;
      });
      if (r1 < 0) return info;
      var range = { s: { r: r0, c: c0 }, e: { r: r1, c: c1 } };
      info.origine = range.s;
      info.righe = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true, range: range });
      return info;
    });
    if (eZip(buffer)) {
      var zip = XLSX.CFB.read(buffer, { type: "array" });
      testiDelPacchetto(zip).forEach(function (x) { aggiungiExtra(x.testo, x.dove); });
      utentiDelPacchetto(zip).forEach(function (u) { if (autori.indexOf(u) < 0) autori.push(u); });
    }
    ((wb.Workbook || {}).Names || []).forEach(function (n) {
      letterali(String(n.Ref || ""), function (t) { aggiungiExtra(t, "nome definito «" + n.Name + "»"); return t; });
    });
    return { fogli: fogli, nomiDefiniti: ((wb.Workbook || {}).Names || []).length, extra: extra, autori: autori };
  }

  // ------------------------------------------------------------------ parti del pacchetto .xlsx
  function percorsoParte(p) { return p.replace(/^Root Entry\//, ""); }
  function testoParte(zip, i) { return new TextDecoder("utf-8").decode(new Uint8Array(zip.FileIndex[i].content)); }
  function trovaParte(zip, nome) {
    for (var i = 0; i < zip.FullPaths.length; i++) if (percorsoParte(zip.FullPaths[i]) === nome && zip.FileIndex[i].content) return i;
    return -1;
  }

  // Nome del foglio per ogni file xl/worksheets/sheetN.xml (il numero del file non e' la posizione del foglio).
  function fogliPerParte(zip) {
    var out = {}, iWb = trovaParte(zip, "xl/workbook.xml"), iRel = trovaParte(zip, "xl/_rels/workbook.xml.rels");
    if (iWb < 0 || iRel < 0) return out;
    var dest = {};
    testoParte(zip, iRel).replace(/<Relationship\b[^>]*>/g, function (tag) {
      var id = /\bId="([^"]*)"/.exec(tag), t = /\bTarget="([^"]*)"/.exec(tag);
      if (id && t) dest[id[1]] = t[1].replace(/^\/?xl\//, "").replace(/^\//, "");
      return tag;
    });
    testoParte(zip, iWb).replace(/<sheet\b[^>]*>/g, function (tag) {
      var n = /\bname="([^"]*)"/.exec(tag), id = /\br:id="([^"]*)"/.exec(tag);
      if (n && id && dest[id[1]]) out["xl/" + dest[id[1]]] = xmlDec(n[1]);
      return tag;
    });
    return out;
  }

  // Nomi degli utenti di Office che Excel scrive da solo nel file: autore e ultimo a modificare, chi ha condiviso o
  // revisionato il file, chi ha aggiornato una pivot, chi ha creato uno scenario, le persone dei commenti moderni,
  // l'utente di SharePoint, le visualizzazioni personali («Mario Rossi - Visualizzazione personale»). Sono persone
  // (colleghi): la pagina li marca da sola, cosi' si sostituiscono ovunque e il controllo finale li cerca.
  function utentiDelPacchetto(zip) {
    var out = [];
    function aggiungi(v) {
      v = xmlDec(String(v || "")).trim();
      if (v.length >= 3 && !/^(tc=|\{|S-1-)/.test(v) && /\p{L}/u.test(v) && out.indexOf(v) < 0) out.push(v);
    }
    zip.FullPaths.forEach(function (pp, i) {
      var p = percorsoParte(pp), f = zip.FileIndex[i];
      if (f.type !== 2 || !f.content || !PARTI_TESTO.test(p)) return;
      var s = testoParte(zip, i);
      if (p === "docProps/core.xml") s.replace(/<(dc:creator|cp:lastModifiedBy)(?:\s[^>]*)?>([^<]*)</g, function (a, t, v) { aggiungi(v); return a; });
      s.replace(/\s(?:userName|refreshedBy|user)="([^"]*)"/g, function (a, v) { aggiungi(v); return a; });
      s.replace(/<userInfo\b[^>]*\sname="([^"]*)"/g, function (a, v) { aggiungi(v); return a; });
      s.replace(/<person\b[^>]*\sdisplayName="([^"]*)"/g, function (a, v) { aggiungi(v); return a; });
      s.replace(/<(?:\w+:)?DisplayName>([^<]*)</g, function (a, v) { aggiungi(v); return a; });
      s.replace(/<customWorkbookView\b[^>]*\sname="([^"]*)"/g, function (a, v) { aggiungi(xmlDec(v).replace(/\s+-\s+[^-]*$/, "")); return a; });
    });
    return out;
  }

  // Codici delle intestazioni e dei piè di pagina di stampa: &L &C &R (posizione), &P &N &D &T &F &A &Z &G (campi),
  // &B &I &U &E &S &X &Y &H (stile), &K... (colore), &"font" e &12 (dimensione). «&&» e' una & scritta.
  var CODICE_INTEST = /&(?:"[^"]*"|K(?:[0-9A-Fa-f]{6}|\d{2}[+-]\d{3})|\d+|[LCRPNDTFAZGBIUESXYH+\-])/g;
  // Applica T ai soli pezzi di testo, lasciando i codici dove sono.
  function trasformaIntestazione(s, T) {
    var out = "", pos = 0, m;
    function letterale(x) { return x ? T(x.replace(/&&/g, "&")).replace(/&/g, "&&") : ""; }
    var re = new RegExp("&&|" + CODICE_INTEST.source, "g");
    while ((m = re.exec(s)) !== null) {
      if (m[0] === "&&") continue;
      out += letterale(s.slice(pos, m.index)) + m[0];
      pos = m.index + m[0].length;
    }
    return out + letterale(s.slice(pos));
  }
  // I pezzi di testo di un'intestazione, gli stessi a cui trasformaIntestazione applica T (con «&&» gia' sciolta).
  function pezziIntestazione(s) {
    var pezzi = [];
    trasformaIntestazione(s, function (x) { if (x.trim()) pezzi.push(x.trim()); return x; });
    return pezzi;
  }
  // Solo il testo leggibile di un'intestazione (per le proposte e per il controllo finale).
  function testoIntestazione(s) { return s.replace(/&&/g, "\u0000").replace(CODICE_INTEST, " ").replace(/\u0000/g, "&"); }

  var TAG_INTEST = /<(oddHeader|oddFooter|evenHeader|evenFooter|firstHeader|firstFooter)>([\s\S]*?)<\/\1>/g;
  var ATTR_CONVALIDA = "prompt|promptTitle|error|errorTitle";

  // Testi che non stanno nelle celle: intestazioni di stampa, messaggi delle convalide, caselle di testo, titoli dei
  // grafici. Non si vedono nella griglia, ma finiscono nel file: l'utente li deve poter vedere e marcare.
  function testiDelPacchetto(zip) {
    var out = [], nomi = fogliPerParte(zip);
    zip.FullPaths.forEach(function (pp, i) {
      var p = percorsoParte(pp), f = zip.FileIndex[i];
      if (f.type !== 2 || !f.content) return;
      if (/^xl\/worksheets\/[^/]*\.xml$/.test(p)) {
        var nome = nomi[p] || p, s = testoParte(zip, i);
        // una voce per ogni pezzo di testo tra due codici: nel file si trasformano uno per uno
        s.replace(TAG_INTEST, function (a, tag, x) {
          pezziIntestazione(xmlDec(x)).forEach(function (pezzo) { out.push({ testo: pezzo, dove: "intestazione di stampa di «" + nome + "»" }); });
          return a;
        });
        s.replace(/<dataValidation\b[^>]*>/g, function (tag) {
          tag.replace(new RegExp("\\b(" + ATTR_CONVALIDA + ")=\"([^\"]*)\"", "g"), function (a, k, v) {
            out.push({ testo: xmlDec(v), dove: "messaggio di un elenco o convalida in «" + nome + "»" });
            return a;
          });
          return tag;
        });
      } else if (/^xl\/(drawings|charts)\/[^/]*\.xml$/.test(p)) {
        var dove = /charts/.test(p) ? "grafico" : "casella di testo o disegno";
        testoParte(zip, i).replace(/<a:p>([\s\S]*?)<\/a:p>/g, function (a, x) {
          var t = [];
          x.replace(/<a:t>([\s\S]*?)<\/a:t>/g, function (b, y) { t.push(xmlDec(y)); return b; });
          if (t.length) out.push({ testo: t.join(""), dove: dove });
          return a;
        });
      }
    });
    return out;
  }

  var PAROLE_SUGGERITE = /nom|cognom|client|fornitor|ragione|social|societ|ditta|azienda|dipendent|intestat|inquilin|conduttor|locator|propriet|referent|responsabil|contatt|beneficiar|debitor|creditor|controparte|indirizz|via\b|email|e-mail|mail|telefon|cellular|\bcf\b|codice fiscale|p\.? ?iva|partita|iban|banca/i;

  // Colonne di un foglio con la riga d'intestazione data (1 = prima riga del foglio).
  function colonne(foglio, rigaIntest) {
    var iR = rigaIntest - 1 - foglio.origine.r;
    var intest = foglio.righe[iR] || [];
    var nCol = foglio.righe.reduce(function (m, r) { return Math.max(m, r ? r.length : 0); }, 0);
    var out = [];
    for (var c = 0; c < nCol; c++) {
      var esempi = [], testi = 0, pieni = 0;
      for (var r = iR + 1; r < foglio.righe.length; r++) {
        var v = (foglio.righe[r] || [])[c];
        if (v == null || v === "") continue;
        pieni++;
        if (typeof v === "string") testi++;
        if (esempi.length < 3 && esempi.indexOf(valoreInTesto(v)) < 0) esempi.push(valoreInTesto(v));
      }
      if (!pieni && intest[c] == null) continue;
      var nomeCol = intest[c] == null ? "" : valoreInTesto(intest[c]);
      var let_ = lettera(c + foglio.origine.c);
      out.push({ indice: c, lettera: let_, intestazione: nomeCol, esempi: esempi, pieni: pieni,
        quotaTesto: pieni ? testi / pieni : 0,
        suggerita: PAROLE_SUGGERITE.test(nomeCol) && testi > 0,
        prefisso: prefissoDa(nomeCol, "COL_" + let_) });
    }
    return out;
  }

  // ------------------------------------------------------------------ testi sparsi
  // Tutti i testi distinti del file, per scegliere i nomi quando non stanno in una colonna (prospetti a blocchi).
  // Parole che non sono nomi: mesi, giorni, intestazioni e voci tipiche di bilancio. Se un testo ne contiene una,
  // non viene proposto (l'utente puo' sempre marcarlo a mano).
  var PAROLE_COMUNI = new RegExp("^(" + [
    "gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre",
    "dicembre", "luned[iì]", "marted[iì]", "mercoled[iì]", "gioved[iì]", "venerd[iì]", "sabato", "domenica",
    "totale", "totali", "subtotale", "saldo", "saldi", "note", "nota", "data", "date", "importo", "importi", "descrizione",
    "anno", "mese", "periodo", "luce", "gas", "acqua", "riscaldamento", "risc", "spese", "spesa", "entrate", "uscite",
    "costi", "costo", "ricavi", "ricavo", "media", "nome", "cognome", "nominativo", "cliente", "clienti", "fornitore",
    "fornitori", "referente", "controparte", "commessa", "progetto", "email", "e-mail", "mail", "telefono", "cellulare",
    "fax", "indirizzo", "via", "citt[aà]", "cap", "provincia", "codice", "fiscale", "partita", "p\\.?iva", "iva", "iban",
    "conto", "dipendente", "dipendenti", "unit[aà]", "fattura", "fatture", "bonifico", "pagamento", "pagamenti",
    "rimborso", "acconto", "rata", "rate", "imposta", "imposte", "accantonamento", "verificare", "riepilogo",
    "prospetto", "bilancio", "budget", "consuntivo", "preventivo", "voce", "voci", "categoria", "tipo", "stato",
    "quantit[aà]", "prezzo", "valore", "percentuale", "differenza", "variazione", "scadenza", "pagato", "incassato",
    "fatturato", "utile", "perdita", "margine", "somma", "conteggio", "etichette", "etichetta", "riga", "righe",
    "colonna", "colonne", "complessivo", "generale", "vuoto", "trimestre", "semestre", "esercizio", "consumi", "consumo",
    // parole tipiche delle intestazioni di stampa e delle note
    "pagina", "pag", "riservato", "riservata", "riservati", "confidenziale", "bozza", "attenzione", "copia",
    "uso", "interno", "interna", "documento", "versione", "aggiornato", "aggiornata", "approvato", "approvata", "stampa"
  ].join("|") + ")$", "i");
  var CONNETTIVI = /^(di|de|del|della|dello|dei|degli|delle|da|dal|dalla|e|ed|&|van|von|der|den|la|le|lo|y)$/i;
  var PAROLA_NOME = /^\p{Lu}[\p{L}'’.&-]*$/u;
  // forme societarie e parole generiche dei nomi di azienda: da sole non identificano nessuno
  var SUFFISSI = /^(s\.?r\.?l\.?|s\.?p\.?a\.?|s\.?n\.?c\.?|s\.?a\.?s\.?|s\.?s\.?|coop|cooperativa|studio|associato|associati|figli|fratelli|societ[aà]|group|gruppo|italia|italiana|holding|consorzio|servizi|ditta|impresa)$/i;

  function formaDiNome(t) {
    var parole = t.split(/\s+/);
    if (parole.length > 5 || !PAROLA_NOME.test(parole[0])) return false;
    return parole.every(function (p) { return (PAROLA_NOME.test(p) || CONNETTIVI.test(p)) && !PAROLE_COMUNI.test(p); });
  }

  // Intestazioni di colonne che contengono nomi di persone o di aziende (non indirizzi, telefoni o codici).
  var INTEST_NOMI = /(^|[^\p{L}])(nom|cognom|client|fornitor|ragione|social|societ|ditta|aziend|dipendent|intestat|inquilin|conduttor|locator|propriet|referent|responsabil|contatt|beneficiar|debitor|creditor|controparte)/iu;

  function testiDistinti(lettura) {
    var perChiave = new Map();
    // dove: «nomi» sotto un'intestazione di nomi, «altro» sotto un'altra intestazione o nella riga d'intestazione,
    // «libero» senza intestazione (prospetti a blocchi, testi fuori dalle celle)
    function conta(v, iF, dove) {
      if (typeof v !== "string") return;
      var t = v.trim().replace(/\s+/g, " ");
      if (!t || /^[\d\s.,:;\/%€$+()-]*$/.test(t)) return;
      var k = norm(t), x = perChiave.get(k);
      if (!x) { x = { testo: t, volte: 0, perFoglio: {}, dove: { nomi: 0, altro: 0, libero: 0 } }; perChiave.set(k, x); }
      x.volte++;
      x.perFoglio[iF] = (x.perFoglio[iF] || 0) + 1;
      x.dove[dove]++;
    }
    // Una riga vale come intestazione di tabella solo se almeno due celle ne hanno le parole («Referente», «Citta'»,
    // «Importo»): nei prospetti a blocchi la prima riga piena di testo e' fatta di unita' e cognomi, non di intestazioni.
    function intestazioneVera(riga) {
      return (riga || []).filter(function (v) {
        if (typeof v !== "string") return false;
        return INTEST_NOMI.test(v) || PAROLE_SUGGERITE.test(v) ||
          v.split(/[\s\/.,:;()-]+/).some(function (p) { return p && PAROLE_COMUNI.test(p); });
      }).length >= 2;
    }
    lettura.fogli.forEach(function (f, iF) {
      var h = rigaIntestazione(f);
      if (h >= 0 && !intestazioneVera(f.righe[h])) h = -1;
      var intest = h >= 0 ? f.righe[h] || [] : [];
      f.righe.forEach(function (riga, r) {
        (riga || []).forEach(function (v, c) {
          var nomeCol = intest[c] == null ? "" : valoreInTesto(intest[c]).trim();
          conta(v, iF, h < 0 || r < h ? "libero" : r === h ? "altro" : !nomeCol ? "libero" :
            INTEST_NOMI.test(nomeCol) ? "nomi" : "altro");
        });
      });
    });
    // i testi fuori dalle celle (commenti, intestazioni, caselle di testo) contano come un foglio a parte
    (lettura.extra || []).forEach(function (x) { conta(x.testo, "fuori", "libero"); });
    var out = [];
    perChiave.forEach(function (x) {
      // Un'etichetta si ripete dentro lo stesso foglio (es. «Risc» in ogni blocco); un nome compare una o due volte
      // per foglio, anche se il file ha molti fogli (un periodo per foglio). Un nome di piu' parole vale sempre.
      var maxFoglio = Math.max.apply(null, Object.keys(x.perFoglio).map(function (k) { return x.perFoglio[k]; }));
      var piuParole = /\s/.test(x.testo);
      // Una parola sola sotto un'intestazione che non parla di nomi («Citta'», «Causale») o in una riga
      // d'intestazione e' quasi sempre un'etichetta: si propone solo se compare anche sotto un'intestazione di nomi.
      var etichetta = !piuParole && x.dove.altro > 0 && x.dove.nomi === 0;
      x.proposto = (maxFoglio <= 2 || piuParole) && !etichetta && x.testo.length >= 3 && x.testo.length <= 60 &&
        formaDiNome(x.testo) && x.testo !== x.testo.toUpperCase();
      delete x.perFoglio;
      delete x.dove;
      out.push(x);
    });
    out.sort(function (a, b) { return (b.proposto - a.proposto) || (a.volte - b.volte) || a.testo.localeCompare(b.testo, "it"); });
    return out;
  }

  // Riga d'intestazione (indice in foglio.righe), o -1: la prima, tra le prime 10, con almeno meta' delle colonne
  // piene e solo testo.
  function rigaIntestazione(foglio) {
    var nCol = foglio.righe.reduce(function (m, r) { return Math.max(m, r ? r.length : 0); }, 0);
    for (var r = 0; r < Math.min(10, foglio.righe.length); r++) {
      var piene = (foglio.righe[r] || []).filter(function (v) { return v != null && v !== ""; });
      if (piene.length >= Math.max(2, nCol / 2) && piene.every(function (v) { return typeof v === "string"; })) return r;
    }
    return -1;
  }

  // Valori delle corrispondenze precedenti presenti nel file: tornano marcati con la loro categoria.
  function proposteDaCorrispondenze(lettura, righe) {
    var voci = [], cat = new Map();
    (righe || []).forEach(function (r) {
      if (!r || r[0] == null || r[1] == null) return;
      var m = /^(.*)_\d+$/.exec(String(r[0]).trim());
      if (!m || /^(IBAN|CF|EMAIL|PIVA|TEL)$/.test(m[1])) return;
      voci.push({ testo: String(r[1]), codice: m[1] });
      cat.set(norm(r[1]), m[1]);
    });
    var cerca = costruisciRicerca(voci);
    if (!cerca) return [];
    var trovati = new Map();
    lettura.fogli.forEach(function (f) {
      f.righe.forEach(function (riga) {
        (riga || []).forEach(function (v) {
          if (typeof v !== "string") return;
          cerca.re.lastIndex = 0;
          var m;
          while ((m = cerca.re.exec(v)) !== null) {
            var k = norm(m[0]);
            if (!trovati.has(k)) trovati.set(k, { testo: m[0].trim(), categoria: cat.get(k) });
          }
        });
      });
    });
    return Array.from(trovati.values());
  }

  // ------------------------------------------------------------------ anonimizzazione
  // config: { fogli: { nome: { rigaIntest, colonne: [{indice, prefisso}] } }, termini: "testo => PREFISSO\n...",
  //           regole: {iban,cf,email,piva,tel: bool}, ovunque: bool, rinominaFogli: bool }
  function anonimizza(lettura, config, righeCorrispondenze) {
    var mappa = new Mappa();
    if (righeCorrispondenze) mappa.carica(righeCorrispondenze);
    var rep = { fogli: [], termini: 0, regole: {}, nonCercati: [], celleColonne: 0, occorrenzeTesto: 0 };
    ORDINE_REGOLE.forEach(function (k) { rep.regole[k] = 0; });

    // 1. codici per i valori delle colonne scelte
    var voci = [];
    lettura.fogli.forEach(function (f) {
      var cf = config.fogli[f.nome];
      if (!cf) return;
      var iR = cf.rigaIntest - 1 - f.origine.r;
      cf.colonne.forEach(function (col) {
        var origine = f.nome + " / " + col.prefisso;
        for (var r = iR + 1; r < f.righe.length; r++) {
          var v = (f.righe[r] || [])[col.indice];
          if (v == null || v === "") continue;
          var t = valoreInTesto(v).trim();
          if (!t) continue;
          var codice = mappa.codice(t, col.prefisso, origine);
          if (config.ovunque && typeof v === "string") {
            if (t.length >= MIN_LUNGHEZZA_OVUNQUE && !/^[\d\s.,/-]+$/.test(t)) voci.push({ testo: t, codice: codice });
            else if (rep.nonCercati.indexOf(t) < 0) rep.nonCercati.push(t);
          }
        }
      });
    });

    // 2a. testi marcati uno per uno, ciascuno con la sua categoria: ogni cella che li contiene, ovunque nel file.
    //     Una voce e' una stringa (categoria = prefissoValori) oppure {testo, categoria}.
    (config.valori || []).forEach(function (x) {
      var t = String(typeof x === "string" ? x : x.testo).trim();
      if (!t) return;
      var pref = prefissoDa(typeof x === "string" ? "" : x.categoria, config.prefissoValori || "NOME");
      voci.push({ testo: t, codice: mappa.codice(t, pref, "marcati") });
      rep.termini++;
    });

    // 2b. elenco dei termini scritto a mano
    String(config.termini || "").split(/\r?\n/).forEach(function (riga) {
      var parti = riga.split("=>"), t = parti[0].trim();
      if (!t) return;
      var pref = prefissoDa(parti[1] || "", "TERMINE");
      voci.push({ testo: t, codice: mappa.codice(t, pref, "elenco") });
      rep.termini++;
    });
    var ricerca = costruisciRicerca(voci);

    // Il testo si cerca in forma NFC (accenti composti); se non cambia niente torna identico all'originale.
    function anonimizzaTesto(s) {
      var nfc = s.normalize("NFC"), nuovo = sostituisci(nfc);
      return nuovo === nfc ? s : nuovo;
    }
    function sostituisci(s) {
      var segs = [{ t: s }];
      ORDINE_REGOLE.forEach(function (k) {
        if (!config.regole || !config.regole[k]) return;
        var R = REGOLE[k];
        segs = applicaRe(segs, R.re, function (trovato) {
          var n = R.valida(trovato);
          if (!n) return null;
          rep.regole[k]++;
          var valore = trovato.slice(0, n).trim();
          var chiave = k === "iban" || k === "piva" || k === "tel" ? valore.replace(/[\s.\/-]/g, "") : valore;
          return { codice: mappa.codice(valore, R.prefisso, "regola " + k, chiave), lunghezza: n };
        });
      });
      if (ricerca) {
        segs = applicaRe(segs, ricerca.re, function (trovato) {
          var c = ricerca.perChiave.get(norm(trovato));
          if (!c) return null;
          rep.occorrenzeTesto++;
          return { codice: c, lunghezza: trovato.length };
        });
      }
      return segs.map(function (x) { return x.t; }).join("");
    }

    // P.IVA salvata come numero (lo zero iniziale si perde). Circa un numero intero su dieci di quella lunghezza
    // passa il controllo per caso: per non trasformare importi o codici in testo (le somme li salterebbero), vale
    // solo dove l'intestazione della colonna o l'etichetta a sinistra dice P.IVA. Ritorna il codice o null.
    var numeriPiva = new Set();
    if (config.regole && config.regole.piva) lettura.fogli.forEach(function (f) {
      // Per colonna: c'e' un'etichetta «P.IVA» sopra? Resta valida scendendo, anche dopo un testo qualunque nella
      // colonna («estero», «n.d.»); smette solo quando sotto compare un'altra etichetta di colonna («Importo»,
      // «Totale», «Cliente»...).
      var etichetta = [];
      f.righe.forEach(function (riga) {
        riga = riga || [];
        riga.forEach(function (v, c) {
          if (typeof v === "string" && v.trim()) {
            var t = v.trim();
            if (ETICHETTA_PIVA.test(t)) etichetta[c] = true;
            else if (t.split(/\s+/).length <= 3 && (PAROLE_SUGGERITE.test(t) || t.split(/\s+/).some(function (p) { return PAROLE_COMUNI.test(p); })))
              etichetta[c] = false;
            return;
          }
          if (!pivaNumerica(v)) return;
          var sinistra = null;
          for (var j = c - 1; j >= 0 && sinistra == null; j--) if (riga[j] != null && riga[j] !== "") sinistra = riga[j];
          if (etichetta[c] || (typeof sinistra === "string" && ETICHETTA_PIVA.test(sinistra))) numeriPiva.add(v);
        });
      });
    });
    function anonimizzaNumero(v) {
      if (!numeriPiva.has(v)) return null;
      rep.regole.piva++;
      return mappa.codice(String(v).padStart(11, "0"), "PIVA", "regola piva numero");
    }

    // 3. nuovo file, solo valori
    var wbOut = XLSX.utils.book_new();
    wbOut.Props = { Title: "", Subject: "", Author: "", Manager: "", Company: "", Category: "", Keywords: "", Comments: "", LastAuthor: "" };
    var nomiUsati = {};
    var restiTesto = new Map(), uscite = [];
    lettura.fogli.forEach(function (f, iF) {
      var cf = config.fogli[f.nome] || { rigaIntest: 1, colonne: [] };
      var iR = cf.rigaIntest - 1 - f.origine.r;
      var scelte = new Set(cf.colonne.map(function (c) { return c.indice; }));
      var celleCol = 0;
      var aoa = f.righe.map(function (riga, r) {
        return (riga || []).map(function (v, c) {
          if (v == null || v === "") return v;
          if (r > iR && scelte.has(c)) {
            celleCol++;
            return mappa.perChiave.get(norm(valoreInTesto(v).trim()));
          }
          if (typeof v === "string") {
            var nuovo = anonimizzaTesto(v);
            var resto = nuovo.replace(/\b[A-Z][A-Z0-9_]*_\d{3,}\b/g, "").replace(/[\s\p{P}]+/gu, "");
            if (resto && /\p{L}/u.test(resto)) {
              var k = nuovo.trim();
              restiTesto.set(k, (restiTesto.get(k) || 0) + 1);
            }
            return nuovo;
          }
          var cod = anonimizzaNumero(v);
          return cod === null ? v : cod;
        });
      });
      rep.celleColonne += celleCol;
      uscite.push(aoa);
      var ws = XLSX.utils.aoa_to_sheet(aoa, { origin: f.origine, cellDates: true });
      if (f.unioni.length) ws["!merges"] = f.unioni;
      if (f.colonne) ws["!cols"] = f.colonne;
      var nome = config.rinominaFogli ? "Foglio" + (iF + 1)
        : anonimizzaTesto(f.nome).replace(/[\[\]:*?\/\\]/g, "_").slice(0, 31) || "Foglio" + (iF + 1);
      var base = nome, n = 2;
      while (nomiUsati[nome.toLowerCase()]) nome = base.slice(0, 28) + "_" + (n++);
      nomiUsati[nome.toLowerCase()] = true;
      XLSX.utils.book_append_sheet(wbOut, ws, nome);
      rep.fogli.push({ nome: nome, nascosto: f.nascosto, righe: f.righe.length, colonneScelte: cf.colonne.length,
        celleColonne: celleCol, commenti: f.commenti, formule: f.formule, collegamenti: f.collegamenti });
    });
    wbOut.Workbook = { Sheets: lettura.fogli.map(function (f) { return { Hidden: f.nascosto ? 1 : 0 }; }) };

    // Frammenti: una parola distintiva di un nome marcato che compare da sola nel file anonimo (es. «Edilnord» da
    // «Edilnord Srl» nel nome di un foglio o in una descrizione). E' la forma breve che sfugge alla ricerca esatta.
    var distintive = new Map();
    mappa.perCodice.forEach(function (v) {
      if (/^regola/.test(v.origine)) return;
      String(v.valore).split(/[\s,;:()\/]+/).forEach(function (p) {
        if (p.length >= 4 && PAROLA_NOME.test(p) && !PAROLE_COMUNI.test(p) && !CONNETTIVI.test(p) && !SUFFISSI.test(p))
          distintive.set(norm(p), p);
      });
    });
    var frammenti = [], visti = new Set();
    function cercaFrammenti(testo, dove) {
      if (typeof testo !== "string" || !distintive.size) return;
      testo.replace(/\b[A-Z][A-Z0-9_]*_\d{3,}\b/g, " ").split(/[^\p{L}'’.&-]+/u).forEach(function (p) {
        p = p.replace(/[.'’-]+$/, "");
        var k = norm(p);
        if (!distintive.has(k) || !PAROLA_NOME.test(p)) return;   // solo con la maiuscola: «verde» in «area verde» no
        var chiave = dove + "|" + k;
        if (visti.has(chiave)) return;
        visti.add(chiave);
        frammenti.push({ parola: distintive.get(k), testo: testo.trim(), dove: dove });
      });
    }
    wbOut.SheetNames.forEach(function (n) { cercaFrammenti(n, "nome del foglio"); });
    uscite.forEach(function (aoa) { aoa.forEach(function (riga) { (riga || []).forEach(function (v) { cercaFrammenti(v, "cella"); }); }); });

    // Testi fuori dalle celle: come li vedra' Claude, e quali hanno ancora parole in chiaro.
    var fuori = (lettura.extra || []).map(function (x) {
      var dopo = anonimizzaTesto(x.testo);
      cercaFrammenti(dopo, x.dove);
      return { dove: x.dove, prima: x.testo, dopo: dopo, nome: sembraNome(dopo) };
    });

    var resti = [];
    restiTesto.forEach(function (n, t) { resti.push({ testo: t, volte: n, nome: sembraNome(t) }); });
    fuori.forEach(function (x) {
      var resto = x.dopo.replace(/\b[A-Z][A-Z0-9_]*_\d{3,}\b/g, "").replace(/[\s\p{P}]+/gu, "");
      if (resto && /\p{L}/u.test(resto)) resti.push({ testo: x.dopo, volte: 1, nome: x.nome, dove: x.dove });
    });
    resti.sort(function (a, b) { return (b.nome - a.nome) || (b.volte - a.volte) || (a.testo < b.testo ? -1 : 1); });
    rep.nuoviCodici = mappa.nuovi;
    rep.nomiDefiniti = lettura.nomiDefiniti;
    return { wb: wbOut, mappa: mappa, report: rep, resti: resti, righe: uscite, frammenti: frammenti, fuori: fuori,
      nomiFogli: wbOut.SheetNames.slice(), trasforma: { testo: anonimizzaTesto, numero: anonimizzaNumero } };
  }

  // ------------------------------------------------------------------ file anonimo con la formattazione
  // Il .xlsx e' un archivio zip di file XML. Invece di ricostruire il file dai valori (che perde stili, formule,
  // larghezze, grafici), si sostituiscono i testi dentro l'XML originale con la stessa trasformazione dell'anteprima.
  // Ogni posto dove un nome puo' stare ha la sua regola; alla fine un controllo cerca i valori veri in tutto il pacchetto.
  function xmlDec(s) {
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, function (m, e) {
      if (e[0] === "#") return String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
      return { amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'" }[e];
    });
  }
  function xmlEnc(s, attributo) {
    s = String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    // in un attributo a capo e tabulazione scritti cosi' come sono diventerebbero spazi: vanno codificati
    return attributo ? s.replace(/"/g, "&quot;").replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;") : s;
  }
  function eZip(buf) { return buf && buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4B; }

  // Contenuto di <si> o <is>: uno o piu' run <r><t>. Se il nome sta dentro un run si tiene la formattazione interna;
  // se e' spezzato tra due run, la cella diventa un testo unico (resta lo stile della cella).
  function trasformaRuns(inner, T) {
    var re = /(<rPh\b[\s\S]*?<\/rPh>)|<t(?:\s[^>]*)?\/>|<t((?:\s[^>]*[^\/>])?)>([\s\S]*?)<\/t>/g, parti = [], m;
    while ((m = re.exec(inner)) !== null) if (!m[1]) parti.push(m[3] === undefined ? "" : xmlDec(m[3]));
    if (!parti.length) return inner;
    var intero = parti.join(""), nuovo = T(intero);
    if (nuovo === intero) return inner;
    var perRun = parti.length === 1 ? [nuovo] : parti.map(T);
    if (perRun.join("") === nuovo) {
      var i = 0;
      return inner.replace(re, function (all, ph) {
        return ph ? ph : "<t xml:space=\"preserve\">" + xmlEnc(perRun[i++]) + "</t>";
      });
    }
    return "<t xml:space=\"preserve\">" + xmlEnc(nuovo) + "</t>";
  }
  // Paragrafo di una casella di testo o di un grafico (<a:p>): uno o piu' run <a:t>. Come per le celle: se il nome
  // e' spezzato tra due run («Mar» + «io Rossi» in grassetto), il paragrafo intero va nel primo run e gli altri si
  // svuotano (resta la formattazione del primo).
  function trasformaParagrafo(inner, T) {
    var re = /<a:t>([\s\S]*?)<\/a:t>|<a:t\/>/g, parti = [], m;
    while ((m = re.exec(inner)) !== null) parti.push(m[1] === undefined ? "" : xmlDec(m[1]));
    if (!parti.length) return inner;
    var intero = parti.join(""), nuovo = T(intero);
    if (nuovo === intero) return inner;
    var perRun = parti.length === 1 ? [nuovo] : parti.map(T), i = 0;
    var unico = perRun.join("") !== nuovo;
    return inner.replace(re, function () {
      var t = unico ? (i === 0 ? nuovo : "") : perRun[i];
      i++;
      return "<a:t>" + xmlEnc(t) + "</a:t>";
    });
  }
  // Testi tra virgolette dentro una formula ("" e' la virgoletta)
  function letterali(f, T) {
    return f.replace(/"((?:[^"]|"")*)"/g, function (a, x) { return "\"" + T(x.replace(/""/g, "\"")).replace(/"/g, "\"\"") + "\""; });
  }
  function attributi(tag, nomi, T) {
    return tag.replace(new RegExp("(^|\\s)(" + nomi + ")=\"([^\"]*)\"", "g"), function (a, sp, k, v) {
      var prima = xmlDec(v), dopo = T(prima);
      return dopo === prima ? a : sp + k + "=\"" + xmlEnc(dopo, true) + "\"";
    });
  }
  // Ogni testo e ogni attributo di una parte XML che non ha una regola sua (affettatori, query, controlli, dati
  // ricchi...). T cambia solo i nomi marcati, quindi i riferimenti tra le parti restano coerenti.
  function trasformaGenerico(s, T) {
    // le sezioni CDATA prima: il loro testo non e' codificato, e dentro possono esserci < e >
    var cdata = [];
    s = s.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, function (a, x) {
      var dopo = T(x);
      cdata.push(dopo === x ? a : "<![CDATA[" + dopo.replace(/\]\]>/g, "]]]]><![CDATA[>") + "]]>");
      return "\u0000" + (cdata.length - 1) + "\u0000";
    });
    s = s.replace(/<[^>]+>/g, function (tag) {
      if (/^<[?!\/]/.test(tag)) return tag;
      return tag.replace(/(\s)([\w:.-]+)=("([^"]*)"|'([^']*)')/g, function (a, sp, k, q, v1, v2) {
        if (/^xmlns/.test(k)) return a;
        var prima = xmlDec(v1 != null ? v1 : v2), dopo = T(prima);
        return dopo === prima ? a : sp + k + "=\"" + xmlEnc(dopo, true) + "\"";
      });
    }).replace(/>([^<]+)</g, function (a, x) {
      if (/\u0000/.test(x)) return a;
      var prima = xmlDec(x), dopo = T(prima);
      return dopo === prima ? a : ">" + xmlEnc(dopo) + "<";
    });
    return s.replace(/\u0000(\d+)\u0000/g, function (a, i) { return cdata[+i]; });
  }

  // Percorsi con il nome dell'utente Windows o di OneDrive («C:\Users\mrossi\...», «.../personal/mario_rossi_.../»):
  // Excel li salva nel file (cartella di lavoro, collegamenti, connessioni). Il nome utente diventa «utente».
  var RE_PERCORSO_UTENTE = /((?:^|[^\p{L}\p{N}])[A-Za-z]:(?:\\|\/|\\\\)(?:Users|Utenti|Documents and Settings)(?:\\|\/|\\\\))([^\\\/"<>&|]+)/giu;
  var RE_PERCORSO_PERSONALE = /((?:sharepoint\.com|onedrive\.com|-my\.sharepoint\.com)\/personal\/)([^\/"<>&?#]+)/gi;
  var RE_PERCORSO_MAC = /((?:^|[\s"'=(\/])\/Users\/)([^\/"<>&]+)/g;   // anche «file:///Users/nome/»
  function togliUtente(s) {
    return s.replace(RE_PERCORSO_UTENTE, function (a, pre, nome) { return /^utente$/i.test(nome) ? a : pre + "utente"; })
      .replace(RE_PERCORSO_PERSONALE, function (a, pre, nome) { return /^utente$/i.test(nome) ? a : pre + "utente"; })
      .replace(RE_PERCORSO_MAC, function (a, pre, nome) { return /^(utente|Shared)$/i.test(nome) ? a : pre + "utente"; });
  }

  // Parti del pacchetto. Ogni parte deve avere una regola: quelle senza regola, se non sono XML leggibile, non si
  // possono controllare e il file formattato non esce (si ripiega sui soli valori).
  var PARTI_IMMAGINE = /^xl\/media\/[^/]*\.(png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$/i;
  // Parti che si tolgono (con le relazioni e i tipi che le nominano): stampanti, macro, e i metadati che Google Fogli
  // aggiunge quando esporta in Excel (in un formato suo, non leggibile qui; quelli dei commenti possono contenere gli
  // account di chi ha commentato). Excel non li usa.
  var PARTI_DA_TOGLIERE = /^xl\/(printerSettings\/[^/]*\.bin|vbaProject\.bin|vbaProjectSignature\.bin|metadata|commentsmeta\d*)$/i;
  var PARTI_TESTO = /\.(xml|rels|vml)$/i;

  function anonimizzaXlsx(buffer, esito) {
    var avvisi = [];
    if (!eZip(buffer)) return { buffer: null, motivo: "formato" };
    var zip = XLSX.CFB.read(buffer, { type: "array" });
    var percorsi = zip.FullPaths.map(percorsoParte);
    if (percorsi.indexOf("xl/workbook.bin") >= 0) return { buffer: null, motivo: "xlsb" };
    var dec = new TextDecoder("utf-8"), enc = new TextEncoder();

    // Parti che non si possono leggere ne' ripulire: modello dati di Power Pivot, oggetti incorporati (altri file
    // Office), controlli ActiveX, query di Power Query. Dentro possono avere una copia dei dati.
    var illeggibili = [];
    percorsi.forEach(function (p, i) {
      var f = zip.FileIndex[i];
      if (f.type !== 2 || !f.content || p === "" || /Sh33tJ5/.test(p)) return;
      if (PARTI_TESTO.test(p)) {
        if (/^customXml\/item\d*\.xml$/i.test(p) && /DataMashup/.test(dec.decode(new Uint8Array(f.content)))) illeggibili.push(p);
        return;
      }
      if (PARTI_IMMAGINE.test(p) || PARTI_DA_TOGLIERE.test(p)) return;
      illeggibili.push(p);
    });
    if (illeggibili.length) return { buffer: null, motivo: "parti", parti: illeggibili };

    // contatori dell'anteprima: questa seconda passata non li deve gonfiare
    var rep = esito.report, salva = JSON.stringify({ r: rep.regole, o: rep.occorrenzeTesto });
    var T = function (s) { return esito.trasforma.testo(s); };
    var N = esito.trasforma.numero;
    var ND = esito.trasforma.numeroDaTesto;

    // fogli: nome vecchio -> nuovo, nell'ordine di workbook.xml (lo stesso dell'anteprima)
    // Tutte le rinomine in un solo passaggio: una dopo l'altra, «Riepilogo» -> «Foglio1» e poi «Foglio1» -> «Foglio2»
    // manderebbero la formula sul foglio sbagliato.
    var rinomine = [], reRinomina = null, perForma = null;
    function rinomina(testo) {
      if (!rinomine.length) return testo;
      if (!reRinomina) {
        perForma = new Map();
        var alt = [];
        rinomine.forEach(function (x) {
          var q = "'" + x.da.replace(/'/g, "''") + "'!";
          perForma.set(q, x.citatoA + "!");
          alt.push(escapeRe(q));
          if (FOGLIO_SEMPLICE.test(x.da)) { perForma.set(x.da + "!", x.citatoA + "!"); alt.push("(?<![\\p{L}\\p{N}_.'\\]])" + escapeRe(x.da) + "!"); }
        });
        alt.sort(function (a, b) { return b.length - a.length; });
        reRinomina = new RegExp(alt.join("|"), "gu");
      }
      return testo.replace(reRinomina, function (m) { return perForma.get(m) || m; });
    }
    function formula(x) { return xmlEnc(rinomina(letterali(xmlDec(x), T))); }

    var testo = function (i) { return dec.decode(new Uint8Array(zip.FileIndex[i].content)); };
    var scrivi = function (i, s) { var b = enc.encode(s); zip.FileIndex[i].content = b; zip.FileIndex[i].size = b.length; };

    // Parti da togliere, e gli Id delle relazioni che puntano a loro, per parte di partenza. Un elemento che richiama
    // una relazione tolta va tolto anche lui: Google Fogli scrive in workbook.xml <go:sheetsCustomData r:id="…"/>, e
    // con la relazione tolta e il richiamo rimasto Excel non apre il file (prova sui file veri, 13/10). Le stampanti
    // fanno eccezione: di <pageSetup> si toglie solo l'attributo r:id, il resto dell'impostazione di pagina resta.
    var daTogliere = percorsi.filter(function (p) { return PARTI_DA_TOGLIERE.test(p); });
    var idTolti = {};
    if (daTogliere.length) percorsi.forEach(function (p, i) {
      var m = /^(.*?)_rels\/([^/]+)\.rels$/.exec(p);
      if (!m || !zip.FileIndex[i].content) return;
      testo(i).replace(/<Relationship\b[^>]*>/g, function (tag) {
        var t = /\sTarget="([^"]*)"/.exec(tag), id = /\sId="([^"]*)"/.exec(tag);
        if (t && id && !/printerSettings/.test(t[1]) &&
            daTogliere.some(function (d) { return d.split("/").pop() === t[1].split("/").pop(); }))
          (idTolti[m[1] + m[2]] = idTolti[m[1] + m[2]] || []).push(id[1]);
        return tag;
      });
    });
    function togliRichiami(p, s) {
      var ids = idTolti[p];
      if (!ids) return s;
      var alt = ids.map(escapeRe).join("|");
      return s.replace(new RegExp("<([\\w:]+)\\b[^>]*\\sr:id=\"(?:" + alt + ")\"[^>]*/>", "g"), "")
        .replace(new RegExp("<([\\w:]+)\\b[^>]*\\sr:id=\"(?:" + alt + ")\"[^>]*>[\\s\\S]*?</\\1>", "g"), "");
    }

    // 1. workbook.xml per primo: rinomina dei fogli e nomi definiti; via il percorso della cartella di lavoro
    //    (<x15ac:absPath>) e il nome di chi ha protetto il file in scrittura (<fileSharing userName>).
    var iWb = percorsi.indexOf("xl/workbook.xml");
    if (iWb >= 0) {
      var nFoglio = 0, wb = testo(iWb);
      wb = wb.replace(/<sheet\b[^>]*>/g, function (tag) {
        var nuovo = esito.nomiFogli[nFoglio++];
        return tag.replace(/\bname="([^"]*)"/, function (a, v) {
          var da = xmlDec(v);
          if (nuovo && nuovo !== da) rinomine.push({ da: da, a: nuovo, citatoA: foglioCitato(nuovo) });
          return "name=\"" + xmlEnc(nuovo || da, true) + "\"";
        });
      });
      wb = wb.replace(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g, function (a, at, x) {
        return "<definedName" + at + ">" + formula(x) + "</definedName>";
      }).replace(/<mc:AlternateContent\b(?:(?!<\/mc:AlternateContent>)[\s\S])*?absPath[\s\S]*?<\/mc:AlternateContent>/g, "")
        .replace(/<x15ac:absPath\b[^>]*\/>/g, "")
        .replace(/<fileSharing\b[^>]*>/g, function (tag) { return tag.replace(/\buserName="[^"]*"/, "userName=\"\""); })
        // visualizzazioni personali: «Mario Rossi - Visualizzazione personale»
        .replace(/<customWorkbookView\b[^>]*>/g, function (tag) { return attributi(tag, "name", T); });
      scrivi(iWb, togliUtente(togliRichiami("xl/workbook.xml", wb)));
    }

    // autori dei commenti: «Autore 1», «Autore 2»... anche nella riga in grassetto con cui Excel apre ogni commento
    var autori = new Map(), persone = new Map();
    function sostituisciAutori(s) {
      autori.forEach(function (nuovo, vecchio) {
        if (vecchio.length >= 2) s = s.split(vecchio).join(nuovo);
      });
      return s;
    }
    percorsi.forEach(function (p, i) {
      if (!/^xl\/comments[^/]*\.xml$/.test(p) && !/^xl\/comments\/[^/]*\.xml$/.test(p)) return;
      testo(i).replace(/<author>([\s\S]*?)<\/author>/g, function (a, x) {
        var k = xmlDec(x).trim();
        // «tc={...}» non e' un nome: lega il commento classico al commento moderno e deve restare com'e'
        if (k && !/^tc=/.test(k) && !autori.has(k)) autori.set(k, T(k) !== k ? T(k) : "Autore " + (autori.size + 1));
        return a;
      });
    });

    // Ripristino: i testi condivisi (sharedStrings), per rimettere numero una P.IVA che nell'originale era numero
    // anche quando Excel o Claude hanno salvato la cella come testo condiviso.
    var condivisi = null;
    if (ND) {
      var iSs = percorsi.indexOf("xl/sharedStrings.xml");
      condivisi = [];
      if (iSs >= 0) testo(iSs).replace(/<si>([\s\S]*?)<\/si>/g, function (a, x) {
        var t = "";
        x.replace(/(<rPh\b[\s\S]*?<\/rPh>)|<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, function (b, ph, y) { if (!ph) t += xmlDec(y); return b; });
        condivisi.push(t);
        return a;
      });
    }

    var tolte = [];
    var tolgoStampanti = daTogliere.some(function (p) { return /^xl\/printerSettings\//.test(p); });
    if (daTogliere.some(function (p) { return /^xl\/(metadata|commentsmeta\d*)$/i.test(p); }))
      avvisi.push("Tolti i metadati che Google Fogli aggiunge al file: non si possono controllare.");
    var macro = percorsi.indexOf("xl/vbaProject.bin") >= 0;
    percorsi.forEach(function (p, i) {
      if (i === iWb || !zip.FileIndex[i].content || zip.FileIndex[i].type !== 2) return;
      if (PARTI_DA_TOGLIERE.test(p)) { tolte.push(p); return; }
      if (!PARTI_TESTO.test(p)) return;
      var s;
      if (p === "xl/sharedStrings.xml") {
        s = testo(i).replace(/<si>([\s\S]*?)<\/si>/g, function (a, x) { return "<si>" + trasformaRuns(x, T) + "</si>"; });
      } else if (/^xl\/worksheets\/[^/]*\.xml$/.test(p) || /^xl\/externalLinks\/externalLink[^/]*\.xml$/.test(p)) {
        s = testo(i).replace(/<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g, function (all, at, ch, inner) {
          if (inner === undefined) return all;
          var t = /\bt="([^"]*)"/.exec(at);
          t = t ? t[1] : "n";
          if (t === "s" && condivisi) {
            var iv = /<v(?:\s[^>]*)?>(\d+)<\/v>/.exec(inner), ns = iv ? ND(condivisi[+iv[1]] || "") : null;
            if (ns !== null) return "<c" + at.replace(/\s*\bt="[^"]*"/, "") + "><v>" + ns + "</v></c>";
          }
          if (t === "inlineStr") {
            if (ND) {
              var solo = /^<is>\s*<t(?:\s[^>]*)?>([^<]*)<\/t>\s*<\/is>$/.exec(inner.trim()), n = solo ? ND(xmlDec(solo[1])) : null;
              if (n !== null) return "<c" + at.replace(/\s*\bt="[^"]*"/, "") + "><v>" + n + "</v></c>";
            }
            return "<c" + at + ">" + inner.replace(/<is>([\s\S]*?)<\/is>/, function (a, x) { return "<is>" + trasformaRuns(x, T) + "</is>"; }) + "</c>";
          }
          if (t === "str") return "<c" + at + ">" + inner.replace(/<v(\s[^>]*)?>([\s\S]*?)<\/v>/, function (a, at2, x) { return "<v" + (at2 || "") + ">" + xmlEnc(T(xmlDec(x))) + "</v>"; }) + "</c>";
          if (t === "n" && !/<f[\s>\/]/.test(inner)) {
            var v = /<v(?:\s[^>]*)?>([^<]*)<\/v>/.exec(inner), cod = v ? N(+v[1]) : null;
            if (cod) return "<c" + at.replace(/\s*\bt="[^"]*"/, "") + " t=\"inlineStr\"><is><t>" + xmlEnc(cod) + "</t></is></c>";
          }
          return all;
        });
        s = s.replace(/<f((?:\s[^>]*[^\/>])?)>([\s\S]*?)<\/f>/g, function (a, at, x) { return "<f" + at + ">" + formula(x) + "</f>"; })
          .replace(/<(formula|formula1|formula2|xm:f|x12ac:list)>([\s\S]*?)<\/\1>/g, function (a, tag, x) { return "<" + tag + ">" + formula(x) + "</" + tag + ">"; })
          .replace(TAG_INTEST, function (a, tag, x) { return "<" + tag + ">" + xmlEnc(trasformaIntestazione(xmlDec(x), T)) + "</" + tag + ">"; })
          .replace(/<hyperlink\b[^>]*>/g, function (tag) { return attributi(attributi(tag, "location", function (v) { return T(rinomina(v)); }), "display|tooltip", T); })
          .replace(/<(filter|customFilter)\b[^>]*>/g, function (tag) { return attributi(tag, "val", T); })
          .replace(/<(dataValidation|x14:dataValidation)\b[^>]*>/g, function (tag) { return attributi(tag, ATTR_CONVALIDA, T); })
          .replace(/<scenario\b[^>]*>/g, function (tag) { return attributi(tag, "name|user|comment", T); })
          .replace(/<customSheetView\b[^>]*>/g, function (tag) { return attributi(tag, "name", T); });
      } else if (/^xl\/worksheets\/_rels\/[^/]*\.rels$/.test(p) || /^xl\/drawings\/_rels\/[^/]*\.rels$/.test(p) ||
                 /^xl\/externalLinks\/_rels\/[^/]*\.rels$/.test(p)) {
        s = testo(i).replace(/<Relationship\b[^>]*TargetMode="External"[^>]*>/g, function (tag) { return attributi(tag, "Target", T); });
      } else if (/^xl\/comments[^/]*\.xml$/.test(p) || /^xl\/comments\/[^/]*\.xml$/.test(p)) {
        s = testo(i).replace(/<author>([\s\S]*?)<\/author>/g, function (a, x) {
          var k = xmlDec(x).trim();
          return /^tc=/.test(k) ? a : "<author>" + xmlEnc(autori.get(k) || "Autore") + "</author>";
        }).replace(/<text>([\s\S]*?)<\/text>/g, function (a, x) {
          return "<text>" + trasformaRuns(x, function (t) { return sostituisciAutori(T(t)); }) + "</text>";
        });
      } else if (/^xl\/threadedComments\/[^/]*\.xml$/.test(p)) {
        // le menzioni (@nome) puntano a posizioni nel testo che cambiano lunghezza: si tolgono, il testo resta
        s = testo(i).replace(/<mentions>[\s\S]*?<\/mentions>/g, "")
          .replace(/<text>([\s\S]*?)<\/text>/g, function (a, x) { return "<text>" + xmlEnc(sostituisciAutori(T(xmlDec(x)))) + "</text>"; });
      } else if (/^xl\/persons\/[^/]*\.xml$/.test(p)) {
        s = testo(i).replace(/<person\b[^>]*>/g, function (tag) {
          var n = persone.size + 1;
          persone.set(tag, n);
          return tag.replace(/\bdisplayName="[^"]*"/, "displayName=\"Persona " + n + "\"").replace(/\buserId="[^"]*"/, "userId=\"persona_" + n + "\"");
        });
      } else if (/^xl\/tables\/[^/]*\.xml$/.test(p)) {
        s = testo(i).replace(/<tableColumn\b[^>]*>/g, function (tag) { return attributi(tag, "name|dataCellStyle", T); })
          .replace(/<(filter|customFilter)\b[^>]*>/g, function (tag) { return attributi(tag, "val", T); })
          .replace(/<(calculatedColumnFormula|totalsRowFormula)((?:\s[^>]*[^\/>])?)>([\s\S]*?)<\/\1>/g, function (a, tag, at, x) { return "<" + tag + at + ">" + formula(x) + "</" + tag + ">"; });
      } else if (/^xl\/pivotCache\/[^/]*\.xml$/.test(p)) {
        s = testo(i).replace(/<worksheetSource\b[^>]*>/g, function (tag) {
          return tag.replace(/\bsheet="([^"]*)"/, function (a, v) {
            var x = rinomine.filter(function (r) { return r.da === xmlDec(v); })[0];
            return "sheet=\"" + xmlEnc(x ? x.a : xmlDec(v), true) + "\"";
          });
        }).replace(/<(s|cacheField)\b[^>]*>/g, function (tag) { return attributi(tag, "v|name|caption", T); })
          .replace(/<(formula)>([\s\S]*?)<\/\1>/g, function (a, tag, x) { return "<formula>" + formula(x) + "</formula>"; });
      } else if (/^xl\/pivotTables\/[^/]*\.xml$/.test(p)) {
        s = testo(i).replace(/<(item|dataField|pivotField|pivotTableDefinition|pageField|field)\b[^>]*>/g, function (tag) { return attributi(tag, "n|name|caption|dataCaption", T); });
      } else if (/^xl\/charts\/[^/]*\.xml$/.test(p) || /^xl\/drawings\/[^/]*\.xml$/.test(p)) {
        // Il prefisso dei tag («c:strCache» o «strCache») dipende da chi ha scritto il file: Excel lo mette, openpyxl
        // (usato spesso dagli LLM per salvare gli Excel) no. Le regole valgono per entrambi.
        s = testo(i).replace(/<((?:\w+:)?(?:strCache|multiLvlStrCache))>([\s\S]*?)<\/\1>/g, function (a, tag, x) {
          return "<" + tag + ">" + x.replace(/<((?:\w+:)?v)(\s[^>]*)?>([\s\S]*?)<\/\1>/g, function (b, v, at2, y) { return "<" + v + (at2 || "") + ">" + xmlEnc(T(xmlDec(y))) + "</" + v + ">"; }) + "</" + tag + ">";
        }).replace(/<((?:\w+:)?f)>([\s\S]*?)<\/\1>/g, function (a, tag, x) { return "<" + tag + ">" + formula(x) + "</" + tag + ">"; })
          .replace(/<a:p>([\s\S]*?)<\/a:p>/g, function (a, x) { return "<a:p>" + trasformaParagrafo(x, T) + "</a:p>"; })
          .replace(/<a:t>([\s\S]*?)<\/a:t>/g, function (a, x) { return "<a:t>" + xmlEnc(T(xmlDec(x))) + "</a:t>"; })
          .replace(/<((?:\w+:)?(?:cNvPr|tx))\b[^>]*>/g, function (tag) { return attributi(tag, "descr|title|name", T); });
      } else if (p === "xl/styles.xml") {
        // nomi degli stili di cella (i formati numerici si ripuliscono sotto, per tutte le parti)
        s = testo(i).replace(/<cellStyle\b[^>]*>/g, function (tag) { return attributi(tag, "name", T); });
      } else if (p === "docProps/core.xml") {
        s = testo(i).replace(/<(dc:creator|cp:lastModifiedBy|dc:title|dc:subject|dc:description|cp:keywords|cp:category|cp:contentStatus)(\s[^>]*)?>[\s\S]*?<\/\1>/g, "<$1$2></$1>");
      } else if (p === "docProps/app.xml") {
        s = testo(i).replace(/<(Company|Manager|HyperlinkBase)>[\s\S]*?<\/\1>/g, "<$1></$1>")
          .replace(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/g, function (a, x) {
            var v = xmlDec(x), r = rinomine.filter(function (y) { return y.da === v; })[0];
            return "<vt:lpstr>" + xmlEnc(r ? r.a : T(rinomina(v))) + "</vt:lpstr>";
          });
        // il resto (nome del programma, versione...): un nome marcato che coincide si sostituisce anche li'
        s = trasformaGenerico(s, T);
      } else if (p === "docProps/custom.xml") {
        s = testo(i).replace(/<(vt:lpwstr|vt:lpstr)>([\s\S]*?)<\/\1>/g, function (a, tag, x) { return "<" + tag + ">" + xmlEnc(T(xmlDec(x))) + "</" + tag + ">"; });
      } else if (p === "[Content_Types].xml" || /(^|\/)_rels\/[^/]*\.rels$/.test(p) || /^xl\/theme\//.test(p) || p === "xl/calcChain.xml") {
        s = testo(i);   // struttura del pacchetto, temi, ordine di calcolo: niente testi dell'utente
      } else {
        // affettatori, sequenze temporali, query, connessioni, controlli, dati ricchi, VML dei commenti...
        s = trasformaGenerico(testo(i), T);
      }
      // Testi fissi dei formati numerici personalizzati («#,##0 "rif. Rossi"»): negli stili, e copiati nei grafici
      // e nelle pivot. E chi ha aggiornato una pivot o modificato il file condiviso: e' il nome dell'utente.
      s = s.replace(/(\sformatCode=")([^"]*)"/g, function (a, pre, v) {
        var prima = xmlDec(v), dopo = letterali(prima, T);
        return dopo === prima ? a : pre + xmlEnc(dopo, true) + "\"";
      }).replace(/<((?:c:)?formatCode)>([\s\S]*?)<\/\1>/g, function (a, tag, x) {
        var prima = xmlDec(x), dopo = letterali(prima, T);
        return dopo === prima ? a : "<" + tag + ">" + xmlEnc(dopo) + "</" + tag + ">";
      }).replace(/(\s(?:refreshedBy|userName))="[^"]*"/g, "$1=\"\"");
      s = togliUtente(togliRichiami(p, s));
      // stampanti (il loro file binario puo' contenere il nome della stampante o dell'utente) e macro (possono
      // contenere nomi e non si possono leggere qui): si tolgono, con i riferimenti che puntano a loro
      // relazioni e tipi che nominano una parte tolta
      daTogliere.forEach(function (t) {
        var nome = escapeRe(t.split("/").pop());
        s = s.replace(new RegExp("<Relationship\\b[^>]*\\sTarget=\"(?:[^\"]*/)?" + nome + "\"[^>]*/>", "g"), "");
        if (p === "[Content_Types].xml") s = s.replace(new RegExp("<Override\\b[^>]*PartName=\"/" + escapeRe(t) + "\"[^>]*/>", "g"), "");
      });
      if (tolgoStampanti) s = s.replace(/(<pageSetup\b[^>]*?)\s+r:id="[^"]*"/g, "$1");
      if (macro) {
        if (p === "[Content_Types].xml") s = s.replace(/<Override\b[^>]*vbaProject[^>]*\/>/g, "")
          .replace(/application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml/g, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        if (p === "xl/_rels/workbook.xml.rels") s = s.replace(/<Relationship\b[^>]*vbaProject[^>]*\/>/g, "");
        if (p === "xl/_rels/vbaProject.bin.rels") { tolte.push(p); return; }
      }
      scrivi(i, s);
    });
    // cfb_del vuole il percorso con la barra iniziale: senza, non trova la parte e non toglie niente
    tolte.forEach(function (p) { if (!XLSX.CFB.utils.cfb_del(zip, "/" + p)) throw new Error("Non riesco a togliere " + p); });

    // Due grafie dello stesso nome («Mario Rossi», «Mario  Rossi») diventano lo stesso codice: dentro una pivot o tra
    // le colonne di una tabella due voci uguali rendono il file illeggibile per Excel. Allora niente file formattato.
    var doppi = [];
    zip.FullPaths.forEach(function (pp, i) {
      var p = percorsoParte(pp), f = zip.FileIndex[i];
      if (f.type !== 2 || !f.content) return;
      var gruppi = [];
      if (/^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(p)) {
        var s = testo(i);
        s.replace(/<sharedItems\b[^>]*>([\s\S]*?)<\/sharedItems>/g, function (a, x) {
          var voci = [];
          x.replace(/<s\b[^>]*\sv="([^"]*)"/g, function (b, v) { voci.push(v); return b; });
          gruppi.push(voci);
          return a;
        });
        var campi = [];
        s.replace(/<cacheField\b[^>]*\sname="([^"]*)"/g, function (a, v) { campi.push(v); return a; });
        gruppi.push(campi);
      } else if (/^xl\/tables\/[^/]*\.xml$/.test(p)) {
        var col = [];
        testo(i).replace(/<tableColumn\b[^>]*\sname="([^"]*)"/g, function (a, v) { col.push(v); return a; });
        gruppi.push(col);
      }
      gruppi.forEach(function (voci) {
        var viste = new Set();
        voci.forEach(function (v) {
          var k = xmlDec(v).toLowerCase();
          if (viste.has(k) && doppi.indexOf(p) < 0) doppi.push(p);
          viste.add(k);
        });
      });
    });
    if (doppi.length) return { buffer: null, motivo: "doppioni", parti: doppi };
    if (macro) avvisi.push("Le macro del file sono state tolte: potevano contenere nomi.");
    if (percorsi.some(function (p) { return /^xl\/externalLinks\//.test(p); }))
      avvisi.push("Il file ha collegamenti ad altri file Excel: i loro percorsi possono contenere nomi, controllali.");
    var immagini = percorsi.filter(function (p) { return PARTI_IMMAGINE.test(p); }).length;
    if (immagini) avvisi.push("Il file contiene " + immagini + (immagini === 1 ? " immagine" : " immagini") +
      ": Claude le vedrà così come sono. Controlla che non mostrino nomi, firme o loghi di clienti.");

    rep.nuoviCodici = esito.mappa.nuovi;
    var r = JSON.parse(salva);
    rep.regole = r.r;
    rep.occorrenzeTesto = r.o;
    var out = new Uint8Array(XLSX.CFB.write(zip, { fileType: "zip", type: "array", compression: true }));
    return { buffer: out, avvisi: avvisi, controllo: controllaPacchetto(out, esito.mappa) };
  }

  // Ripristino che tiene la formattazione: stessa macchina, con la trasformazione inversa (codice -> valore vero).
  // Per i formati non zip (.xls, .csv) si ricade sul ripristino dei soli valori.
  function ripristinaXlsx(buffer, righe) {
    if (!eZip(buffer)) { var v = ripristinaExcel(buffer, righe); v.formattazione = false; return v; }
    var mappa = mappaRipristino(righe), trovati = 0, sconosciuti = [], storpiati = [];
    function R(s) {
      var r = ripristinaTesto(s, mappa);
      trovati += r.trovati;
      r.sconosciuti.forEach(function (x) { if (sconosciuti.indexOf(x) < 0) sconosciuti.push(x); });
      r.storpiati.forEach(function (x) { if (storpiati.indexOf(x) < 0) storpiati.push(x); });
      return r.testo;
    }
    var nomi = XLSX.read(buffer, { type: "array", bookSheets: true }).SheetNames;
    var finto = {
      report: { regole: {}, occorrenzeTesto: 0 },
      mappa: { nuovi: 0, perCodice: new Map() },
      nomiFogli: nomi.map(function (n) { return nomeFoglioRipristinato(n, mappa); }),
      trasforma: { testo: R, numero: function () { return null; },
        numeroDaTesto: function (s) { var n = numeroRipristinato(s, mappa); if (n !== null) trovati++; return n; } }
      // (per le celle a testo condiviso il codice si conta due volte, qui e nel testo condiviso: il totale mostrato e'
      // indicativo, i «non trovati» sono esatti)
    };
    var res = anonimizzaXlsx(buffer, finto);
    if (!res.buffer) { var w = ripristinaExcel(buffer, righe); w.formattazione = false; return w; }
    return { buffer: res.buffer, trovati: trovati, sconosciuti: sconosciuti, storpiati: storpiati, formattazione: true };
  }

  // Controllo finale: nessun valore vero (delle corrispondenze) deve comparire nel testo di nessuna parte del
  // pacchetto, e nessun percorso con il nome dell'utente. Le parti che non sono testo non si possono controllare:
  // se ce ne sono (a parte le immagini) il controllo non passa. Riporta solo dove e quante volte, mai il valore.
  function controllaPacchetto(buffer, mappa) {
    var valori = [];
    mappa.perCodice.forEach(function (v) {
      if (norm(v.valore).length >= 3) valori.push(v.valore);
    });
    var alternative = [];
    valori.forEach(function (t) {
      formeRicerca(t).forEach(function (f) {
        // cifre e codici (telefoni, IBAN, P.IVA): anche con spazi o separatori in mezzo, e senza gli zeri iniziali
        // (una P.IVA salvata come numero li perde)
        var compatto = f.replace(/[\s.\/-]/g, ""), alt;
        if (/^[0-9a-z]{6,34}$/.test(compatto) && /\d{4}/.test(compatto)) {
          var zeri = /^\d+$/.test(compatto) ? /^0*/.exec(compatto)[0].length : 0;
          alt = (zeri ? "(?:0[\\s./-]?){0," + zeri + "}" : "") + compatto.slice(zeri).split("").map(escapeRe).join("[\\s./-]?");
        } else alt = escapeRe(f).replace(/ /g, "\\s+");
        if (alternative.indexOf(alt) < 0) alternative.push(alt);
      });
    });
    alternative.sort(function (a, b) { return b.length - a.length; });
    var re = alternative.length ? new RegExp(CONFINE_PRIMA + "(?:" + alternative.join("|") + ")" + CONFINE_DOPO, "giu") : null;
    var problemi = [];
    if (!eZip(buffer)) return { ok: true, problemi: problemi };
    var dec = new TextDecoder("utf-8");
    var zip = XLSX.CFB.read(buffer, { type: "array" });
    // quante volte compare un valore vero, nel testo cosi' com'e' e nella forma NFKC (legature, «´»)
    function conta(t) { return re ? Math.max((t.match(re) || []).length, (t.normalize("NFKC").match(re) || []).length) : 0; }
    zip.FullPaths.forEach(function (pp, i) {
      var p = percorsoParte(pp), f = zip.FileIndex[i];
      if (f.type !== 2 || !f.content || p === "" || /Sh33tJ5/.test(p)) return;
      if (!PARTI_TESTO.test(p)) {
        // le immagini EMF/WMF (un intervallo copiato come immagine) contengono il testo delle celle in UTF-16
        if (/\.(emf|wmf)$/i.test(p)) {
          var nImm = conta(new TextDecoder("utf-16le").decode(new Uint8Array(f.content)).replace(/\u0000/g, " ").normalize("NFC"));
          if (nImm) problemi.push({ parte: p, volte: nImm, tipo: "valore" });
        } else if (!PARTI_IMMAGINE.test(p)) problemi.push({ parte: p, volte: 1, tipo: "illeggibile" });
        return;
      }
      var grezzo = dec.decode(new Uint8Array(f.content));
      if (/^customXml\//.test(p) && /DataMashup/.test(grezzo)) { problemi.push({ parte: p, volte: 1, tipo: "illeggibile" }); return; }
      grezzo = grezzo
        .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, function (a, x) { return xmlEnc(x); })
        .replace(TAG_INTEST, function (a, tag, x) { return " " + xmlEnc(testoIntestazione(xmlDec(x))) + " "; });
      // 1. i tag diventano spazi, il testo degli attributi resta (nomi dei fogli, cache delle pivot, percorsi)
      //    Il separatore non e' uno spazio: due testi vicini («…Mario» in una cella, «Rossi…» nella successiva) non
      //    devono sembrare un nome intero. Un nome spezzato dentro lo stesso testo lo trova il punto 2.
      var s = xmlDec(grezzo.replace(/<[^>]*>/g, function (tag) {
        var attr = tag.match(/"[^"]*"|'[^']*'/g);
        return attr ? "\u0001" + attr.join("\u0001") + "\u0001" : "\u0001";
      })).normalize("NFC");
      // 2. dentro ogni testo formattato (testo condiviso, cella con testo proprio, paragrafo di una casella o di un
      //    commento) i tag spariscono: un nome spezzato tra due pezzi di formattazione («Mar» + «io Rossi») torna
      //    intero. Solo li': unire le celle di un foglio creerebbe numeri e parole che nel file non ci sono.
      var pezzi = [];
      grezzo.replace(/<(si|is|text|a:p)\b[^>]*>([\s\S]*?)<\/\1>/g, function (a, tag, x) { pezzi.push(x.replace(/<[^>]*>/g, "")); return a; });
      // separatore che non e' uno spazio: un testo che finisce con «Mario» e il successivo che comincia con «Rossi»
      // non devono sembrare il nome intero
      var unito = xmlDec(pezzi.join("\u0001")).normalize("NFC");
      var n = Math.max(conta(s), conta(unito));
      if (n) problemi.push({ parte: p, volte: n, tipo: "valore" });
      if (togliUtente(s) !== s) problemi.push({ parte: p, volte: 1, tipo: "percorso" });
    });
    return { ok: problemi.length === 0, problemi: problemi };
  }

  // Euristica per mettere in cima i testi che potrebbero contenere nomi: una parola con iniziale maiuscola che non
  // e' la prima del testo (le intestazioni di una parola sola non contano).
  function sembraNome(t) {
    // nome rimasto a meta': una parola maiuscola attaccata a un codice («Mario PERSONA_001», «PERSONA_001 Rossi»),
    // anche se e' la prima del testo
    var accanto = /(\p{Lu}\p{Ll}{2,})\s+[A-Z][A-Z0-9_]*_\d{3,}\b|\b[A-Z][A-Z0-9_]*_\d{3,}\s+(\p{Lu}\p{Ll}{2,})/gu, m;
    while ((m = accanto.exec(t)) !== null) {
      var w = m[1] || m[2];
      if (!PAROLE_COMUNI.test(w) && !SUFFISSI.test(w) && !CONNETTIVI.test(w)) return true;
    }
    var parole = t.replace(/\b[A-Z][A-Z0-9_]*_\d{3,}\b/g, " ").split(/[\s,;:()\/]+/).filter(Boolean);
    return parole.slice(1).some(function (p) { return /^\p{Lu}\p{Ll}{2,}/u.test(p) && !PAROLE_COMUNI.test(p) && !SUFFISSI.test(p); });
  }

  // ------------------------------------------------------------------ file di uscita
  function scrivi(wb) { return XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }); }

  function fileCorrispondenze(mappa) {
    var wb = XLSX.utils.book_new();
    wb.Props = { Title: "Corrispondenze", Author: "" };
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Codice", "Valore vero", "Origine"]].concat(mappa.righe())), "Corrispondenze");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ["ATTENZIONE: questo file contiene i dati veri."],
      ["Non mandarlo mai a Claude, a ChatGPT o a qualunque servizio esterno."],
      ["Serve solo per ripristinare i nomi veri e per dare gli stessi codici ai file successivi."]]), "Leggimi");
    return scrivi(wb);
  }

  function leggiCorrispondenze(buffer) {
    var wb = XLSX.read(buffer, { type: "array" });
    var ws = wb.Sheets["Corrispondenze"] || wb.Sheets[wb.SheetNames[0]];
    var righe = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
    return righe.slice(1);
  }

  // ------------------------------------------------------------------ ripristino
  var RE_CODICE = /\b[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)*_\d{3,}\b/g;
  // codice -> {valore, origine}; si costruisce una volta sola e si passa al posto delle righe
  function mappaRipristino(righe) {
    if (righe instanceof Map) return righe;
    return new Map(righe.map(function (r) { return [String(r[0]).trim(), { valore: String(r[1]), origine: r[2] == null ? "" : String(r[2]) }]; }));
  }
  function ripristinaTesto(testo, righe) {
    var m = mappaRipristino(righe);
    var trovati = 0, sconosciuti = [], storpiati = [];
    // Prima i codici conosciuti, anche se attaccati ad altro testo: un IBAN scritto attaccato ad altre cifre diventa
    // «IBAN_001» seguito da quelle cifre, e deve tornare com'era. Poi, quelli che sembrano codici ma non lo sono.
    var reNoti = codiciNoti(m);
    var out = reNoti ? String(testo).replace(reNoti, function (c) { trovati++; return m.get(c).valore; }) : String(testo);
    // (si cercano nel testo di partenza: un valore vero rimesso potrebbe somigliare a un codice)
    String(testo).replace(RE_CODICE, function (c) {
      if (reNoti) { reNoti.lastIndex = 0; var k = reNoti.exec(c); if (k && k.index === 0) return c; }
      if (sconosciuti.indexOf(c) < 0) sconosciuti.push(c);
      return c;
    });
    // Codici scritti in un altro modo dall'assistente («persona_001», «PERSONA_1», «Persona 001»): non si rimettono
    // da soli («la persona 1 del gruppo» puo' essere testo normale), ma si segnalano perche' l'utente li controlli.
    var reV = variantiCodici(m);
    if (reV) out.replace(reV, function (x, pref, num) {
      var giusto = pref.toUpperCase() + "_" + String(+num).padStart(3, "0");
      if (x !== giusto && m.has(giusto) && storpiati.indexOf(x) < 0) storpiati.push(x);
      return x;
    });
    return { testo: out, trovati: trovati, sconosciuti: sconosciuti, storpiati: storpiati };
  }
  function codiciNoti(m) {
    if (m._noti !== undefined) return m._noti;
    var codici = Array.from(m.keys()).sort(function (a, b) { return b.length - a.length; });
    m._noti = codici.length ? new RegExp("(?<![A-Za-z0-9])(?:" + codici.map(escapeRe).join("|") + ")", "g") : null;
    return m._noti;
  }
  function variantiCodici(m) {
    if (m._varianti !== undefined) return m._varianti;
    var pref = [];
    m.forEach(function (v, c) { var p = /^(.*)_\d+$/.exec(c); if (p && pref.indexOf(p[1]) < 0) pref.push(p[1]); });
    m._varianti = pref.length ? new RegExp("(?<![\\p{L}\\p{N}_])(" + pref.map(escapeRe).join("|") + ")[ _-]?(\\d{1,6})(?![\\p{L}\\p{N}_])", "giu") : null;
    return m._varianti;
  }
  // Una cella che contiene solo il codice di una P.IVA che nell'originale era un numero torna un numero.
  function numeroRipristinato(s, m) {
    var x = m.get(String(s).trim());
    return x && x.origine === "regola piva numero" && /^\d{11}$/.test(x.valore) ? +x.valore : null;
  }
  function nomeFoglioRipristinato(n, m) { return ripristinaTesto(n, m).testo.replace(/[\[\]:*?\/\\]/g, "_").slice(0, 31); }

  function ripristinaExcel(buffer, righe) {
    var m = mappaRipristino(righe);
    var wb = XLSX.read(buffer, { type: "array", cellDates: true });
    var tot = 0, sconosciuti = [], storpiati = [];
    function nota(r) {
      r.sconosciuti.forEach(function (s) { if (sconosciuti.indexOf(s) < 0) sconosciuti.push(s); });
      r.storpiati.forEach(function (s) { if (storpiati.indexOf(s) < 0) storpiati.push(s); });
    }
    wb.SheetNames.forEach(function (nome) {
      var ws = wb.Sheets[nome];
      Object.keys(ws).forEach(function (k) {
        if (k[0] === "!") return;
        var c = ws[k];
        if (c.t !== "s" || typeof c.v !== "string") return;
        var n = numeroRipristinato(c.v, m);
        if (n !== null) { c.t = "n"; c.v = n; delete c.w; delete c.h; delete c.r; tot++; return; }
        var r = ripristinaTesto(c.v, m);
        if (r.trovati) { c.v = r.testo; delete c.w; delete c.h; delete c.r; tot += r.trovati; }
        nota(r);
      });
    });
    // nomi dei fogli (senza doppioni: Excel non li accetta)
    var nuovi = {}, usati = {};
    wb.SheetNames.forEach(function (nome) {
      var r = ripristinaTesto(nome, m);
      nota(r);
      var n = r.trovati ? nomeFoglioRipristinato(nome, m) : nome, base = n, i = 2;
      while (usati[n.toLowerCase()]) n = base.slice(0, 28) + "_" + (i++);
      usati[n.toLowerCase()] = true;
      if (n !== nome) { nuovi[nome] = n; tot += r.trovati; }
    });
    wb.SheetNames = wb.SheetNames.map(function (nome) {
      if (!nuovi[nome]) return nome;
      wb.Sheets[nuovi[nome]] = wb.Sheets[nome];
      delete wb.Sheets[nome];
      return nuovi[nome];
    });
    return { buffer: scrivi(wb), trovati: tot, sconosciuti: sconosciuti, storpiati: storpiati };
  }

  root.AnonCore = {
    init: function (xlsx) { XLSX = xlsx; },
    leggi: leggi, colonne: colonne, testiDistinti: testiDistinti, anonimizza: anonimizza, scrivi: scrivi,
    rigaIntestazione: rigaIntestazione, proposteDaCorrispondenze: proposteDaCorrispondenze,
    anonimizzaXlsx: anonimizzaXlsx, controllaPacchetto: controllaPacchetto, ripristinaXlsx: ripristinaXlsx, eZip: eZip,
    ricerca: costruisciRicerca, norm: norm, prefissoDa: prefissoDa,
    fileCorrispondenze: fileCorrispondenze, leggiCorrispondenze: leggiCorrispondenze,
    ripristinaTesto: ripristinaTesto, ripristinaExcel: ripristinaExcel, mappaRipristino: mappaRipristino,
    _test: { formaDiNome: formaDiNome, pivaValida: pivaValida, cfValido: cfValido, ibanPrefissoValido: ibanPrefissoValido, norm: norm,
      prefissoDa: prefissoDa, sembraNome: sembraNome, trasformaIntestazione: trasformaIntestazione, testoIntestazione: testoIntestazione,
      togliUtente: togliUtente, valoreInTesto: valoreInTesto }
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
