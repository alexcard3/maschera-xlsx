// Collaudo dei posti del pacchetto .xlsx dove un nome si puo' nascondere, su file costruiti qui con dati INVENTATI.
// Non serve Excel: si parte da un file scritto da SheetJS e si aggiungono a mano le parti che Excel scriverebbe.
// Uso: node collaudo_pacchetto.js      (esce con codice 1 se un controllo fallisce)
"use strict";
const XLSX = require("../vendor/xlsx.full.min.js");
require("../src/anonimizza_core.js");
const A = globalThis.AnonCore;
A.init(XLSX);

let falliti = 0;
function controlla(nome, ok, dettaglio) {
  console.log((ok ? "PASS " : "FAIL ") + nome + (dettaglio ? " — " + dettaglio : ""));
  if (!ok) falliti++;
}
const dec = new TextDecoder(), enc = new TextEncoder();
const NOMI = ["Mario Rossi", "Müller", "Nicol\u00f2 Ferr\u00e8", "Rossi & Figli", "Sara Moretti"];

// file di base: un foglio con un nome che non e' solo ASCII, uno di riepilogo con formule che lo citano
function base() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Cliente", "P.IVA", "Importo"],
    ["Mario Rossi", 4321768915, 1234567897],     // P.IVA salvata come numero; importo che passa il controllo per caso
    ["Rossi & Figli", null, 500],
    ["Telefonata con Nicolo\u0300 Ferre\u0300", null, 0]]), "Müller");
  const r = XLSX.utils.aoa_to_sheet([["Totale"]]);
  r["A2"] = { t: "n", v: 0, f: "Müller!C2+Müller!C3" };
  r["!ref"] = "A1:A2";
  XLSX.utils.book_append_sheet(wb, r, "Riepilogo");
  return XLSX.CFB.read(new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" })), { type: "array" });
}
function parte(zip, nome) { const i = zip.FullPaths.findIndex((p) => p.replace(/^Root Entry\//, "") === nome); return i; }
function leggiParte(zip, nome) { const i = parte(zip, nome); return i < 0 ? null : dec.decode(zip.FileIndex[i].content); }
function scriviParte(zip, nome, testo) {
  const b = typeof testo === "string" ? enc.encode(testo) : testo, i = parte(zip, nome);
  if (i < 0) XLSX.CFB.utils.cfb_add(zip, nome, b);
  else { zip.FileIndex[i].content = b; zip.FileIndex[i].size = b.length; }
}
function cambia(zip, nome, f) { scriviParte(zip, nome, f(leggiParte(zip, nome))); }
function buffer(zip) { return new Uint8Array(XLSX.CFB.write(zip, { fileType: "zip", type: "array" })); }
function anonimizza(buf) {
  const lettura = A.leggi(buf);
  const valori = NOMI.map((t) => ({ testo: t, categoria: "PERSONA" }));
  const esito = A.anonimizza(lettura, { fogli: {}, termini: "", valori, ovunque: true, rinominaFogli: false,
    regole: { iban: true, cf: true, email: true, piva: true, tel: true } }, null);
  return { lettura, esito, res: A.anonimizzaXlsx(buf, esito) };
}
// ogni parte di testo dell'uscita, con i riferimenti XML risolti e in NFC
function testi(buf) {
  const zip = XLSX.CFB.read(buf, { type: "array" }), out = {};
  zip.FullPaths.forEach((p, i) => {
    const f = zip.FileIndex[i];
    if (f.type === 2 && f.content && /\.(xml|rels|vml)$/.test(p)) {
      out[p.replace(/^Root Entry\//, "")] = dec.decode(f.content).replace(/&amp;/g, "&").replace(/&#10;/g, "\n").normalize("NFC");
    }
  });
  return out;
}
function cerca(buf, ago) {
  const t = testi(buf);
  return Object.keys(t).filter((p) => t[p].toLowerCase().includes(ago.toLowerCase()));
}

// ------------------------------------------------------------ 1. file con tutti i nascondigli testuali
const zip = base();
cambia(zip, "xl/workbook.xml", (s) => s
  .replace(/<workbook([^>]*)>/, '<workbook$1 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac">' +
    '<fileSharing readOnlyRecommended="1" userName="Sara Moretti"/>' +
    '<mc:AlternateContent><mc:Choice Requires="x15ac"><x15ac:absPath url="C:\\Users\\mrossi\\Documenti\\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent>'));
cambia(zip, "xl/worksheets/sheet1.xml", (s) => s.replace("</worksheet>",
  '<dataValidations count="1"><dataValidation type="list" sqref="D2" prompt="Chiedere a Mario Rossi" promptTitle="Nota"><formula1>"Mario Rossi,Altro"</formula1></dataValidation></dataValidations>' +
  '<headerFooter><oddHeader>&amp;LMario Rossi&amp;C&amp;"Arial,Bold"&amp;12Rossi &amp;&amp; Figli&amp;RPagina &amp;P</oddHeader></headerFooter></worksheet>'));
cambia(zip, "xl/styles.xml", (s) => s
  .replace(/<numFmts count="\d+">/, (m) => m + '<numFmt numFmtId="200" formatCode="#,##0 &quot;rif. Mario Rossi&quot;"/>')
  .replace(/<\/styleSheet>/, "</styleSheet>")
  .replace(/<cellStyles count="(\d+)">/, (m, n) => '<cellStyles count="' + (+n + 1) + '"><cellStyle name="Stile Mario Rossi" xfId="0" customBuiltin="1"/>'));
if (!/numFmtId="200"/.test(leggiParte(zip, "xl/styles.xml")))
  cambia(zip, "xl/styles.xml", (s) => s.replace(/(<styleSheet[^>]*>)/, '$1<numFmts count="1"><numFmt numFmtId="200" formatCode="#,##0 &quot;rif. Mario Rossi&quot;"/></numFmts>'));
// parte senza una regola sua (affettatore): nome in un attributo, con un a capo che deve restare codificato
scriviParte(zip, "xl/slicerCaches/slicerCache1.xml",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Filtro_Mario Rossi" sourceName="Voce&#10;conto"><data><tabular><items><i x="0"/></items></tabular></data></slicerCacheDefinition>');
// connessione a un file nella cartella dell'utente
scriviParte(zip, "xl/connections.xml",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" name="Query" odcFile="C:\\Users\\mrossi\\Documents\\Origini dati\\q.odc" type="5"/></connections>');
// commento classico: autore e riga in grassetto con il suo nome
scriviParte(zip, "xl/comments1.xml",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Laura Bassi</author></authors><commentList><comment ref="C2" authorId="0"><text><r><rPr><b/></rPr><t>Laura Bassi:</t></r><r><t xml:space="preserve">\nimporto da verificare</t></r></text></comment></commentList></comments>');
// commento moderno con menzione
scriviParte(zip, "xl/threadedComments/threadedComment1.xml",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A2" dT="2025-01-01T10:00:00.00" personId="{1}" id="{2}"><text>@Mario Rossi guarda qui</text><mentions><mention mentionpersonId="{1}" mentionId="{3}" startIndex="0" length="12"/></mentions></threadedComment></ThreadedComments>');
const buf1 = buffer(zip);
const { lettura, esito, res } = anonimizza(buf1);
controlla("pacchetto: uscita formattata prodotta e controllo finale pulito", !!res.buffer && res.controllo.ok,
  res.motivo || JSON.stringify(res.controllo && res.controllo.problemi));
if (res.buffer) {
  const t = testi(res.buffer);
  controlla("percorso della cartella di lavoro tolto", !/absPath/.test(t["xl/workbook.xml"]) && cerca(res.buffer, "mrossi").length === 0,
    cerca(res.buffer, "mrossi").join(","));
  controlla("nome di chi ha protetto il file tolto", !/Sara Moretti/.test(t["xl/workbook.xml"]));
  controlla("intestazione: codici al loro posto, nomi tolti", /&L\S*_\d{3}&C&"Arial,Bold"&12\S*_\d{3}&RPagina &P/.test(t["xl/worksheets/sheet1.xml"]),
    (/<oddHeader>([^<]*)/.exec(t["xl/worksheets/sheet1.xml"]) || [])[1]);
  controlla("messaggio della convalida anonimizzato", !/Chiedere a Mario Rossi/.test(t["xl/worksheets/sheet1.xml"]));
  controlla("formato numerico e stile di cella anonimizzati", !/Mario Rossi/.test(t["xl/styles.xml"]));
  controlla("parte senza regola sua (affettatore) anonimizzata", !/Mario Rossi/.test(t["xl/slicerCaches/slicerCache1.xml"]));
  controlla("a capo in un attributo resta codificato", /sourceName="Voce&#10;conto"/.test(leggiParte(XLSX.CFB.read(res.buffer, { type: "array" }), "xl/slicerCaches/slicerCache1.xml")));
  controlla("connessione: cartella dell'utente tolta", /Users\\utente\\/.test(t["xl/connections.xml"]));
  controlla("commento: autore e riga in grassetto senza il nome", !/Laura Bassi/.test(t["xl/comments1.xml"]));
  controlla("commento moderno: menzione tolta, nome nel testo sostituito", !/<mentions>/.test(t["xl/threadedComments/threadedComment1.xml"]) &&
    !/Mario Rossi/.test(t["xl/threadedComments/threadedComment1.xml"]));
  controlla("accento scritto separato (NFD) trovato", cerca(res.buffer, "Nicol\u00f2").length === 0, cerca(res.buffer, "Nicol\u00f2").join(","));
  const fogli = XLSX.read(res.buffer, { type: "array", cellFormula: true });
  const nuovo = fogli.SheetNames[0], f = fogli.Sheets[fogli.SheetNames[1]]["A2"].f;
  controlla("foglio con nome non ASCII rinominato anche nelle formule senza apici", /^PERSONA_\d{3}$/.test(nuovo) && f === nuovo + "!C2+" + nuovo + "!C3", nuovo + " | " + f);
  const ws = fogli.Sheets[nuovo];
  controlla("P.IVA come numero sotto l'intestazione «P.IVA»: diventa codice", /^PIVA_\d{3}$/.test(String(ws["B2"].v)), String(ws["B2"].v));
  controlla("importo che sembra una P.IVA: resta numero", ws["C2"].t === "n" && ws["C2"].v === 1234567897, ws["C2"].t + " " + ws["C2"].v);
  // testi fuori dalle celle: visibili all'utente prima di scaricare (i commenti si provano nel collaudo con Excel:
  // qui il commento non e' collegato al foglio e SheetJS non lo legge)
  const dove = lettura.extra.map((x) => x.dove).join(" | ");
  controlla("testi fuori dalle celle elencati (intestazione, convalida)", /intestazione/.test(dove) && /convalida/.test(dove), dove);
  const parti = lettura.extra.filter((x) => /intestazione/.test(x.dove)).map((x) => x.testo);
  controlla("intestazione: una voce per parte (sinistra, centro, destra), come la trasforma il file",
    parti.includes("Mario Rossi") && parti.includes("Rossi & Figli") && parti.includes("Pagina"), parti.join(" | "));
  // ripristino: la P.IVA numerica torna numero, il foglio torna «Müller»
  const rip = A.ripristinaXlsx(res.buffer, esito.mappa.righe());
  const back = XLSX.read(rip.buffer, { type: "array" });
  controlla("ripristino formattato: foglio e P.IVA numerica", back.SheetNames[0] === "Müller" && back.Sheets["Müller"]["B2"].t === "n" &&
    back.Sheets["Müller"]["B2"].v === 4321768915, back.SheetNames[0] + " " + (back.Sheets[back.SheetNames[0]]["B2"] || {}).t);
  const intest = /<oddHeader>([^<]*)/.exec(testi(rip.buffer)["xl/worksheets/sheet1.xml"])[1];
  controlla("ripristino: la & del nome nell'intestazione torna doppia", intest.includes("Rossi && Figli") && intest.startsWith("&LMario Rossi"), intest);
}
// il controllo finale deve mordere: sul file originale trova i nomi e il percorso
const morso = A.controllaPacchetto(buf1, esito.mappa);
const tipi = morso.problemi.map((p) => p.parte + ":" + p.tipo).join(", ");
controlla("il controllo finale scatta sul file originale (stili, intestazione, percorso)",
  !morso.ok && /styles\.xml:valore/.test(tipi) && /sheet1\.xml:valore/.test(tipi) && /workbook\.xml:percorso/.test(tipi), tipi);

// ------------------------------------------------------------ 2. parti che non si possono leggere
function conParte(nome, contenuto) { const z = base(); scriviParte(z, nome, contenuto); return buffer(z); }
const modello = anonimizza(conParte("xl/model/item.data", new Uint8Array([1, 2, 3, 77, 97, 114, 105, 111]))).res;
controlla("modello dati di Power Pivot: niente file formattato", !modello.buffer && modello.motivo === "parti" && modello.parti.includes("xl/model/item.data"),
  modello.motivo + " " + (modello.parti || []).join(","));
const incorporato = anonimizza(conParte("xl/embeddings/oleObject1.bin", new Uint8Array([9, 9, 9]))).res;
controlla("oggetto incorporato: niente file formattato", !incorporato.buffer && incorporato.motivo === "parti");
const query = anonimizza(conParte("customXml/item1.xml", '<?xml version="1.0"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">AAAA</DataMashup>')).res;
controlla("query di Power Query: niente file formattato", !query.buffer && query.motivo === "parti");
controlla("il controllo finale segnala le parti illeggibili",
  !A.controllaPacchetto(conParte("xl/model/item.data", new Uint8Array([1])), { perCodice: new Map() }).ok);

// ------------------------------------------------------------ 3. stampante: si toglie con i suoi riferimenti
(function () {
  const z = base();
  scriviParte(z, "xl/printerSettings/printerSettings1.bin", new Uint8Array([80, 82, 78]));
  cambia(z, "xl/worksheets/sheet1.xml", (s) => s.replace("</worksheet>", '<pageSetup paperSize="9" r:id="rId9"/></worksheet>'));
  const rels = leggiParte(z, "xl/worksheets/_rels/sheet1.xml.rels") ||
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  scriviParte(z, "xl/worksheets/_rels/sheet1.xml.rels", rels.replace("</Relationships>",
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/></Relationships>'));
  const r = anonimizza(buffer(z)).res;
  const t = r.buffer ? testi(r.buffer) : {};
  const parti = r.buffer ? XLSX.CFB.read(r.buffer, { type: "array" }).FullPaths.join("|") : "";
  controlla("stampante tolta, con relazione e riferimento", !!r.buffer && !/printerSettings/.test(parti) &&
    !/printerSettings/.test(t["xl/worksheets/_rels/sheet1.xml.rels"]) && !/r:id="rId9"/.test(t["xl/worksheets/sheet1.xml"]),
    r.motivo || parti);
})();

// ------------------------------------------------------------ 4. macro: si tolgono davvero
(function () {
  const z = base();
  scriviParte(z, "xl/vbaProject.bin", new Uint8Array([1, 2, 3]));
  cambia(z, "xl/_rels/workbook.xml.rels", (s) => s.replace("</Relationships>",
    '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
  cambia(z, "[Content_Types].xml", (s) => s.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml").replace("</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'));
  const r = anonimizza(buffer(z)).res;
  const t = r.buffer ? testi(r.buffer) : {};
  const parti = r.buffer ? XLSX.CFB.read(r.buffer, { type: "array" }).FullPaths.join("|") : "";
  controlla("macro tolte davvero, con relazione e tipo di file", !!r.buffer && !/vbaProject/.test(parti) &&
    !/vbaProject/.test(t["xl/_rels/workbook.xml.rels"]) && !/sheet\.macroEnabled\.main\+xml/.test(t["[Content_Types].xml"]) &&
    r.avvisi.some((a) => /macro/.test(a)), r.motivo || parti);
})();

// ------------------------------------------------------------ 5. secondo giro di revisione
function libro(fogli, opzioni) {
  const wb = XLSX.utils.book_new();
  Object.keys(fogli).forEach((n) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fogli[n]), n));
  return XLSX.CFB.read(new Uint8Array(XLSX.write(wb, Object.assign({ type: "array", bookType: "xlsx" }, opzioni || {}))), { type: "array" });
}
function anonimizzaCon(buf, nomi, altro) {
  const lettura = A.leggi(buf);
  const esito = A.anonimizza(lettura, Object.assign({ fogli: {}, termini: "", ovunque: true, rinominaFogli: false,
    valori: nomi.map((t) => ({ testo: t, categoria: "PERSONA" })), regole: { iban: true, cf: true, email: true, piva: true, tel: true } }, altro || {}), null);
  return { lettura, esito, res: A.anonimizzaXlsx(buf, esito) };
}

// a. caratteri che NFC e NFKC scrivono diversi: accento acuto come apostrofo, legatura da PDF, İ turca
(function () {
  const nomi = ["Luca D´Amico", "Raﬀaele Neri", "İlkay Demir"];
  const { res } = anonimizzaCon(buffer(libro({ F: [["Cliente"]].concat(nomi.map((n) => ["Telefonata con " + n])) })), nomi);
  const resti = nomi.filter((n) => cerca(res.buffer, n.normalize("NFC")).length);
  controlla("nomi con ´, legatura e İ: sostituiti", !!res.buffer && res.controllo.ok && resti.length === 0, resti.join(", ") || res.motivo);
  // il controllo li riconosce anche se nel file sono scritti nell'altra forma («Raffaele» con due f)
  const c = A.controllaPacchetto(buffer(libro({ F: [["Raffaele Neri"]] })), { perCodice: new Map([["PERSONA_002", { valore: "Raﬀaele Neri" }]]) });
  controlla("controllo finale: legatura riconosciuta nell'altra forma", !c.ok);
})();

// b. P.IVA come numero: etichetta che resta anche dopo un testo nella colonna; due zeri iniziali persi
(function () {
  const z = libro({ F: [["Fornitore", "P.IVA"], ["Alfa", 4321768915], ["Beta", "estero"], ["Gamma", 488410010], ["Delta", 1234567897]] });
  const { esito } = anonimizzaCon(buffer(z), []);
  const col = esito.righe[0].slice(1).map((r) => r[1]);
  controlla("P.IVA dopo «estero» e con due zeri persi: codici", /^PIVA_/.test(col[0]) && col[1] === "estero" && /^PIVA_/.test(col[2]) &&
    /^PIVA_/.test(col[3]), col.join(" | "));
})();

// c. nomi degli utenti di Office: raccolti per la marcatura automatica, poi sostituiti; percorso Mac tolto
(function () {
  const z = base();
  scriviParte(z, "xl/revisions/userNames.xml", '<?xml version="1.0"?><users xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><userInfo guid="{A}" name="Giulia Ferri" id="1" dateTime="2025-01-01T00:00:00"/></users>');
  cambia(z, "xl/workbook.xml", (s) => s.replace("</workbook>", '<customWorkbookViews><customWorkbookView name="Piero Sala - Visualizzazione personale" guid="{B}" windowWidth="100" windowHeight="100" activeSheetId="1"/></customWorkbookViews></workbook>'));
  cambia(z, "xl/worksheets/sheet1.xml", (s) => s.replace("</sheetData>", '</sheetData><scenarios><scenario name="Base" count="1" user="Carla Monti" comment="Creato da Carla Monti"><inputCells r="C2" val="1"/></scenario></scenarios>'));
  scriviParte(z, "customXml/item1.xml", '<?xml version="1.0"?><p:properties xmlns:p="http://schemas.microsoft.com/office/2006/metadata/properties"><documentManagement><Responsabile><UserInfo><DisplayName>Ugo Testa</DisplayName></UserInfo></Responsabile><Fonte>file:///Users/ugotesta/Documents/x.xlsx</Fonte></documentManagement></p:properties>');
  const buf = buffer(z), l = A.leggi(buf);
  const attesi = ["Giulia Ferri", "Piero Sala", "Carla Monti", "Ugo Testa"];
  controlla("utenti di Office raccolti per la marcatura", attesi.every((u) => l.autori.includes(u)), l.autori.join(", "));
  const { res } = anonimizzaCon(buf, NOMI.concat(l.autori));
  const resti = attesi.concat(["ugotesta"]).filter((u) => res.buffer && cerca(res.buffer, u).length);
  controlla("utenti di Office e percorso Mac tolti dal file", !!res.buffer && res.controllo.ok && resti.length === 0, resti.join(", ") || res.motivo);
})();

// d. casella di testo con il nome spezzato tra due pezzi di formattazione
(function () {
  const z = base();
  scriviParte(z, "xl/drawings/drawing1.xml", '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:sp><xdr:txBody><a:p><a:r><a:t>Visto da Mar</a:t></a:r><a:r><a:rPr b="1"/><a:t>io Rossi</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:wsDr>');
  const buf = buffer(z);
  const { res, esito } = anonimizzaCon(buf, NOMI);
  controlla("nome spezzato in una casella di testo: sostituito", !!res.buffer && res.controllo.ok && cerca(res.buffer, "io Rossi").length === 0);
  controlla("controllo finale: nome spezzato riconosciuto sull'originale", A.controllaPacchetto(buf, esito.mappa).problemi.some((p) => /drawing1/.test(p.parte)));
})();

// e. attributi tra apici singoli e sezioni CDATA in una parte senza regola sua
(function () {
  const z = base();
  scriviParte(z, "xl/webextensions/webextension1.xml", "<?xml version=\"1.0\"?><we:webextension xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\"><we:property name='autore' value='Mario Rossi'/><we:note><![CDATA[Nota per Mario Rossi <urgente>]]></we:note></we:webextension>");
  const buf = buffer(z);
  const { res, esito } = anonimizzaCon(buf, NOMI);
  controlla("apici singoli e CDATA: sostituiti", !!res.buffer && res.controllo.ok && cerca(res.buffer, "Mario Rossi").length === 0, res.motivo);
  controlla("controllo finale: apici singoli e CDATA visti sull'originale", A.controllaPacchetto(buf, esito.mappa).problemi.some((p) => /webextension/.test(p.parte)));
})();

// f. due grafie dello stesso nome in una pivot: niente file formattato (Excel non lo aprirebbe)
(function () {
  const z = base();
  scriviParte(z, "xl/pivotCache/pivotCacheDefinition1.xml", '<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheFields count="1"><cacheField name="Cliente"><sharedItems count="2"><s v="Mario Rossi"/><s v="Mario  Rossi"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>');
  const { res } = anonimizzaCon(buffer(z), NOMI);
  controlla("doppioni in una pivot: niente file formattato", !res.buffer && res.motivo === "doppioni", res.motivo);
})();

// g. fogli rinominati in Foglio1, Foglio2: le formule puntano ancora al foglio giusto
(function () {
  const wb = XLSX.utils.book_new();
  const r = XLSX.utils.aoa_to_sheet([["x"]]); r["A2"] = { t: "n", v: 0, f: "Foglio1!A1" }; r["!ref"] = "A1:A2";
  const f1 = XLSX.utils.aoa_to_sheet([["y"]]); f1["A2"] = { t: "n", v: 0, f: "Riepilogo!A1" }; f1["!ref"] = "A1:A2";
  XLSX.utils.book_append_sheet(wb, r, "Riepilogo");
  XLSX.utils.book_append_sheet(wb, f1, "Foglio1");
  const buf = new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
  const { res } = anonimizzaCon(buf, [], { rinominaFogli: true });
  const u = XLSX.read(res.buffer, { type: "array", cellFormula: true });
  const f = [u.Sheets[u.SheetNames[0]]["A2"].f, u.Sheets[u.SheetNames[1]]["A2"].f];
  controlla("rinomina dei fogli in un passaggio solo: formule sul foglio giusto",
    u.SheetNames.join(",") === "Foglio1,Foglio2" && f[0] === "Foglio2!A1" && f[1] === "Foglio1!A1", u.SheetNames.join(",") + " | " + f.join(" | "));
})();

// h. piè di pagina «Pagina &P di &N»: nessuna proposta di nome
(function () {
  const z = base();
  cambia(z, "xl/worksheets/sheet1.xml", (s) => s.replace("</worksheet>", "<headerFooter><oddFooter>Pagina &amp;P di &amp;N</oddFooter><oddHeader>&amp;CRiservato</oddHeader></headerFooter></worksheet>"));
  const l = A.leggi(buffer(z));
  const proposti = A.testiDistinti(l).filter((t) => t.proposto).map((t) => t.testo);
  controlla("piè di pagina e «Riservato»: nessuna proposta", !proposti.some((p) => /Pagina|Riservato|^di$/.test(p)), proposti.join(", "));
})();

// i. ripristino: P.IVA numerica tornata testo condiviso (file salvato da Excel o da Claude) torna numero
(function () {
  const { res, esito } = anonimizzaCon(buffer(base()), NOMI);
  const condiviso = new Uint8Array(XLSX.write(XLSX.read(res.buffer, { type: "array" }), { type: "array", bookType: "xlsx", bookSST: true }));
  const testoSs = dec.decode(XLSX.CFB.read(condiviso, { type: "array" }).FileIndex[parte(XLSX.CFB.read(condiviso, { type: "array" }), "xl/sharedStrings.xml")].content);
  const rip = A.ripristinaXlsx(condiviso, esito.mappa.righe());
  const b2 = XLSX.read(rip.buffer, { type: "array" }).Sheets["Müller"]["B2"];
  controlla("ripristino da testo condiviso: P.IVA torna numero", /PIVA_\d{3}/.test(testoSs) && b2.t === "n" && b2.v === 4321768915, b2.t + " " + b2.v);
})();

// j. autore «tc={...}» del commento classico legato a un commento moderno: resta com'e'
(function () {
  const z = base();
  scriviParte(z, "xl/comments1.xml", '<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>tc={11111111-2222-3333-4444-555555555555}</author></authors><commentList><comment ref="C2" authorId="0"><text><t>[Commento in thread]</t></text></comment></commentList></comments>');
  const { res } = anonimizzaCon(buffer(z), NOMI);
  controlla("autore tc={...} lasciato com'e'", !!res.buffer && /<author>tc=\{11111111/.test(testi(res.buffer)["xl/comments1.xml"]));
})();

// k. immagine EMF con il testo di un intervallo copiato: il controllo lo legge
(function () {
  const z = base();
  const testoEmf = "\u0001\u0000Mario Rossi\u0000", b = new Uint8Array(testoEmf.length * 2);
  for (let i = 0; i < testoEmf.length; i++) { b[2 * i] = testoEmf.charCodeAt(i) & 255; b[2 * i + 1] = testoEmf.charCodeAt(i) >> 8; }
  scriviParte(z, "xl/media/image1.emf", b);
  const { res } = anonimizzaCon(buffer(z), NOMI);
  controlla("immagine EMF con un nome: il controllo la blocca", !!res.buffer && !res.controllo.ok && res.controllo.problemi.some((p) => /image1\.emf/.test(p.parte)));
})();

// k2. ritorno da un LLM: grafico scritto senza prefisso «c:» (come fa openpyxl), codici scritti male segnalati
(function () {
  const { res, esito } = anonimizzaCon(buffer(base()), NOMI);
  const z = XLSX.CFB.read(res.buffer, { type: "array" });
  const cod = esito.mappa.righe().find((r) => r[1] === "Mario Rossi")[0];
  scriviParte(z, "xl/charts/chart9.xml", '<?xml version="1.0"?><chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"><chart><plotArea><barChart><ser><cat><strRef><f>Foglio!A1</f><strCache><ptCount val="1"/><pt idx="0"><v>' + cod + '</v></pt></strCache></strRef></cat></ser></barChart></plotArea></chart></chartSpace>');
  const rip = A.ripristinaXlsx(buffer(z), esito.mappa.righe());
  controlla("ritorno: grafico senza prefisso, nomi rimessi", testi(rip.buffer)["xl/charts/chart9.xml"].includes("<v>Mario Rossi</v>"));
  const t = A.ripristinaTesto("Sentire " + cod.toLowerCase() + ", " + cod.replace(/_0+/, "_") + " e la persona 7 del gruppo", esito.mappa.righe());
  controlla("ritorno: codici scritti male segnalati, non rimessi", t.storpiati.length === 2 && t.testo.includes(cod.toLowerCase()), t.storpiati.join(", "));
})();

// k3. file esportato da Google Fogli: metadati tolti con la relazione e con l'elemento che la richiama
(function () {
  const z = base();
  scriviParte(z, "xl/metadata", new Uint8Array([49, 46, 48, 52, 32, 77, 97, 114, 105, 111]));
  cambia(z, "xl/_rels/workbook.xml.rels", (s) => s.replace("</Relationships>", '<Relationship Id="rId77" Type="http://customschemas.google.com/relationships/workbookmetadata" Target="metadata"/></Relationships>'));
  cambia(z, "[Content_Types].xml", (s) => s.replace("</Types>", '<Override ContentType="application/binary" PartName="/xl/metadata"/></Types>'));
  cambia(z, "xl/workbook.xml", (s) => s.replace("</workbook>", '<extLst><ext uri="GoogleSheetsCustomDataVersion2"><go:sheetsCustomData xmlns:go="http://customooxmlschemas.google.com/" r:id="rId77" roundtripDataChecksum="x"/></ext></extLst></workbook>'));
  const { res } = anonimizzaCon(buffer(z), NOMI);
  const t = res.buffer ? testi(res.buffer) : {};
  controlla("Google Fogli: metadati, relazione e richiamo tolti", !!res.buffer && res.controllo.ok && !("xl/metadata" in t) &&
    !/rId77/.test(t["xl/workbook.xml"]) && !/metadata"/.test(t["xl/_rels/workbook.xml.rels"]) && !/\/xl\/metadata"/.test(t["[Content_Types].xml"]),
    res.motivo || "");
})();

// l. nome rimasto a meta': segnalato
controlla("«Mario PERSONA_001» segnalato come possibile nome", A._test.sembraNome("Mario PERSONA_001") && A._test.sembraNome("PERSONA_001 Rossi"));
controlla("«Riunione con PERSONA_013» non segnalato (niente rumore)", !A._test.sembraNome("Riunione con PERSONA_013 e PERSONA_006") &&
  !A._test.sembraNome("Fattura PERSONA_001"));

console.log("\n" + (falliti ? falliti + " controlli FALLITI" : "tutti i controlli superati"));
process.exit(falliti ? 1 : 0);
