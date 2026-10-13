# Experiment: rizzo-pii as a name-suggestion engine for Maschera — prediction · Esperimento: rizzo-pii come motore di suggerimenti per Maschera — predizione

**English** · [Italiano](#italiano)

Written on 2026-10-13 **before** running the measurement, and committed before it: the git history shows the order.
It is not edited after the results; results go to `RISULTATI.md`.

## Question

Maschera proposes likely names with simple rules (words with capital initials, not common accounting words, at most
twice per sheet). rizzo-pii is an Italian PII model that, on a spreadsheet pasted as text, missed half of the
surnames (test of 2026-10-12). Can the two together, rizzo-pii suggesting and the user confirming in Maschera, find
clearly more names than Maschera alone, without flooding the user with false suggestions?

## Set-up

- `banco.py` writes an **invented** workbook (no real data): suppliers, transactions, staff (all capitals, first name
  and surname in separate cells), free-text notes. 48 written forms of names in 8 groups (see `NOMI`), plus many text
  cells without names (cities, ledger reasons, headers, notes).
- Methods, each giving a set of suggested values that the user would accept as they are:
  1. **Maschera**: the proposals the page makes when the file is loaded (`testiDistinti`, `proposto`).
  2. **rizzo-pii, cella**: each text cell sent alone; values = `FULLNAME` and `ORG` entities.
  3. **rizzo-pii, intestazione**: «Header: text», the natural input for Maschera, which knows the headers.
  4. **rizzo-pii, riga**: the whole row «Header: value; …», as in the 2026-10-12 test.
  5. **Maschera + rizzo-pii intestazione**: the integration being considered. (Also reported: Maschera + cella.)
- Only text cells are sent: Maschera never touches numbers outside the chosen columns, by design.
- Scoring (`valuta.js`): accepted values are applied as Maschera does (a cell equal to a value is replaced whole;
  values of 4+ characters are also searched inside texts, with Maschera's own search). Each occurrence of a true name
  is **covered** (all characters hidden), **partial** or **uncovered**. A **false suggestion** is a value that, applied
  alone, touches no true name.
- rizzo-pii 2.0.0, model rizzo-pii-0.3B v1.5.0, Docker, CPU only, on the author's PC.
- Before writing this prediction, one call was made only to see the response format: «Referente: Marco Galli»
  (a cell of the benchmark) → `FULLNAME`. No other result was seen.

## Prediction

Share of occurrences covered, by group:

| group | Maschera | rizzo-pii, cella | rizzo-pii, intestazione |
|---|---|---|---|
| person, capital initials | 50–65% | ≥ 85% | ≥ 85% |
| person, all capitals | 0% | at least 1 of 2 | at least 1 of 2 |
| person, all lower case | 0% | at most 2 of 3 | at least as cella |
| person, surname first or initial | 100% | ≥ 1 of 2 | ≥ 1 of 2 |
| person, first name or surname alone | ≤ 10% | ≤ 30% | cella + at least 20 points |
| company, with legal form | ≥ 90% | ≥ 70% | ≥ 70% |
| company, without legal form | 50–70% | ≤ 60% | ≤ 60% |
| company, all capitals | 0% | at most 1 of 2 | at most 1 of 2 |

- **False suggestions:** Maschera ≥ 15 (single capitalised words: cities, ledger reasons); rizzo-pii cella ≤ 5.
- **riga** covers at least 5 points less than **intestazione**.
- **Maschera + intestazione** covers at least 20 points more than Maschera alone.
- **Time:** under 1 s per call, under 2 minutes per method.

## Criterion (decided now)

Propose the integration to rizzo-pii's author if **both** hold:
- (a) Maschera + rizzo-pii intestazione covers **at least 15 percentage points** more occurrences than Maschera alone;
- (b) rizzo-pii intestazione alone gives **at most 5** false suggestions.

If (a) fails: no integration proposal, only the cross-links. If (a) holds and (b) fails: proposal with that caveat.
The report is sent in any case, with the numbers as they are.

---

## Italiano

Scritta il 13/10/2026 **prima** della misura e salvata in git prima di farla: la storia di git mostra l'ordine. Non
si modifica dopo i risultati, che vanno in `RISULTATI.md`.

## Domanda

Maschera propone i probabili nomi con regole semplici (parole con l'iniziale maiuscola, non parole comuni da
bilancio, al massimo due volte per foglio). rizzo-pii è un modello italiano per i dati personali che, su un foglio
incollato come testo, ha mancato metà dei cognomi (prova del 12/10/2026). Insieme, con rizzo-pii che suggerisce e
l'utente che conferma in Maschera, trovano chiaramente più nomi di Maschera da sola, senza sommergere l'utente di
suggerimenti sbagliati?

## Impianto

- `banco.py` scrive una cartella di lavoro **inventata** (nessun dato vero): fornitori, movimenti, personale (tutto
  maiuscolo, nome e cognome in celle separate), appunti in testo libero. 48 forme scritte di nomi in 8 gruppi (vedi
  `NOMI`), più molte celle di testo senza nomi (città, causali, intestazioni, note).
- Metodi, ognuno dà un insieme di valori proposti che l'utente accetterebbe così come sono:
  1. **Maschera**: le proposte che la pagina fa quando si carica il file (`testiDistinti`, `proposto`).
  2. **rizzo-pii, cella**: ogni cella di testo inviata da sola; valori = entità `FULLNAME` e `ORG`.
  3. **rizzo-pii, intestazione**: «Intestazione: testo», il modo naturale per Maschera, che conosce le intestazioni.
  4. **rizzo-pii, riga**: tutta la riga «Intestazione: valore; …», come nella prova del 12/10/2026.
  5. **Maschera + rizzo-pii intestazione**: l'integrazione in esame. (Si riporta anche Maschera + cella.)
- Si mandano solo le celle di testo: Maschera non tocca mai i numeri fuori dalle colonne scelte, per progetto.
- Punteggio (`valuta.js`): i valori accettati si applicano come fa Maschera (la cella uguale a un valore si
  sostituisce tutta; i valori di almeno 4 caratteri si cercano anche dentro i testi, con la ricerca di Maschera).
  Ogni occorrenza di un nome vero è **coperta** (tutte le lettere nascoste), **parziale** o **scoperta**. Un **falso
  suggerimento** è un valore che, applicato da solo, non tocca nessun nome vero.
- rizzo-pii 2.0.0, modello rizzo-pii-0.3B v1.5.0, Docker, solo CPU, sul PC dell'autore.
- Prima di scrivere questa predizione è stata fatta una sola chiamata, per vedere il formato della risposta:
  «Referente: Marco Galli» (una cella del banco) → `FULLNAME`. Nessun altro risultato è stato visto.

## Predizione

Quota di occorrenze coperte, per gruppo:

| gruppo | Maschera | rizzo-pii, cella | rizzo-pii, intestazione |
|---|---|---|---|
| persona, iniziali maiuscole | 50–65% | ≥ 85% | ≥ 85% |
| persona, tutto maiuscolo | 0% | almeno 1 su 2 | almeno 1 su 2 |
| persona, tutto minuscolo | 0% | al massimo 2 su 3 | almeno come cella |
| persona, cognome prima o iniziale | 100% | ≥ 1 su 2 | ≥ 1 su 2 |
| persona, solo nome o solo cognome | ≤ 10% | ≤ 30% | cella + almeno 20 punti |
| azienda, con forma societaria | ≥ 90% | ≥ 70% | ≥ 70% |
| azienda, senza forma societaria | 50–70% | ≤ 60% | ≤ 60% |
| azienda, tutto maiuscolo | 0% | al massimo 1 su 2 | al massimo 1 su 2 |

- **Falsi suggerimenti:** Maschera ≥ 15 (parole singole con la maiuscola: città, causali); rizzo-pii cella ≤ 5.
- **riga** copre almeno 5 punti meno di **intestazione**.
- **Maschera + intestazione** copre almeno 20 punti più di Maschera da sola.
- **Tempo:** meno di 1 s per chiamata, meno di 2 minuti per metodo.

## Criterio (deciso adesso)

Si propone l'integrazione all'autore di rizzo-pii se valgono **entrambe**:
- (a) Maschera + rizzo-pii intestazione copre **almeno 15 punti percentuali** di occorrenze in più di Maschera da sola;
- (b) rizzo-pii intestazione da solo dà **al massimo 5** falsi suggerimenti.

Se (a) non vale: nessuna proposta di integrazione, solo i rimandi. Se vale (a) ma non (b): proposta con quella
avvertenza. Il rapporto si manda comunque, con i numeri come sono.
