# Experiment: rizzo-pii as a name-suggestion engine for Maschera — results · Esperimento: rizzo-pii come motore di suggerimenti per Maschera — risultati

**English** · [Italiano](#italiano)

Run on 2026-10-13, after the prediction ([PREDIZIONE.md](PREDIZIONE.md), commit `417c71b`). Raw outputs:
[`risultati/rizzo.json`](risultati/rizzo.json) and [`risultati/valutazione.txt`](risultati/valutazione.txt).
rizzo-pii 2.0.0, model rizzo-pii-0.3B v1.5.0, Docker, CPU only.

## In one sentence

On this invented workbook, rizzo-pii given one cell at a time finds **90% of the name occurrences with 1 false
suggestion**; Maschera alone finds 53% with 23 false suggestions; together they cover **100%**. The success criterion
is met. Caveat: invented data, probably easier than real files.

## Main numbers

59 occurrences of names (17 inside a sentence) in 120 text cells.

| method | suggested values | covered | partial | uncovered | false suggestions |
|---|---|---|---|---|---|
| Maschera | 44 | 31 (53%) | 1 | 27 | 23 |
| rizzo-pii, cella | 47 | 54 (92%) | 3 | 2 | 5 |
| rizzo-pii, intestazione | 42 | 53 (90%) | 3 | 3 | 1 |
| rizzo-pii, riga | 42 | 49 (83%) | 8 | 2 | 0 |
| **Maschera + rizzo-pii intestazione** | 73 | **59 (100%)** | 0 | 0 | 24 |
| Maschera + rizzo-pii cella | 73 | 58 (98%) | 0 | 1 | 24 |
| *(exploratory) Maschera filtered + rizzo-pii intestazione* | 53 | 57 (97%) | 0 | 2 | 5 |

By group (covered, `+Np` = partial):

| group (occurrences) | Maschera | cella | intestazione | riga | Maschera + intestazione |
|---|---|---|---|---|---|
| person, capital initials (16) | 9 +1p | 16 | 16 | 16 | 16 |
| person, all capitals (2) | 0 | 2 | 2 | 2 | 2 |
| person, all lower case (3) | 1 | 3 | 3 | 3 | 3 |
| person, surname first or initial (2) | 2 | 1 +1p | 1 +1p | 1 +1p | 2 |
| person, first name or surname alone (13) | 1 | 12 | 13 | 13 | 13 |
| company, with legal form (10) | 10 | 10 | 10 | 7 +3p | 10 |
| company, without legal form (10) | 7 | 7 +2p | 5 +2p | 5 +3p | 10 |
| company, all capitals (3) | 1 | 3 | 3 | 2 +1p | 3 |

Time: 111 calls in 3.8 s (cella), 116 in 3.4 s (intestazione), 50 in 1.7 s (riga); slowest call 0.08 s.

## Criterion

- (a) Maschera + intestazione covers **+47 points** over Maschera alone (100% vs 53%; threshold 15): **met**.
- (b) rizzo-pii intestazione gives **1** false suggestion («NAPOLI», from «TRASFERITO DA NAPOLI»; threshold 5): **met**.

→ By the rule written before the run, the integration is worth proposing.

## Prediction vs result

21 of 29 predicted items right, 8 wrong. Six of the wrong ones underestimated rizzo-pii:
- all lower case: predicted at most 2 of 3, got 3 of 3;
- first name or surname alone (cella): predicted ≤ 30%, got 92%; intestazione did not add 20 points (100% vs 92%,
  ceiling);
- company without legal form (cella): predicted ≤ 60%, got 70%;
- company all capitals: predicted at most 1, got 3 of 3 (the group also had 3 occurrences, not 2), cella and
  intestazione.

Two were wrong about Maschera (predicted 0% for all lower case and for company all capitals, got 1 each): the
prediction forgot that Maschera's search is case-insensitive, so «Conti Davide» also covers «conti davide» and
«Tecnofarm S.p.A.» also covers «TECNOFARM S.P.A.».

## What the numbers do not show

- **Invented data, written by us.** Typical Italian names, clean cells. In the 2026-10-12 test on a real
  spreadsheet pasted as text, rizzo-pii masked 7 surnames out of 14. Real files are probably harder; this run does
  not prove the 90% on them.
- **One run, one workbook, 59 occurrences.** Small numbers: one cell more or less moves a group by many points.
- **Labels sometimes swapped:** «F.lli Ruggeri» comes as a person, «Ruggeri» (a person) as a company. For suggestions
  this does not matter; the category Maschera proposes would need the user's check.
- **Company names cut:** «Grafiche Serra di Fabio Serra» → only «Fabio Serra»; «Officina Meccanica Pautasso» → only
  «Officina Meccanica». Maschera covers them, which is why the union reaches 100%.
- **Maschera's false suggestions are a Maschera problem.** Single capitalised words (cities, ledger reasons) are
  proposed as names. In real files they often repeat and get filtered, but here they did not.
- The exploratory row was added **after** seeing the results: Maschera's one-word proposals are kept only if
  rizzo-pii also recognises them. It was not predicted and does not count for the criterion.

## What an integration could look like

- A **separate, opt-in build** of Maschera (for example `build.py --rizzo-pii`) whose Content-Security-Policy allows
  connections **only** to `http://127.0.0.1:5005`. The standard build stays fully offline.
- Maschera sends only text cells, as «Header: text», to the local rizzo-pii; `FULLNAME` and `ORG` become proposals
  the user confirms, like today's. Numbers are never sent.
- If rizzo-pii is not running, Maschera works as today.

## Reproduce

```
# image built from https://github.com/Rizzo-AI-Academy/rizzo-pii at release 2.0.0
docker run -d --name rizzo-pii -p 127.0.0.1:5005:5005 rizzo-pii:2.0.0
python experiments/rizzo-pii/banco.py <dir>
python experiments/rizzo-pii/rizzo.py <dir>
node experiments/rizzo-pii/valuta.js <dir>
```

---

## Italiano

Eseguito il 13/10/2026, dopo la predizione ([PREDIZIONE.md](PREDIZIONE.md), commit `417c71b`). Uscite grezze:
[`risultati/rizzo.json`](risultati/rizzo.json) e [`risultati/valutazione.txt`](risultati/valutazione.txt).
rizzo-pii 2.0.0, modello rizzo-pii-0.3B v1.5.0, Docker, solo CPU.

## In una frase

Su questa cartella di lavoro inventata, rizzo-pii con una cella alla volta trova **il 90% delle occorrenze di nomi
con 1 falso suggerimento**; Maschera da sola ne trova il 53% con 23 falsi suggerimenti; insieme coprono il **100%**.
Il criterio di successo è raggiunto. Avvertenza: dati inventati, probabilmente più facili dei file veri.

## I numeri principali

59 occorrenze di nomi (17 dentro una frase) in 120 celle di testo.

| metodo | valori proposti | coperte | parziali | scoperte | falsi suggerimenti |
|---|---|---|---|---|---|
| Maschera | 44 | 31 (53%) | 1 | 27 | 23 |
| rizzo-pii, cella | 47 | 54 (92%) | 3 | 2 | 5 |
| rizzo-pii, intestazione | 42 | 53 (90%) | 3 | 3 | 1 |
| rizzo-pii, riga | 42 | 49 (83%) | 8 | 2 | 0 |
| **Maschera + rizzo-pii intestazione** | 73 | **59 (100%)** | 0 | 0 | 24 |
| Maschera + rizzo-pii cella | 73 | 58 (98%) | 0 | 1 | 24 |
| *(esplorativo) Maschera filtrata + rizzo-pii intestazione* | 53 | 57 (97%) | 0 | 2 | 5 |

Per gruppo (coperte, `+Np` = parziali):

| gruppo (occorrenze) | Maschera | cella | intestazione | riga | Maschera + intestazione |
|---|---|---|---|---|---|
| persona, iniziali maiuscole (16) | 9 +1p | 16 | 16 | 16 | 16 |
| persona, tutto maiuscolo (2) | 0 | 2 | 2 | 2 | 2 |
| persona, tutto minuscolo (3) | 1 | 3 | 3 | 3 | 3 |
| persona, cognome prima o iniziale (2) | 2 | 1 +1p | 1 +1p | 1 +1p | 2 |
| persona, solo nome o solo cognome (13) | 1 | 12 | 13 | 13 | 13 |
| azienda, con forma societaria (10) | 10 | 10 | 10 | 7 +3p | 10 |
| azienda, senza forma societaria (10) | 7 | 7 +2p | 5 +2p | 5 +3p | 10 |
| azienda, tutto maiuscolo (3) | 1 | 3 | 3 | 2 +1p | 3 |

Tempo: 111 chiamate in 3,8 s (cella), 116 in 3,4 s (intestazione), 50 in 1,7 s (riga); la più lenta 0,08 s.

## Criterio

- (a) Maschera + intestazione copre **47 punti** più di Maschera da sola (100% contro 53%; soglia 15): **raggiunto**.
- (b) rizzo-pii intestazione dà **1** falso suggerimento («NAPOLI», da «TRASFERITO DA NAPOLI»; soglia 5):
  **raggiunto**.

→ Per la regola scritta prima della misura, l'integrazione merita di essere proposta.

## Predizione e risultato

21 voci previste giuste su 29, 8 sbagliate. Sei delle sbagliate sottostimavano rizzo-pii:
- tutto minuscolo: previsto al massimo 2 su 3, ottenuto 3 su 3;
- solo nome o solo cognome (cella): previsto ≤ 30%, ottenuto 92%; intestazione non ha aggiunto 20 punti (100% contro
  92%, al tetto);
- azienda senza forma societaria (cella): previsto ≤ 60%, ottenuto 70%;
- azienda tutto maiuscolo: previsto al massimo 1, ottenuto 3 su 3 (il gruppo aveva 3 occorrenze, non 2), sia cella
  sia intestazione.

Due erano sbagliate su Maschera (previsto 0% per tutto minuscolo e per azienda tutto maiuscolo, ottenuto 1 ciascuno):
la predizione dimenticava che la ricerca di Maschera ignora maiuscole e minuscole, quindi «Conti Davide» copre anche
«conti davide» e «Tecnofarm S.p.A.» copre anche «TECNOFARM S.P.A.».

## Cosa i numeri non dicono

- **Dati inventati, scritti da noi.** Nomi italiani tipici, celle pulite. Nella prova del 12/10/2026 su un foglio
  vero incollato come testo, rizzo-pii ha mascherato 7 cognomi su 14. I file veri sono probabilmente più difficili:
  questa misura non dimostra il 90% su di loro.
- **Una misura, una cartella di lavoro, 59 occorrenze.** Numeri piccoli: una cella in più o in meno sposta un gruppo
  di molti punti.
- **Etichette a volte scambiate:** «F.lli Ruggeri» esce come persona, «Ruggeri» (una persona) come azienda. Per
  suggerire non conta; la categoria proposta da Maschera andrebbe controllata dall'utente.
- **Nomi di azienda tagliati:** «Grafiche Serra di Fabio Serra» → solo «Fabio Serra»; «Officina Meccanica Pautasso»
  → solo «Officina Meccanica». Maschera li copre, ed è per questo che l'unione arriva al 100%.
- **I falsi suggerimenti di Maschera sono un problema di Maschera.** Le parole singole con la maiuscola (città,
  causali) vengono proposte come nomi. Nei file veri spesso si ripetono e vengono filtrate, qui no.
- La riga esplorativa è stata aggiunta **dopo** aver visto i risultati: le proposte di Maschera di una sola parola
  restano solo se le riconosce anche rizzo-pii. Non era prevista e non conta per il criterio.

## Come potrebbe essere un'integrazione

- Una **versione separata e facoltativa** di Maschera (per esempio `build.py --rizzo-pii`), con una
  Content-Security-Policy che ammette connessioni **solo** verso `http://127.0.0.1:5005`. La versione normale resta
  del tutto senza rete.
- Maschera manda a rizzo-pii in locale solo le celle di testo, come «Intestazione: testo»; `FULLNAME` e `ORG`
  diventano proposte che l'utente conferma, come oggi. I numeri non partono mai.
- Se rizzo-pii non è acceso, Maschera funziona come oggi.

## Riprodurre

```
# immagine costruita da https://github.com/Rizzo-AI-Academy/rizzo-pii alla versione 2.0.0
docker run -d --name rizzo-pii -p 127.0.0.1:5005:5005 rizzo-pii:2.0.0
python experiments/rizzo-pii/banco.py <cartella>
python experiments/rizzo-pii/rizzo.py <cartella>
node experiments/rizzo-pii/valuta.js <cartella>
```
