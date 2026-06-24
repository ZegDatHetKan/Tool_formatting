# Processo standard per creare formatter documentali

Questo documento trasforma il lavoro fatto sulle lettere Bergamo Legal in un
processo riusabile per nuovi studi, nuovi dataset e nuove famiglie documentali.

L'obiettivo non è copiare le regole delle lettere. L'obiettivo è ripetere il
metodo: partire da esempi approvati, estrarre regole generalizzabili, costruire
uno script deterministico, validarlo contro gli esempi storici e ottenere il
consenso finale del cliente tramite un artefatto visivo.

## 1. Obiettivo del processo

Per ogni famiglia documentale bisogna produrre:

- una classificazione chiara del dataset;
- una specifica `.md` ragionata delle regole di formattazione;
- una specifica `.md` specchio 1:1 del codice;
- uno script Python deterministico che riproduce la struttura;
- una validazione contro gli esempi forniti;
- un documento HTML di consenso cliente;
- una versione finale aggiornata dopo il feedback del cliente.

La regola di fondo è: **AI per analisi e astrazione, Python per misurazione,
rendering, confronto e produzione ripetibile**.

## 2. Input richiesti

Ogni nuovo progetto dovrebbe partire da un dataset con:

- documenti input grezzi;
- output già formattati o approvati dal cliente, quando disponibili;
- eventuali report/verification storici;
- template, carta intestata, footer o modello vuoto;
- feedback cliente, anche informale;
- indicazione della famiglia documentale da trattare per prima.

Struttura consigliata:

```text
assets/
  Template_Vuoto.docx
previous_works/
  manifest.json
  input_001.docx
  output_001.docx
  report_001.md
  verification_001.md
  ...
docs/
formatters/
tool/scripts/
out/
```

`previous_works/manifest.json` deve diventare la fonte di verità per ID,
famiglia documentale, soggetto, path input/output/report/verification e stato di
review. I file storici dentro `previous_works/` sono read-only.

## 3. Fase 1 - Profilazione del dataset

Scopo: capire cosa c'è nel dataset e scegliere una famiglia documentale per il
primo formatter.

Attività:

- verificare numero di documenti, coppie input/output e file mancanti;
- classificare i documenti per famiglia (`letters`, `acts`, `contracts`, ecc.);
- scegliere una sola famiglia per il primo ciclo;
- identificare esempi positivi, esempi problematici e casi limite;
- leggere feedback cliente e associarlo agli ID corretti;
- misurare i DOCX: margini, header/footer, font, dimensioni, spaziature,
  allineamenti, rientri, keep-with-next, grassetti e corsivi.

Output consigliati:

- `previous_works/manifest.json`;
- `feedback.md`;
- un report di corpus con ID analizzati, esclusi e ambigui.

Nel caso lettere, questa fase ha prodotto e usato:

- `previous_works/manifest.json`;
- `feedback.md`;
- `docs/report_esecuzione_lettere.md`;
- i 15 record `document_type == "letters"` come prima famiglia.

## 4. Fase 2 - Regole e formatter deterministico

Scopo: trasformare le osservazioni in una struttura riusabile.

Per ogni famiglia documentale creare:

- una specifica ragionata, per esempio
  `docs/regole_formattazione_<famiglia>.md`;
- una specifica specchio del codice, per esempio
  `docs/regole_di_formattazione_<famiglia>_in_lavorazione.md`;
- uno script Python, per esempio `formatters/<famiglia>.py`;
- eventuali helper o CLI di validazione.

Principi:

- il formatter deve essere **content-driven**: riceve contenuti semantici e li
  rende secondo regole fisse;
- lo script è la fonte di verità operativa;
- il documento specchio deve riportare esattamente costanti, allineamenti,
  spaziature, rientri e comportamenti del codice;
- non si modifica mai a mano un output generato per correggere un problema di
  layout: si corregge lo script e si rigenera;
- non si implementano più famiglie insieme se non strettamente necessario.

Nel caso lettere:

- regole ragionate: `docs/regole_formattazione_lettere.md`;
- specchio 1:1: `docs/regole_di_formattazione_in_lavorazione.md`;
- script: `formatters/letters.py`;
- validazione: `validate_letters.py`.

## 5. Fase 3 - Validazione cieca e confronto

Scopo: verificare se lo script generalizza davvero.

Procedura:

1. prendere gli input storici della famiglia scelta;
2. ricostruire contenuto semantico senza usare gli output come guida durante il
   rendering;
3. generare nuovi output con lo script;
4. solo dopo la generazione, confrontare gli output prodotti con gli output
   storici approvati;
5. annotare differenze, regressioni, differenze volute e dubbi.

La distinzione è importante:

- **output storico**: documento fornito o approvato dal cliente;
- **output generato**: documento prodotto dal nostro script;
- **differenza voluta**: modifica decisa per uniformare o migliorare;
- **regressione**: comportamento da correggere;
- **dubbio**: scelta che richiede feedback utente/cliente.

Nel caso lettere:

- le ricostruzioni sono in `documenti_generati/storiche/`;
- i dubbi e confronti sono in `docs/dubbi_formattazione.md`;
- il report di validazione è in `docs/report_esecuzione_lettere.md`.

## 6. Fase 4 - Feedback, decisioni e revisione

Scopo: trasformare incongruenze e preferenze in decisioni esplicite.

Attività:

- associare ogni feedback al giusto ID del corpus;
- verificare se il feedback è sensato sul DOCX associato;
- separare feedback sulla famiglia corrente da feedback su famiglie future;
- porre domande al cliente/utente solo sulle scelte che cambiano davvero la
  formattazione;
- applicare le decisioni nel codice e nella spec specchio;
- rigenerare gli output e ripetere la validazione.

Nel caso lettere:

- `feedback.md` collega nomi cliente, ID corpus, input/output e verifica;
- i feedback prioritari erano `027` e `028`;
- alcune note su `acts` sono state conservate ma non applicate alle lettere.

## 7. Fase 5 - HTML di consenso cliente

Scopo: fornire al cliente un artefatto visuale per approvare la formattazione
prima del consolidamento finale.

L'HTML deve contenere:

- un documento fittizio, senza dati reali;
- una preview della pagina o del documento;
- legenda per ogni sezione;
- colori coerenti tra legenda e sezione evidenziata;
- riepilogo rapido visibile;
- dettagli completi su hover/focus: font, dimensione, allineamento, rientro,
  spaziatura, grassetto, corsivo, keep-with-next, template/header/footer;
- report con fonti usate, conflitti e punti da approvare.

Idealmente l'HTML non deve essere disegnato a mano. Deve essere prodotto da uno
script ripetibile che usa l'output reale del formatter o dati misurati dal DOCX.

Nel caso lettere:

- generatore: `tool/scripts/build_review_html.py`;
- HTML: `out/html/lettera_fittizia_review.html`;
- report: `out/html/lettera_fittizia_review_report.md`;
- il flusso reale è: contenuto fittizio -> `LetterDocument` -> DOCX -> PDF
  LibreOffice -> immagine/coordinate -> HTML con overlay e legenda.

## 8. Fase 6 - Consolidamento finale

Dopo il feedback sul documento HTML:

- aggiornare lo script Python con le decisioni approvate;
- aggiornare la spec specchio 1:1;
- aggiornare la spec ragionata;
- aggiornare eventuali prompt/istruzioni agent;
- rigenerare esempi e HTML;
- rieseguire validazioni;
- registrare cosa è approvato e cosa resta fuori ambito.

Il consenso cliente deve chiudere il ciclo: dopo questa fase non devono restare
dubbi su come la famiglia documentale approvata viene formattata.

## 9. Checklist per un nuovo progetto

Prima di iniziare:

- [ ] dataset raccolto;
- [ ] template/carta intestata disponibile;
- [ ] coppie input/output copiate in area read-only;
- [ ] manifest creato;
- [ ] feedback cliente importato;
- [ ] famiglia documentale prioritaria scelta.

Durante la profilazione:

- [ ] documenti classificati;
- [ ] esempi rappresentativi selezionati;
- [ ] DOCX misurati programmaticamente;
- [ ] regole ricorrenti separate dai casi isolati;
- [ ] dubbi documentati.

Durante l'implementazione:

- [ ] script Python deterministico creato;
- [ ] spec ragionata scritta;
- [ ] spec specchio sincronizzata col codice;
- [ ] output generati fuori da `previous_works/`;
- [ ] nessun dato storico modificato.

Durante la validazione:

- [ ] rigenerazione dagli input eseguita senza usare gli output come guida;
- [ ] confronto con output storici completato;
- [ ] regressioni separate da differenze volute;
- [ ] feedback/domande preparate.

Durante il consenso cliente:

- [ ] documento fittizio creato;
- [ ] HTML con legenda e sezioni colorate generato;
- [ ] punti sensibili evidenziati;
- [ ] report HTML scritto;
- [ ] feedback finale raccolto.

Finalizzazione:

- [ ] decisioni cliente riportate nel codice;
- [ ] decisioni riportate nella spec specchio;
- [ ] decisioni riportate nella spec ragionata;
- [ ] validazioni rieseguite;
- [ ] artefatti finali consegnabili.

## 10. Ruoli: AI e Python

AI dovrebbe occuparsi di:

- classificare e descrivere strutture;
- proporre blocchi semantici;
- confrontare differenze e formulare ipotesi;
- redigere regole, dubbi e domande;
- trasformare feedback informale in decisioni verificabili.

Python dovrebbe occuparsi di:

- leggere e misurare DOCX;
- generare DOCX;
- confrontare file e layout;
- creare report ripetibili;
- generare HTML/PDF/immagini;
- proteggere i riferimenti storici da modifiche accidentali.

Il processo deve rimanere model-agnostic: Claude, Codex o altri modelli possono
aiutare nell'analisi, ma il risultato operativo deve vivere in script,
specifiche e dati riproducibili.

## 11. Prompt per Claude per completare/correggere questo processo

```text
Sei in /home/giulio/code/LexForward/moduli/formatter.

Leggi docs/processo_standard_formatter.md e completalo/correggilo usando i
passaggi reali che hai eseguito in questo progetto.

Obiettivo: trasformare il lavoro sulle lettere Bergamo Legal in un processo
standard riusabile per altri studi, altri dataset e altre famiglie documentali.

Non riscrivere il documento in modo astratto. Mantieni una pipeline concreta:
dataset -> manifest -> analisi -> regole -> script Python -> validazione cieca
-> confronto output -> feedback -> HTML di consenso -> consolidamento finale.

Usa come fonti:
- docs/regole_formattazione_lettere.md
- docs/regole_di_formattazione_in_lavorazione.md
- docs/report_esecuzione_lettere.md
- docs/dubbi_formattazione.md
- feedback.md
- out/html/lettera_fittizia_review_report.md
- tool/scripts/build_review_html.py
- formatters/letters.py

Non modificare previous_works/ e non cambiare gli output storici.
Se trovi passaggi mancanti o inaccurati, correggi il documento e segnala in
fondo cosa hai cambiato.
```

