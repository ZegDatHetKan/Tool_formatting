# Report — Artefatto HTML di revisione formattazione (lettera fittizia)

Accompagna `out/html/lettera_fittizia_review.html`.

## Cos'è

Una pagina HTML statica, pensata per il **cliente** (professionisti legali, non
sviluppatori), che mostra una **lettera fittizia** Bergamo Legal con una legenda
a sinistra. Ogni sezione semantica è evidenziata con un colore; la voce di
legenda corrispondente mostra un riepilogo breve e, **al passaggio del mouse e al
focus da tastiera**, tutti i dettagli di formattazione: font, dimensione,
allineamento, spaziatura, rientro, grassetto/corsivo, keep-with-next ed eredità
dal template. La pagina si apre direttamente come file statico (nessun build).

## Come è stata prodotta (importante)

La lettera fittizia **non è HTML disegnato a mano**: è **generata dal formatter**.
Il flusso è:

1. si inventa il contenuto fittizio e si costruisce un `LetterDocument`;
2. lo si passa a `render_letter` → `out/html/lettera_fittizia.docx` (DOCX reale,
   formattato dallo script, eredita header/footer/margini dal template);
3. si **rilegge** il `.docx` e si emette l'HTML usando le proprietà **effettive**
   di ogni paragrafo (allineamento, dimensione, grassetto/corsivo, rientro,
   spaziatura, keep-with-next).

Quindi l'anteprima e i dettagli della legenda riflettono ciò che lo script
produce davvero, non un'approssimazione. Lo script di generazione è
`tool/scripts/build_review_html.py` (rilancia per rigenerare).

## File prodotti

- `out/html/lettera_fittizia.docx` — la lettera fittizia **formattata dal tool**.
- `out/html/lettera_fittizia_review.html` — l'artefatto di review (generato dal `.docx`).
- `out/html/lettera_fittizia_review_report.md` — questo report.
- `tool/scripts/build_review_html.py` — generatore (render + estrazione + HTML).

## Contenuto fittizio (nessun dato reale)

Lettera inventata: diffida ad adempiere di *Aurora Servizi S.r.l.* verso *Sig.
Mario Rossi* (nomi, indirizzi, importi, numeri di fattura e l'avvocato firmatario
*Avv. Giulia Conti* sono tutti **inventati**). Nessun cliente reale, nessuna
controversia reale, nessun dato preso da `previous_works/`.

La lettera esercita tutte le sezioni richieste: metodo di trasmissione, data/
luogo, destinatario, OGGETTO (etichetta + contenuto), apertura, paragrafo di
corpo, titoletto a sinistra (“Fatto”), titolo rituale centrato (“DIFFIDA”), voci
di elenco, congedo, blocco firma, allegati e postilla.

## Fonti usate

**Valori implementati** (ciò che la pagina mostra come “regola attuale”):
- `formatters/letters.py` — letto direttamente per costanti e default
  (`SIZE_BODY=12`, `SIZE_HEADING=16`, `SIZE_SUBHEADING=14`, `SIZE_SMALL=10`,
  `RECIPIENT_INDENT_CM=8.5`, `LIST_INDENT_CM=0.5`; default `subject_style=SPLIT`,
  `subject_label_center=True`, `subject_content_center=False`, `opening_formal=True`,
  `date_above_recipient=True`, `merge_courtesy_appellative=True`) e per le
  spaziature di ogni `_emit_*`.
- `docs/regole_di_formattazione_in_lavorazione.md` — specchio 1:1 del codice.

**Razionale e cautele** (note al cliente e punti sensibili):
- `docs/regole_formattazione_lettere.md`
- `docs/client_approval_points.md`
- `feedback.md`
- `docs/dubbi_formattazione.md`

**Eredità dal template**: `assets/Template_Vuoto.docx` (header/footer/margini),
non modificato.

## Conflitti tra fonti

Verifica eseguita confrontando i valori letti **a runtime** da `formatters/letters.py`
con `docs/regole_di_formattazione_in_lavorazione.md` e `docs/regole_formattazione_lettere.md`.

- **Nessun conflitto doc↔codice riscontrato**: font, dimensioni, allineamenti,
  rientri, spaziature, default e regole speciali coincidono fra codice e
  documenti. La pagina riflette quindi i valori del codice.
- **Tensione nei dati di feedback (non un conflitto doc↔codice)**: la PEC `003`
  ottenne 5/5 pur con OGGETTO centrato, mentre `027` (4/5) contestava la forma
  centrata/split. È un'incoerenza del cliente, **già risolta per decisione**
  (D1): OGGETTO split con etichetta centrata 16 pt + contenuto giustificato 12 pt.
  Segnalata nella pagina come “punto sensibile / da approvare”.

## Punti da far approvare al cliente

Mostrati nella pagina con linguaggio non definitivo (“Regola attuale”, “Da
approvare”, “Punto sensibile”, “Ereditato dal template”, “Generato dallo script”):

1. **OGGETTO** (punto sensibile): etichetta `OGGETTO:` centrata 16 pt grassetto;
   contenuto a capo, giustificato, 12 pt grassetto. Uniforma anche le lettere che
   storicamente usavano la forma inline.
2. **Destinatario**: rientro 8,5 cm; appellativo personale unito al nome
   (“Egr. Sig. Mario Rossi”); “Spett.le” davanti a società su riga propria.
3. **Gerarchia dei titoli**: titoletti a sinistra 14 pt vs titoli rituali
   centrati 16 pt; verbi dispositivi (“CHIEDE ALTRESÌ”) resi come titoletto +
   contenuto a capo.
4. **Blocco firma**: allineato a destra, nomi in grassetto, spaziatura ariosa
   (prima riga 18 pt, separazione fra più firmatari) — da approvare.
5. **Eredità dal template**: header (letterhead Garamond + logo) e footer (dati
   Studio) provengono dal template e non sono generati; il corpo non li duplica.
6. **Principio “mai completare”**: data, destinatario, saluto, oggetto, firma
   sono opzionali e resi solo se forniti; lo script non inventa contenuti.

## Interazione e accessibilità

- Dettagli completi rivelati su **hover** del mouse **e** su **focus** da
  tastiera (voci di legenda e sezioni della lettera sono focusabili con `Tab`).
- Evidenziazione reciproca: attivando una voce di legenda si evidenzia la sezione
  corrispondente nella lettera, e viceversa (colore condiviso sempre visibile +
  rinforzo al focus/hover).
- In **stampa/PDF** tutti i dettagli sono espansi e le evidenziazioni restano
  visibili (media query dedicata).
- Layout responsive: legenda fissa a sinistra su desktop, sopra la lettera su
  mobile.

## Esito checklist di accettazione

- ✔ ogni voce di legenda ha una sezione evidenziata corrispondente (16/16);
- ✔ ogni sezione evidenziata ha una voce di legenda (16 chiavi, corrispondenza piena);
- ✔ la lettera è prodotta da `render_letter` (DOCX reale) e l'HTML ne legge i valori;
- ✔ ogni voce espone i dettagli completi su hover e focus;
- ✔ la lettera fittizia include tutti i tipi di sezione richiesti;
- ✔ la pagina si apre come file HTML statico, senza build;
- ✔ nessun dato cliente reale;
- ✔ `previous_works/` e `assets/Template_Vuoto.docx` non modificati; formatter non modificato;
- ✔ il report elenca fonti, conflitti (nessuno doc↔codice) e punti di approvazione.

## Dopo il riscontro del cliente

Le decisioni approvate andranno ripiegate (manualmente, in fase successiva) in:
`formatters/letters.py`, `docs/regole_di_formattazione_in_lavorazione.md`,
`docs/regole_formattazione_lettere.md`. Questo artefatto **non modifica** il
formatter.
