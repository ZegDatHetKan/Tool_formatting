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

La lettera fittizia **non è HTML disegnato a mano** e **non è una ricostruzione
CSS**: l'anteprima è il **render reale del DOCX** prodotto dal formatter. Il flusso
(`tool/scripts/build_review_html.py`):

1. si inventa il contenuto fittizio e si costruisce un `LetterDocument`;
2. `render_letter` → `out/lettera_fittizia.docx` (DOCX reale, eredita
   header/footer/margini dal template);
3. **LibreOffice headless** converte il DOCX in PDF → `out/lettera_fittizia.pdf`
   (impaginazione, font, intestazione e piè di pagina effettivi: è il deliverable);
4. **solo per l'anteprima web** si crea una copia a **pagina unica** (altezza
   pagina calcolata per contenere tutto senza interruzioni), resa in **immagine**
   (`pdftoppm`, 150 dpi) → `lettera_fittizia_review.png`;
5. con `pdftotext -bbox` si estraggono le **coordinate** di ogni parola; le
   parole sono raggruppate in righe e classificate per sezione, ottenendo i
   **riquadri** (in % della pagina) sovrapposti all'immagine;
6. le schede-sezione mostrano i dettagli **letti dal DOCX reale** (allineamento,
   dimensione, grassetto/corsivo, rientro, spaziatura, keep-with-next).

Quindi font e impaginazione sono quelli effettivi dello script; l'anteprima è
un'unica pagina continua. L'immagine è **pulita** (senza colori); tinte e linee
sono **interattive**, disegnate nell'HTML sopra l'immagine.

### Nota sui font del rendering

Il DOCX dichiara i font reali: **Times New Roman** (corpo, come le lettere di
output) e **Garamond** (intestazione, dal template). Sulla macchina di rendering
questi font proprietari non sono installati, quindi per l'anteprima si usano i
sostituti standard, installati nello spazio utente:
- **Garamond → EB Garamond** (font libero OFL, equivalente aperto di Garamond);
- **Times New Roman → Liberation Serif** (sostituto *metric-identico*, indistinguibile nel corpo).

Aperto in Word con i font proprietari installati, il documento appare nel font
nominale. La sostituzione riguarda solo l'immagine di anteprima, non il `.docx`.

## File prodotti

- `out/lettera_fittizia.docx` — lettera fittizia **formattata dal tool** (pulita).
- `out/lettera_fittizia.pdf` — stessa lettera in PDF (render LibreOffice, pulita).
- `out/html/lettera_fittizia_review.png` — render reale a pagina unica (anteprima).
- `out/html/lettera_fittizia_review.html` — artefatto di review (cartellone interattivo).
- `out/html/lettera_fittizia_review_report.md` — questo report.
- `tool/scripts/build_review_html.py` — generatore (render → PDF → immagine → HTML).

## Contenuto fittizio (nessun dato reale)

Lettera inventata: diffida ad adempiere di *Aurora Servizi S.r.l.* verso *Sig.
Mario Rossi* (nomi delle parti, indirizzi, importi e numeri di fattura sono
**inventati**). Il firmatario è *Avv. Matteo Bertocchi* (socio dello Studio, già
presente nel letterhead del template). Nessun cliente reale, nessuna controversia
reale, nessun dato preso da `previous_works/`.

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

- **Cartellone unico**: la lettera (a sinistra) e le schede-sezione (a destra)
  stanno su un'unica tela, non in due colonne indipendenti. La lettera è una
  **pagina unica continua** (niente taglio fra pagine).
- **Tinte permanenti leggere** su ogni sezione della lettera, che si
  **intensificano** al passaggio del mouse/focus (sulla sezione o sulla scheda).
- **Schede a fianco** in colonna ordinata (in flusso, non si sovrappongono mai);
  riepilogo sempre visibile, dettagli completi che si espandono su hover/focus.
- **Linee di collegamento permanenti e leggere** (nel colore della sezione) fra
  ogni sezione e la sua scheda; si rinforzano sull'elemento attivo.
- Riquadri ricavati dalle **coordinate reali del PDF** (`pdftotext -bbox`).
- Navigabile da **tastiera** (`Tab`); in **stampa/PDF** dettagli espansi.
- Responsive: su schermi stretti le schede vanno sotto la lettera (linee disattivate).
- Niente “hero”: solo una sottile riga informativa in testa.

## Esito checklist di accettazione

- ✔ ogni voce di legenda ha una sezione evidenziata corrispondente (16/16);
- ✔ ogni sezione evidenziata ha una voce di legenda (16 chiavi, corrispondenza piena);
- ✔ l'anteprima è il **render reale** del DOCX (DOCX → PDF LibreOffice → immagine), non CSS;
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
