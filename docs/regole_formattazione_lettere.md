# Regole di formattazione — Lettere Bergamo Legal

Documento di specifica per il formatter deterministico delle **lettere** dello
Studio Bergamo Legal. Le regole sono inferite dal corpus storico
(`previous_works/`), usando **`previous_works/manifest.json` come unica fonte di
classificazione**. Sono stati analizzati esclusivamente i record con
`document_type == "letters"`.

> Ambito: **solo lettere**. Atti, mediazioni, ricorsi, memorie e querele sono
> fuori ambito e non sono trattati qui (cfr. `docs/letter_formatter_goal.md`).

## Fonte di verità e correzioni ad-hoc

> Per lo **specchio 1:1 dei valori applicati dal codice** (font, dimensioni,
> spaziature, condizioni) vedi `docs/regole_di_formattazione_in_lavorazione.md`.
> Questo documento spiega invece il *perché* (corpus, feedback, incertezze).

Lo **script Python `formatters/letters.py` è la fonte di verità** della
formattazione: ogni regola di stile vive nel codice, non in patch manuali. I
documenti si producono **sempre** facendo passare il contenuto semantico
attraverso questo formatter; non si modifica mai a mano il `.docx` di output (la
modifica andrebbe persa al render successivo e l'errore si ripresenterebbe).

Se un testo **non riesce a essere formattato correttamente** seguendo lo script
così com'è, **è consentito intervenire ad-hoc sullo script** (estendere uno
stile, aggiungere un tipo di blocco, correggere un default) — ma è da **evitare
se non necessario**. Ogni intervento deve preservare uno stile e un decoro
adatti a uno studio legale: **niente errori, niente impaginazioni strane**,
gerarchia visiva coerente. Dopo ogni modifica: rigenerare il documento col tool
ed eseguire `validate_letters.py` come verifica di regressione.

## 0. Record di lettere analizzati

Tutti i 15 record `letters` del manifest (tutti `classification_confidence:
high`, `needs_review: false`):

| ID  | Soggetto (subject del manifest)                                  | Tipo funzionale            |
|-----|------------------------------------------------------------------|----------------------------|
| 001 | lettera-informativa-consiglieri-mueller-madone                   | Informativa / invito       |
| 002 | lettera-consiglieri-comunali-muller-madone (versione difensori)  | Informativa / invito       |
| 003 | pec-contestazione-fattura-papi-solutions                         | Contestazione (PEC)        |
| 004 | comunicazione-ritiro-marchio-uibm-bergamo-legal                  | Comunicazione/istanza      |
| 005 | comunicazione-apple-business-duns-bertocchi                      | Comunicazione (e-mail)     |
| 007 | diffida-rimozione-recensioni-casu                                | Diffida + allegati         |
| 010 | diffida-pagamento-restituzione-tecnico-bertocchi                 | Diffida e messa in mora    |
| 011 | comunicazione-consenso-differimento-mediazione-chebotareva       | Comunicazione breve        |
| 013 | lettera-riscontro-contestazione-de-luca                          | Riscontro / contestazione  |
| 015 | recesso-gravi-motivi-sublocazione-avantgarde                     | Recesso (premesso/dichiara)|
| 025 | lettera-trustpilot-recensioni-hormonal-casula                    | Lettera a piattaforma      |
| 027 | lettera-trustpilot-hormonal-holding-recensioni                   | Lettera a piattaforma      |
| 028 | diffida-rimozione-recensioni-locafaro                            | Diffida + Fatto/Diritto    |
| 029 | pec-contestazione-fattura-papi-solutions                         | Contestazione (PEC)        |
| 032 | aggiornamento-duns-apple-business-bertocchi                      | Comunicazione (e-mail)     |

I codici di regola citati (`HARD_*`, `STYLE_*`) provengono dai `report_*.md` e
`verification_*.md` storici e sono usati qui solo come riferimento di
tracciabilità.

## 1. Scheletro canonico della lettera

Le lettere sono composte da **blocchi semantici** in quest'ordine. I blocchi
contrassegnati *(opzionale)* compaiono solo quando la lettera lo richiede.

```
[ delivery_method  ]   metodo di trasmissione      (opzionale)
[ date_place       ]   luogo e data
[ recipient_block  ]   destinatario (Spett.le / Egr. …)
[ subject          ]   OGGETTO
[ opening          ]   formula di apertura (saluto)
[ body_blocks      ]   corpo: paragrafi, titoli, elenchi   (1..N)
[ closing          ]   formula di congedo
[ signature_block  ]   firma/e
[ attachments      ]   Allegati                    (opzionale)
[ postscript       ]   nota/disclaimer finale      (opzionale)
```

`delivery_method` e `date_place` possono scambiarsi di posizione: in alcune
lettere il metodo di trasmissione è la **prima riga del blocco destinatario**
(es. `004` "Tramite servizio online UIBM / PEC a:", `025`/`027` "Via PEC"),
in altre è una riga autonoma in alto (`003`/`029` "Trasmessa a mezzo PEC" a
destra in corsivo; `010` "Raccomandata a.r. anticipata a mezzo PEC" a sinistra;
`007` "A mezzo Raccomandata A/R…" giustificato). È un **punto di variazione**
(cfr. §7).

## 2. Eredità dal template `assets/Template_Vuoto.docx`

Il formatter **apre il template e ne preserva integralmente**:

- **Formato pagina**: A4 (21 × 29.7 cm), verticale.
- **Margini**: sinistro/destro 2.0 cm, **superiore 5.91 cm**, inferiore 4.54 cm.
  Il margine superiore ampio serve a lasciare spazio all'intestazione grafica.
- **Header (letterhead)**: "BERGAMO LEGAL — Società tra Avvocati s.r.l.",
  `www.bergamo.legal`, nomi degli avvocati. Font **Garamond** 14–16 pt. Include
  un'immagine logo (`word/media/image1.jpg`).
- **Footer**: dati di registrazione dello Studio (R.E.A., capitale, C.F./P.IVA,
  Ordine, recapiti, sedi di Bergamo e Milano). Font Garamond 10 pt.

Il corpo del documento viene **svuotato** (l'unico paragrafo vuoto presente)
e ricomposto; `header`, `footer`, `sectPr` (margini/pagina) **non si toccano**.

> **Regola critica (HARD_001 / STYLE_006)** — *Il letterhead non si duplica nel
> corpo.* Negli input storici le prime righe ripetono "BERGAMO LEGAL / Società
> tra Avvocati / indirizzo" e in coda i dati di footer: sono **omessi** perché
> già presenti nel template (cfr. `report_003`, `verification_028`). Il
> composer è content-driven: non emette mai il letterhead nel body.

## 3. Regole tipografiche globali (HARD)

| Regola | Dettaglio | Evidenza |
|--------|-----------|----------|
| Font corpo | **Times New Roman** su **ogni** run, esplicito (HARD_002). Il Garamond resta solo nell'header del template. | tutti gli output |
| Size esplicita | Ogni run ha la dimensione esplicita (HARD_003). Corpo 12 pt. | tutti |
| Giustificazione corpo | I paragrafi di corpo sono **JUSTIFY** 12 pt. | 001,003,007,025,028 |
| Allineamenti per blocco | destinatario **LEFT**; data/congedo/firma **RIGHT**; titoli rituali **CENTER**. | tutti |
| Normalizzazione trattini | en dash `–` (U+2013) ed em dash `—` (U+2014) → trattino ASCII `-` (HARD_004 / HARD_004B). | report 001,003,028 |
| Divisori decorativi vietati | sequenze `***` / `* * *` rimosse; la gerarchia è data dalla spaziatura (STYLE_011). | report 001, verification 028 |
| Data e firma separate | luogo+data e firma in **paragrafi distinti** (HARD_005). | tutti |
| Congedo lato firma | la formula di congedo è allineata **a destra**, come la firma (HARD_005). | 003,010,013,028,029 |
| Nomi firmatari in bold | i nomi dei firmatari sono **bold** anche se non lo erano nel sorgente (STYLE_016). | 002,007,028 |
| Titoli `keep_with_next` | i titoli di sezione non restano orfani a fine pagina. | report 001 |
| Placeholder preservati | `[DA INSERIRE: …]`, `[DATA INVIO]` ecc. **non** vengono alterati (HARD_007); però attivano `needs_review` (§6). | 001,007,010,028 |

### Dimensioni e spaziature per tipo di blocco

Valori derivati dagli output (spaziature in pt, indent in cm):

| Blocco | Font/Size | Allineamento | Indent sx | Spazio prima / dopo |
|--------|-----------|--------------|-----------|---------------------|
| delivery_method (autonomo) | TNR 12, spesso *italic* | RIGHT o LEFT | — | sa 8 |
| date_place | TNR 12 (talora bold) | RIGHT | — | sb 6 / sa 10 |
| recipient_block | TNR 12 | LEFT | **8.5 cm** (STYLE_008B) | sa 2 per riga |
| subject — label "OGGETTO:" | TNR **16** bold | CENTER (o inline JUSTIFY) | — | sb 10 / sa 4 |
| subject — contenuto | TNR 12 **bold** | CENTER o JUSTIFY | — | sa 8 |
| opening | TNR 12 (bold nelle formali) | JUSTIFY o LEFT | — | sb 8 / sa 6 |
| corpo: paragrafo | TNR 12 | JUSTIFY | — | sa 6 |
| corpo: titolo rituale | TNR **16** bold | CENTER | — | sb 14 / sa 8 |
| corpo: sottotitolo | TNR **14** bold | LEFT | — | sb 12 / sa 6 |
| corpo: voce di elenco | TNR 12 | JUSTIFY | **0.5 cm** | sa 4 |
| closing | TNR 12 (talora *italic*) | RIGHT | — | sb 16 / sa 6 |
| signature: nome | TNR 12 **bold** | RIGHT | — | sb 8 / sa 2 |
| signature: ruolo/studio | TNR 12 | RIGHT | — | sa 2 |
| signature: riga firma `____` | TNR 12 | RIGHT | — | sa 6 |
| attachments: "Allegati:" | TNR 12 bold | LEFT | — | sb 16 / sa 4 |
| attachments: voce | TNR 12 | LEFT | 0.5 cm | sa 3 |
| postscript/disclaimer | TNR **10** *italic* | RIGHT | — | sb 4 |

## 4. Dettaglio dei blocchi semantici

### 4.1 `delivery_method` (opzionale)
Metodo di trasmissione. Formule osservate: "A mezzo PEC", "Trasmessa a mezzo
PEC" (`003`,`029` — RIGHT corsivo), "Via PEC" (`025`,`027` — prima riga del
destinatario), "Raccomandata a.r. anticipata a mezzo PEC" (`010`), "A mezzo
Raccomandata A/R, anticipata a mezzo posta elettronica" (`007`,`028`),
"Tramite servizio online UIBM / PEC a:" (`004`).

### 4.2 `date_place`
"Bergamo, lì [DATA INVIO]" / "Bergamo, 5 giugno 2026" / "Madone, lì …". Allineato
a **destra**. Compare sopra il destinatario (`003`,`004`,`013`,`025`,`027`,`029`)
oppure subito dopo (`001`,`010`).

### 4.3 `recipient_block` (STYLE_008B)
Elenco di righe, **LEFT con rientro sinistro 8.5 cm**. Tipicamente: appellativo
("Spett.le", "Egr.", "Preg.ma Sig.ra"), denominazione (spesso **bold**), indirizzo
(tondo), riga PEC/e-mail (spesso bold). Esempi: `003` (Papi Solutions),
`025`/`027` (Trustpilot A/S), `028` (Sig.ra Locafaro).

> **Feedback cliente — 028 (4/5)**: un appellativo di cortesia *personale*
> ("Egr.", "Preg.ma", "Gent.ma", "Ill.mo", …) **non** deve stare su una riga a
> sé sopra il nome, ma sulla **stessa riga** ("mono linea"): `Egr. Sig.ra Simona
> Locafaro`, non `Egr.` + `Sig.ra Simona Locafaro`. Fa eccezione **"Spett.le"**
> davanti a una **denominazione societaria**, che resta su riga propria (forma
> approvata 5/5 in `003` e `025`). Il formatter applica l'unione automatica
> (`merge_courtesy_appellative=True`).

### 4.4 `subject` — OGGETTO (STYLE_008)
Due varianti, entrambe presenti nel corpus:
- **inline** (default): "OGGETTO: <testo>" in un unico paragrafo **JUSTIFY**, run
  label 16 pt bold + run contenuto 12 pt bold (`001`,`005`,`025`,`029`,`032`);
- **split**: "OGGETTO:" CENTER 16 pt bold (con `keep_with_next`) seguito da un
  paragrafo contenuto 12 pt bold (`003`,`007`,`010`,`013`,`027`,`028`).

Regola stabile: **label 16 pt bold, contenuto 12 pt bold**.

> **Feedback cliente — 027 (4/5)**: il cliente chiede esplicitamente
> «*oggetto: testo giustificato*» in un **unico paragrafo**, e contesta la forma
> *split* + centrata usata in `output_027`. La forma inline+giustificata è
> confermata dal riferimento **5/5 `025` ("dritto dritto")**, che usa proprio
> `OGGETTO:` inline JUSTIFY. → **Default = inline, giustificato.** Lo *split*
> resta disponibile come opzione.

### 4.5 `opening`
Formula di apertura: "Egregio/a Consigliere/a,", "Spett.le Società,", "Egregia
Signora,", "Ill.mo Sig. Legale Rappresentante,", "Egregi Signori,", "Gentile
Dott.ssa Corna,", "Gentile Mariachiara,". **Bold + JUSTIFY** nelle lettere
formali (`001`,`007`,`025`,`027`,`028`); **tondo + LEFT** nelle comunicazioni
stile e-mail (`005`,`011`,`032`).

### 4.6 `body_blocks` — il corpo
Lista ordinata di blocchi tipizzati. Tipi (cfr. §3 per gli stili):

- **paragraph** — paragrafo discorsivo, JUSTIFY 12 pt. Numerazione manuale
  ammessa nel testo ("1. …", "5. …").
- **heading_center** — titolo rituale centrato 16 pt bold: "IN FATTO",
  "IN DIRITTO" (`001`,`002`), "PREMESSO CHE", "CONSIDERATO CHE" (`010`,`015`),
  "DICHIARA" (`004`,`015`), "DIFFIDA" (`007`,`028`), "INVITA" / "INVITA LA S.V."
  (`001`,`015`), "D I F F I D A  E  I N T I M A" (`010`). `keep_with_next`.
- **heading_left** (*titoletto*) — sottotitolo a sinistra **14 pt bold**, con il
  **contenuto a capo** (riga successiva): "Fatto", "Diritto" (`007`,`028`),
  "MOTIVA il ritiro come segue." (`004`), titoli numerati di argomento
  "1. Pagina Trustpilot duplicata…" (`025`,`027`). `keep_with_next`.
- **list_item** — voce di elenco JUSTIFY con rientro sinistro (default 0.5 cm).
  Marcatori osservati: "(i)/(ii)/(iii)", "i)/ii)", "(a)/(b)", "(α)/(β)", "•".
  Il marcatore fa parte del testo della voce. Rientro variabile 0.5–1.0 cm.

> **Due livelli di titolo, da non confondere (regola cliente):**
> - **titoletti** → sempre **14 pt, a SINISTRA**, con il **contenuto a capo**
>   (es. `Fatto`, `Diritto`, `MOTIVA`, `CHIEDE ALTRESÌ`, titoli numerati);
> - **intestazioni rituali e label OGGETTO** → **16 pt** (le rituali centrate:
>   `IN FATTO`, `IN DIRITTO`, `DICHIARA`, `DIFFIDA`, `PREMESSO CHE`, `INVITA`…).
>
> **Verbi dispositivi con testo a seguire** (es. `CHIEDE ALTRESÌ, ai sensi…`):
> il verbo è un *titoletto* 14 pt a sinistra e il testo che introduce va **a
> capo** come paragrafo giustificato — **non** in grassetto inline sulla stessa
> riga. Il formatter offre l'helper `disposition(verbo, contenuto)` che produce
> esattamente questa coppia titoletto+paragrafo, così la resa è coerente con gli
> altri titoletti (es. `MOTIVA`). Il **grassetto leggero** nel corpo resta
> ammesso per enfasi su termini chiave, ma non per i verbi dispositivi.

### 4.7 `closing`
"Distinti saluti.", "Con osservanza," (*italic*), "Cordiali saluti,", "In fede,",
"In attesa di un cortese… si porgono distinti saluti.", "Nell'auspicio… distinti
saluti." → **RIGHT** 12 pt.

### 4.8 `signature_block` (HARD_005 / STYLE_016)
Righe allineate a **destra**: eventuale denominazione Studio (tondo o bold),
**nome/i firmatario/i in bold**, ruolo (tondo), riga di firma "____" (tondo),
nota "(Firma digitale)" (*italic*). Più firmatari → ciascuno con la propria riga
di firma (`002`,`015`).

### 4.9 `attachments` (opzionale)
"Allegati:" LEFT bold, poi una voce per allegato "All. N - …" LEFT, rientro
0.5 cm. Unico caso nel corpus: `007`.

### 4.10 `postscript` (opzionale)
Nota finale/disclaimer in coda, 10 pt *italic* a destra (`001`: "(trasmessa per
il tramite dei difensori…)").

## 5. Regole di normalizzazione del testo

Applicate a **ogni** run emesso:
1. `–` (U+2013) → `-`; `—` (U+2014) → `-` (HARD_004B).
2. Rimozione di righe/sequenze divisorie `***` / `* * *` (STYLE_011).
3. Gli spazi non significativi a fine riga vengono compattati; i placeholder
   `[…]` restano **invariati** (HARD_007).

## 6. Casi che impostano `needs_review`

Il formatter segnala (non blocca) quando:
- restano **placeholder** non risolti: `[DA INSERIRE…]`, `[DATA…]`, `[…]`;
- la **data è assente**: lo script scrive comunque `[INSERISCI QUI LA DATA]`
  nella posizione della data e segnala l'omissione;
- un blocco **obbligatorio** è vuoto: destinatario, oggetto o firma assenti;
- residua un **en/em dash** dopo la normalizzazione (indice di bug a monte);
- il contenuto contiene marcatori di divisore decorativo non rimossi.

Questi controlli sono di **forma**, non entrano nel merito giuridico
(`docs/letter_formatter_goal.md`: nessuna modifica di contenuto legale).

## 7. Punti di variazione (parametrizzabili)

| Variabile | Valori osservati | Default scelto |
|-----------|------------------|----------------|
| Posizione `delivery_method` | riga autonoma alto / prima riga destinatario / dopo data | riga autonoma |
| Posizione `date_place` | sopra o sotto il destinatario | sopra il destinatario |
| OGGETTO | inline vs split; CENTER vs JUSTIFY | **inline, JUSTIFY** (feedback 027/025) |
| Appellativo personale | riga a sé vs mono-linea col nome | **mono-linea** (feedback 028) |
| `opening` | bold+JUSTIFY (formale) vs tondo+LEFT (e-mail) | bold + JUSTIFY |
| Rientro `list_item` | 0.5 / 0.7 / 1.0 cm | 0.5 cm |

## 8. Incertezze — stato dopo il feedback cliente (`feedback.md`)

### Risolte dal feedback
- ✅ **Forma dell'OGGETTO**: era incerta (split/center vs inline/justify); il
  feedback `027` (4/5, richiesta esplicita di «oggetto: testo giustificato») e
  il riferimento `025` (5/5) la fissano su **inline + JUSTIFY**. Default
  aggiornato.
- ✅ **Appellativo personale nel destinatario**: il feedback `028` (4/5) chiede
  la **mono-linea** (`Egr. Sig.ra Nome`). Default aggiornato (con eccezione per
  "Spett.le" davanti a società).

### Ancora aperte (da validare con revisione umana)
- **Posizione di data e metodo di trasmissione**: entrambe le disposizioni sono
  attestate; nessun feedback specifico sulle lettere le disambigua (il rilievo
  `024` su data/scriventi a destra riguarda gli `acts`, fuori ambito). Default
  sul caso prevalente, da confermare per famiglia.
- **Bold del saluto di apertura**: dipende dal registro (formale vs e-mail);
  non sempre desumibile automaticamente dal contenuto. Nessun feedback dirimente.
- **Rientro delle voci di elenco** (0.5/0.7/1.0 cm): nessun feedback; default 0.5.
- **Marcatori di elenco**: il marcatore è considerato testo; non viene generato
  automaticamente, resta responsabilità del contenuto fornito.

Queste scelte restano **default espliciti e sovrascrivibili**: un futuro
sviluppatore cambia il contenuto dei blocchi (e, se serve, i pochi flag di
variazione) senza riscrivere la logica di formattazione.

> Nota di scope: i feedback `024`, `026`, `030` riguardano `acts` e, come indicato
> in `feedback.md`, sono conservati come evidenza futura ma **non** guidano il
> formatter delle lettere. Il caso `009` (`other_pending_name`) è ambiguo e non
> prova un errore sull'output associato: ignorato.
