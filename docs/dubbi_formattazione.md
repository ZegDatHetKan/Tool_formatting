# Dubbi di formattazione — confronto ricostruzione vs output storici

## Metodo

Ho **ricostruito tutte le 15 lettere** del corpus partendo **solo dagli input
grezzi** (`input_*.docx`), mappandole nello schema `LetterDocument` e rendendole
**con il nostro formatter** — senza guardare gli output storici. Poi ho
confrontato la mia resa con `output_*.docx`. Le ricostruzioni stanno in
`documenti_generati/storiche/` (non versionate: contengono dati cliente).

## Esito sintetico

**Buona notizia:** lo schema riproduce **fedelmente** la struttura del corpus.
Le uniche differenze sistematiche sono **3 cambi voluti** (feedback/tue
istruzioni) e **1 lacuna** dello schema (comunicazioni email/PEC). Tutto il
resto (destinatario a 8,5 cm, titoli rituali 16 centrati, titoletti 14 sx, firma
a destra, elenchi, normalizzazione) **coincide**.

### Tabella per lettera (paragrafi mia/storica; OGGETTO mia/storica)

| ID | par mia/hist | OGGETTO mia | OGGETTO storico | Note |
|----|-------------|-------------|-----------------|------|
| 001 | 32/33 | J inline | **J inline** | coincide |
| 002 | 29/30 | J inline | **J inline** | coincide |
| 003 | 16/17 | J inline | **C split** | OGGETTO diverso |
| 004 | 21/20 | J inline | C inline | OGGETTO + CHIEDE titoletto |
| 005 | 22/21 | J inline | L (sx) | email-style (vedi D2) |
| 007 | 47/48 | J inline | **C split** | OGGETTO diverso |
| 010 | 32/33 | J inline | **C split** | OGGETTO diverso |
| 011 | 10/8 | J inline | C (Subject) | email-style (vedi D2) |
| 013 | 15/17 | J inline | **C split** | OGGETTO diverso |
| 015 | 38/40 | J inline | **C split** | OGGETTO diverso |
| 025 | 36/35 | J inline | **J inline** | coincide |
| 027 | 46/47 | J inline | **C split** | OGGETTO diverso |
| 028 | 45/47 | J inline | **C split** | OGGETTO + CHIEDE titoletto |
| 029 | 15/15 | J inline | **J inline** | coincide |
| 032 | 21/19 | J inline | C inline | email-style (vedi D2) |

Conteggio OGGETTO nello **storico**: inline-giustificato 4 (001, 002, 025, 029);
**centrato** 9 (003, 004, 007, 010, 013, 015, 027, 028, 032); a sinistra 1 (005).
→ Il corpus storico era **incoerente**; il nostro formatter standardizza tutto a
**inline-giustificato**.

## Differenze VOLUTE (non sono regressioni)

1. **OGGETTO inline-giustificato** invece di centrato/split. Deciso dopo il
   **feedback 027** (4/5: «*oggetto: testo giustificato*», contestava il
   centrato/split) e il **5/5 di 025** (inline). Supportato anche dal feedback
   generale `009` («spesso fa Oggetto: a capo» = lamentela sul form *split*).
2. **Verbi dispositivi come titoletti** (`CHIEDE ALTRESÌ` → 14 sx + contenuto a
   capo) invece di grassetto inline. Tua istruzione esplicita.
3. **Data sempre presente** con placeholder `[INSERISCI QUI LA DATA]` se assente,
   segnalata dal review. Tua istruzione esplicita.

## DUBBI APERTI (il feedback non li risolve del tutto)

> **D1 — RISOLTO e implementato.** Resa decisa col cliente: etichetta `OGGETTO:`
> **centrata 16 pt grassetto**, poi **a capo** il testo dell'oggetto a sinistra/
> **giustificato 12 pt grassetto** (forma *split*). Applicata come default a
> tutte le famiglie; le 4 lettere storiche che usavano l'inline sono uniformate.
> L'etichetta è lasciata come scritta dall'autore (nessuna forzatura maiuscolo).

### D1 — OGGETTO inline-giustificato per TUTTE le famiglie di lettera?

- **Dove sta il feedback:** `feedback.md`, riga
  `Lettera_Trustpilot_v3_collaborativa.docx` = corpus **027**, voto **4/5**.
- **Feedback verbatim del cliente:** «*la parte in alto l'ha spezzettata un po'
  come voleva invece di farla sensata e Ha messo la parte oggetto con oggetto e
  poi sotto il resto e centrale invece di oggetto: testo giustificato*».
- **Decodifica (parole poco chiare):** «ha messo la parte oggetto con oggetto e
  poi sotto il resto» = ha messo `OGGETTO:` su una riga e il contenuto sulla
  riga **sotto** (forma *split*, due righe); «e centrale» = centrato; «**invece
  di** oggetto: testo giustificato» = mentre lui voleva `Oggetto: testo` su
  **una sola riga, giustificato**. Quindi: split+centrato = ciò che NON voleva;
  inline+giustificato = ciò che voleva.
- **Risposta alla tua domanda** ("quando usava oggetto centrato, il resto era
  inline?"): in `027` **no, era split** (label centrata su una riga + contenuto
  centrato sotto). In 2 casi del corpus (`004`, `032`) era invece centrato ma su
  **una** riga. La lamentela era sulla forma *split* centrata.
- **Cosa fa ora lo script:** OGGETTO sempre inline + giustificato, label 16 /
  testo 12 (entrambi grassetto) = **esattamente la forma che il cliente voleva**.
- **Cosa dice il feedback:** lo risolve **solo per i Trustpilot** (`027` 4/5
  voleva inline; `025` 5/5 inline). Segnale generale a favore dell'inline anche
  da `009`. **MA**: la PEC `003` era **5/5 con OGGETTO centrato/split**, e le
  diffide/recessi storici (`007`,`010`,`013`,`015`,`028`) usavano il **centrato**
  senza che risultino lamentele sull'oggetto.
- **Conflitto:** stessa label resa centrata = 5/5 su `003`, ma 4/5 (contestata)
  su `027`.
- **❓ Domanda:** confermo **inline-giustificato per tutte** le lettere, oppure
  alcune famiglie (diffide / PEC / istanze) devono mantenere l'**OGGETTO
  centrato 16**? Se sì, quali?
- **Risposta cliente:** ______________________________________________

> **D2 — RISOLTO e implementato.** Decisione cliente: «si formatta solo ciò che
> viene mandato, MAI completare». Lo script ora: non forza la data (nessun
> placeholder); destinatario, saluto, firma, oggetto sono tutti **opzionali**
> (resi solo se forniti); l'OGGETTO non viene mai anteposto se assente. Esiste
> di fatto una "modalità email/PEC" (basta non fornire destinatario né data).
>
> **D3 — RISOLTO e implementato.** Blocco firma reso più arioso: stacco dal
> corpo (18 pt), riga `____` aderente al nome con stacco sotto, nuovo firmatario
> con aria, righe ruolo strette.

### D2 — Manca una "modalità email/PEC" (lettere 005, 011, 032)

- **Cosa ho osservato:** le comunicazioni stile e-mail **non hanno** né blocco
  destinatario postale né data. La `011` con questo form minimale (solo
  Oggetto + saluto + corpo + firma) era **5/5**.
- **Cosa fa ora lo script:** `recipient_block` è **obbligatorio** e la **data è
  forzata** → nella ricostruzione l'agente ha dovuto **inventare un destinatario**
  (duplicando il saluto, es. "Gentile Stefania") e compare un `[INSERISCI QUI LA
  DATA]` che nello storico non c'era. Inoltre l'oggetto storico è "Subject:" /
  "Oggetto:" **centrato**, non "OGGETTO:" giustificato.
- **Cosa dice il feedback:** nulla di esplicito, ma il **5/5 di `011`** con il
  form minimale suggerisce che destinatario e data **non vanno forzati** in
  questa modalità.
- **❓ Domanda:** aggiungo una **modalità "corpo email/PEC"** in cui:
  (a) il destinatario è opzionale; (b) la data non è forzata (niente placeholder);
  (c) l'oggetto può restare "Oggetto:/Subject:" ed essere eventualmente centrato?
- **Risposta cliente:** ______________________________________________

### D3 (minore) — Segmentazione del blocco firma

- In alcune lettere il numero di righe a destra (firma) differisce di ±1 (come
  sono divise denominazione Studio / nome / riga di firma `____`). Non è una
  regola di stile: dipende da come si passano le righe in `signature_block`.
- **❓ Domanda:** c'è una regola fissa per il blocco firma (es. sempre:
  denominazione Studio + nome avvocato in grassetto + riga `____`), o si lascia
  libero al contenuto? Probabilmente non serve cambiare nulla.
- **Risposta cliente:** ______________________________________________

## Nota

D1, D2 e D3 sono stati **decisi e implementati nel formatter** (vedi i riquadri
in testa a ciascun dubbio). Le lettere si rigenerano col tool su questi default.
