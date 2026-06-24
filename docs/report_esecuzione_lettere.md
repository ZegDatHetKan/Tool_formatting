# Report di esecuzione — Letters formatter

Data: 2026-06-24

## Cosa è stato consegnato

| File | Ruolo |
|------|-------|
| `docs/regole_formattazione_lettere.md` | Regole di formattazione inferite dal corpus |
| `formatters/letters.py` | Composer deterministico (python-docx), content-driven |
| `formatters/__init__.py` | Package |
| `validate_letters.py` | Harness di validazione + 4 lettere di esempio |
| `requirements.txt` | Dipendenza `python-docx` |
| `out/example_*.docx` | Output generati dalla validazione (fuori da `previous_works/`) |

## Lettere del corpus ispezionate

Fonte di classificazione: **`previous_works/manifest.json`** (unica). Selezionati
i 15 record con `document_type == "letters"`, tutti ispezionati a livello di
`output_*.docx` (layout di riferimento), con lettura dei relativi
`report_*.md`/`verification_*.md` e di alcuni `input_*.docx`:

`001, 002, 003, 004, 005, 007, 010, 011, 013, 015, 025, 027, 028, 029, 032`

Distribuzione del manifest: 15 `letters`, 14 `acts`, 3 `other_pending_name`
(questi ultimi due gruppi **ignorati**, come da consegna). Nessuna lettera ha
`needs_review: true` o note/issue nel manifest; tutte `classification_confidence:
high`.

## Regole confermate (alta evidenza nel corpus)

- **Eredità dal template**: A4, margine superiore 5.91 cm, header "BERGAMO LEGAL"
  in Garamond + logo, footer con dati Studio. Preservati integralmente; il corpo
  viene svuotato e ricomposto senza toccare `sectPr`/header/footer (verificato
  riaprendo gli output).
- **Letterhead non duplicato nel corpo** (le prime righe "BERGAMO LEGAL…" e il
  footer presenti negli input vengono omessi).
- **Times New Roman** con font+size **espliciti su ogni run** del corpo.
- Allineamenti per blocco: destinatario LEFT con **rientro 8.5 cm**;
  data/congedo/firma **RIGHT**; titoli rituali (IN FATTO, DIFFIDA, DICHIARA,
  PREMESSO CHE…) **CENTER 16 pt bold**; sottotitoli (Fatto, Diritto, titoli
  numerati) **LEFT 14 pt bold**; corpo **JUSTIFY 12 pt**.
- **OGGETTO**: label 16 pt bold, contenuto 12 pt bold.
- **Normalizzazione**: en/em dash → `-`; divisori `***` rimossi.
- **Firma**: congedo e firma a destra; nomi firmatari in bold; data e firma in
  paragrafi distinti; più firmatari ciascuno con riga di firma.
- **Elenchi**: voci con rientro sinistro (default 0.5 cm), il marcatore è testo.

## Come è stato validato

`.venv/bin/python validate_letters.py` (esito: **VALIDAZIONE OK**) verifica:

1. il template apre;
2. vengono creati 4 output in `out/` (mai in `previous_works/`);
3. `previous_works/` resta **byte-for-byte intatta** (confronto SHA-256 dei 129
   file prima/dopo);
4. **4 strutture rappresentative** rese tramite lo schema semantico:
   diffida (Fatto/Diritto + DIFFIDA + elenchi + 2 firmatari), comunicazione
   stile e-mail (oggetto inline, apertura non formale, elenco puntato),
   contestazione PEC (nota di trasmissione, oggetto split, firma Studio+avvocato),
   informativa (data sotto il destinatario, postscript, placeholder → `needs_review`);
5. ogni output, riaperto, conserva letterhead/footer/margini e rispetta le
   regole tipografiche (TNR su ogni run, nessun en/em dash, firma a destra).

Confronto strutturale a campione fra `out/example_diffida.docx` e gli output
storici (es. `028`/`007`): coincidono indent del destinatario (8.5 cm), split
OGGETTO 16/12 pt, sottotitoli 14 pt, titolo rituale 16 pt centrato, elenchi a
0.5 cm, blocco firma a destra con nomi in bold.

## Integrazione del feedback cliente (`feedback.md`)

Due dubbi sono stati **risolti** dal feedback (lettere `027` e `028`, indicati
in `feedback.md` come prioritari):

- ✅ **Forma OGGETTO** → da split+centrato a **inline + giustificato**
  (richiesta esplicita su `027`, confermata dal 5/5 di `025`). Cambiato il
  default `subject_style = INLINE`.
- ✅ **Appellativo personale del destinatario** → **mono-linea** `Egr. Sig.ra
  Nome` (feedback `028`), con eccezione per "Spett.le" davanti a società.
  Aggiunto `merge_courtesy_appellative=True`.

## Cosa resta incerto / da revisione umana

- **Posizione di data e metodo di trasmissione**: entrambe attestate; nessun
  feedback dirimente sulle lettere (il rilievo `024` è su `acts`). Default su
  caso prevalente, da confermare per famiglia.
- **Bold dell'apertura** (formale vs e-mail): dipende dal registro, non sempre
  inferibile dal contenuto.
- **Rientro delle voci di elenco** (0.5/0.7/1.0 cm): default 0.5 cm.
- **Placeholder** `[DA INSERIRE…]`/`[DATA…]`: conservati intatti ma segnalati via
  `needs_review`; richiedono compilazione umana prima dell'invio.
- Il merito **giuridico** non è toccato: il formatter è solo struttura/forma.
- Feedback su `acts` (`024`,`026`,`030`) e su `009`: fuori ambito lettere,
  conservati come evidenza ma non applicati.

## Uso

```python
from formatters.letters import LetterDocument, render_letter, paragraph, section_heading

letter = LetterDocument(
    delivery_method="A mezzo PEC",
    date_place="Bergamo, 24 giugno 2026",
    recipient_block=["Spett.le ...", "Via ..."],
    subject="OGGETTO: ...",
    opening="Egregi Signori,",
    body_blocks=[paragraph("..."), section_heading("DIFFIDA"), paragraph("...")],
    closing="Distinti saluti,",
    signature_block=["Avv. Matteo Bertocchi"],
)
result = render_letter(letter, "assets/Template_Vuoto.docx", "out/lettera.docx")
print(result.needs_review, result.warnings)
```
