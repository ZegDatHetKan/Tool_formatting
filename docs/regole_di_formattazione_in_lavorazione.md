# Regole di formattazione (in lavorazione) — specchio di `formatters/letters.py`

Questo documento è lo **specchio in prosa** di ciò che fa lo script
`formatters/letters.py`: ogni valore qui (font, dimensioni, allineamenti,
rientri, spaziature, condizioni) corrisponde **1:1** a una riga di codice.

> **Patto di sincronizzazione.** Il codice è la **fonte di verità**
> (cfr. `regole_formattazione_lettere.md`, §"Fonte di verità"). Quando si cambia
> uno dei due, si aggiorna **subito** anche l'altro: se modifichi una costante o
> un `_emit_*` in `letters.py`, riporta il nuovo valore qui; se qui scrivi una
> regola nuova, deve già esistere nel codice. Questo file descrive il
> comportamento *attuale* dello script, non l'analisi del corpus (per quella
> vedi `regole_formattazione_lettere.md`).
>
> Spaziature in **punti (pt)**, rientri in **cm**. "honor span" = rispetta il
> grassetto/corsivo dei singoli `Span` passati; un valore esplicito (es.
> `bold=True`) invece forza tutta la riga.

## 1. Costanti di stile (`letters.py`, sezione "Style constants")

| Costante | Valore | Uso |
|----------|--------|-----|
| `BODY_FONT` | `Times New Roman` | impostato esplicitamente su **ogni** run (attributi `w:ascii`, `w:hAnsi`, `w:cs`) |
| `SIZE_BODY` | `12` pt | corpo, destinatario, data, congedo, firma, contenuto oggetto |
| `SIZE_HEADING` | `16` pt | label `OGGETTO:` + intestazioni rituali (centrate) |
| `SIZE_SUBHEADING` | `14` pt | titoletti a sinistra |
| `SIZE_SMALL` | `10` pt | postscript / disclaimer |
| `RECIPIENT_INDENT_CM` | `8.5` cm | rientro sinistro blocco destinatario |
| `LIST_INDENT_CM` | `0.5` cm | rientro sinistro voci di elenco e allegati |
| `DEFAULT_TEMPLATE` | `assets/Template_Vuoto.docx` | template di base |

## 2. Eredità dal template (`render_letter` → `_clear_body`)

Si apre `assets/Template_Vuoto.docx`, si **rimuovono solo i paragrafi del corpo**
e si ricompone. Restano intatti: `sectPr` (formato pagina A4, margini — superiore
≈5.91 cm), **header** (letterhead "BERGAMO LEGAL" in Garamond + logo), **footer**
(dati Studio). Il corpo è in Times New Roman; il Garamond resta solo nell'header.

## 3. Ordine di emissione (`render_letter`)

**Principio (regola cliente): si formatta SOLO ciò che viene fornito, MAI si
inventa o completa.** Ogni blocco è opzionale e viene reso solo se valorizzato;
nessun placeholder/destinatario/saluto/data viene aggiunto dal formatter.

1. `delivery_method` — solo se valorizzato **e** `delivery_inline_with_recipient == False`;
2. `date_place` — **solo se fornita** (nessun placeholder se manca); sopra il destinatario se `date_above_recipient == True`;
3. `recipient_block` — solo se non vuoto;
4. `date_place` — se `date_above_recipient == False`, la data va qui invece che al punto 2;
5. `subject` — solo se non vuoto;
6. `opening` — solo se non vuoto;
7. `body_blocks`;
8. `closing` — solo se non vuoto;
9. `signature_block` — solo se non vuoto;
10. `attachments` — solo se non vuoto;
11. `postscript` — solo se non vuoto.

La cartella di output viene creata se non esiste; il file non finisce mai in
`previous_works/`.

## 4. Stile per blocco (valori esatti, da `_emit_*`)

| Blocco / funzione | Allin. | Size | Grassetto | Corsivo | Rientro sx | Sp. prima | Sp. dopo | keep_with_next |
|-------------------|--------|------|-----------|---------|------------|-----------|----------|----------------|
| `delivery_method` (autonomo) — `_emit_delivery` | LEFT | 12 | honor span | **sì (forzato)** | — | — | 8 | no |
| `date_place` — `_emit_date` (solo se fornita) | RIGHT | 12 | honor span | honor span | — | 6 | 10 | no |
| `recipient_block` riga — `_emit_recipient` | LEFT | 12 | honor span | honor span | **8.5** | — | 2 (8 sull'ultima riga) | no |
| OGGETTO **split** label (default) — `_emit_subject` | CENTER³ | 16 | **sì** | no | — | 10 | 4 | **sì** |
| OGGETTO **split** contenuto (default) — `_emit_subject` | JUSTIFY³ | 12 | **sì** | no | — | — | 8 | no |
| OGGETTO **inline** (opzione) — `_emit_subject` | JUSTIFY¹ | label 16 / testo 12 | **sì** | no² | — | 8 | 8 | no |
| `opening` — `_emit_opening` | JUSTIFY⁴ | 12 | sì se formale, altrim. honor⁴ | honor span | — | 8 | 6 | no |
| corpo: **paragraph** — `_emit_body` | JUSTIFY | 12 | honor block | honor block | — | — | 6 | no |
| corpo: **heading_center** (rituale) — `_emit_body` | CENTER | 16 | **sì** | no | — | 14 | 8 | **sì** |
| corpo: **heading_left** (titoletto) — `_emit_body` | LEFT | 14 | **sì** | no | — | 12 | 6 | **sì** |
| corpo: **list_item** — `_emit_body` | JUSTIFY | 12 | honor block | honor block | block o 0.5⁵ | — | 4 | no |
| `closing` — `_emit_closing` | RIGHT | 12 | honor span | honor span | — | 16 | 6 | no |
| firma: riga — `_emit_signature` | RIGHT | 12 | vedi §5.3 | vedi §5.3 | — | vedi §5.3⁶ | 2 (8 dopo riga firma) | no |
| `attachments` "Allegati:" — `_emit_attachments` | LEFT | 12 | **sì** | no | — | 16 | 4 | no |
| `attachments` voce — `_emit_attachments` | LEFT | 12 | honor span | honor span | **0.5** | — | 3 | no |
| `postscript` — `_emit_postscript` | RIGHT | 10 | honor span | **sì (forzato)** | — | 4 | — | no |

Note:
1. CENTER se `subject_content_center == True`.
2. La **label** è ciò che precede il primo `:` (es. `OGGETTO:`, `Oggetto:`,
   `Subject:`): preservata verbatim a 16 pt, il resto a 12 pt. **Nessuna label
   viene inventata**: se il testo non ha `:`, tutto è reso a 12 pt grassetto.
3. Default split: **label CENTER** (`subject_label_center == True`) e
   **contenuto JUSTIFY** (`subject_content_center == False`). Invertendo i flag:
   label JUSTIFY se `False`, contenuto CENTER se `True`.
4. LEFT e grassetto-honor se `opening_formal == False` (registro e-mail);
   JUSTIFY + grassetto forzato se `opening_formal == True` (registro formale).
5. `block.indent_cm` se specificato sulla singola voce, altrimenti
   `letter.list_indent_cm` (default 0.5 cm).
6. `space_before`: **18** pt sulla prima riga del blocco (stacco dal corpo);
   **2** pt sulla riga di sola firma `____` (aderisce al nome); **12** pt su un
   nome che segue una riga di firma (nuovo firmatario → aria); **2** pt sulle
   righe ruolo/denominazione (restano strette sotto il nome).

## 5. Regole speciali

### 5.1 Normalizzazione del testo (`normalize_text`)
Su **ogni** run: en dash `–` (U+2013) → `-`; em dash `—` (U+2014) → `-`. I
placeholder `[...]` non vengono toccati.

### 5.2 Divisori decorativi (`_DIVIDER_RE`, in `_emit_body`)
Un blocco di corpo il cui testo combacia con `^\s*(?:\*\s*){2,}\*?\s*$`
(es. `***`, `* * *`) viene **scartato**.

### 5.3 Riconoscimento righe di firma (`_emit_signature`)
Tipi di riga e relativo grassetto/corsivo:
- riga di **sola firma** (solo caratteri `_`): `bold=False`;
- riga "(Firma digitale)" (contiene "firma digitale", case-insensitive):
  `italic=True`, `bold=False`;
- ogni altra riga: `bold` honor-span (i nomi firmatari si passano in grassetto).

Spaziatura (D3 — blocco arioso ed elegante): prima riga `space_before=18`
(stacco); riga di firma `____` `space_before=2`, `space_after=8` (separa più
firmatari); un nome che **segue** una riga di firma `space_before=12` (nuovo
firmatario); righe ruolo/denominazione `space_before=2` (strette sotto il nome).

### 5.4 Unione appellativo di cortesia (`_merge_appellatives`)
Attiva con `merge_courtesy_appellative == True` (default). Se una riga del
destinatario, normalizzata (strip + lower), coincide con un appellativo
**personale** dell'insieme `_COURTESY_APPELLATIVES` ed esiste una riga
successiva, le due righe vengono **unite** in una sola ("Egr. " + nome).

Insieme `_COURTESY_APPELLATIVES`: `egr.`, `egr.mo`, `egr.ma`, `egregio`,
`egregia`, `preg.mo`, `preg.ma`, `pregiatissimo`, `pregiatissima`, `gent.mo`,
`gent.ma`, `gentile`, `gentilissimo`, `gentilissima`, `ill.mo`, `ill.ma`,
`illustrissimo`, `illustrissima`, `chiar.mo`, `chiar.ma`.
**Esclusa di proposito**: `Spett.le` (le denominazioni societarie restano su
riga propria).

### 5.5 Verbi dispositivi (`disposition`)
`disposition(verbo, contenuto)` produce **due blocchi**: un `heading_left`
(titoletto 14 pt sinistra, vedi §4) con il verbo + un `paragraph` (giustificato
12 pt) con il contenuto **a capo**. Da usare per `CHIEDE ALTRESÌ`, `MOTIVA` ecc.
quando trascinano testo sulla stessa riga; mai grassetto inline. Le liste di
blocchi annidate (prodotte da `disposition`) sono appiattite da `_flatten_blocks`.

## 6. Default dei punti di variazione (`LetterDocument`)

| Campo | Default | Effetto |
|-------|---------|---------|
| `date_above_recipient` | `True` | data sopra il destinatario |
| `delivery_inline_with_recipient` | `False` | metodo di trasmissione come riga autonoma in alto |
| `subject_style` | `SPLIT` | `OGGETTO:` su riga propria + testo a capo (feedback 027) |
| `subject_label_center` | `True` | label `OGGETTO:` **centrata** 16 pt |
| `subject_content_center` | `False` | contenuto **giustificato** 12 pt |
| `opening_formal` | `True` | apertura grassetto + JUSTIFY |
| `list_indent_cm` | `0.5` | rientro voci di elenco |
| `merge_courtesy_appellative` | `True` | unione appellativo personale + nome |

## 7. Validazione → `needs_review` (`collect_warnings`)

Controlli **di forma** (non di merito giuridico) sul **solo testo fornito**: un
blocco mancante NON è un errore (non si completa nulla), quindi non viene
segnalato. `RenderResult.needs_review` è `True` se la lista non è vuota:
- placeholder non risolti **lasciati dall'autore** che combaciano con `[... DA INSERIRE | DATA | INSERIRE | ... | … ...]` (case-insensitive);
- en/em dash residuo **dopo** normalizzazione (difensivo: indica una variante non coperta da `normalize_text`);
- divisore decorativo `***` presente (verrà comunque rimosso in resa).

## 8. Superficie API (`__all__`)

- `LetterDocument`, `RenderResult`, `Span`, `BlockKind`, `SubjectStyle`;
- costruttori blocchi: `paragraph`, `section_heading`, `subheading`,
  `list_item`, `disposition`;
- funzioni: `normalize_text`, `collect_warnings`, `render_letter`.

`render_letter(letter, template_path=DEFAULT_TEMPLATE, output_path="out/letter.docx")`
→ `RenderResult(output_path, warnings, paragraph_count, needs_review)`.

## 9. Mappa sezione ↔ codice (per tenere lo specchio allineato)

| Sezione di questo doc | Simbolo in `formatters/letters.py` |
|-----------------------|------------------------------------|
| §1 Costanti | blocco "Style constants" |
| §2 Template | `_clear_body`, `render_letter` |
| §3 Ordine | `render_letter` |
| §4 Stile per blocco | `_emit_delivery/_date/_recipient/_subject/_opening/_body/_closing/_signature/_attachments/_postscript`, `_add_paragraph` |
| §5.1 Normalizzazione | `normalize_text` |
| §5.2 Divisori | `_DIVIDER_RE`, `_emit_body` |
| §5.3 Firma | `_emit_signature` |
| §5.4 Appellativi | `_merge_appellatives`, `_COURTESY_APPELLATIVES` |
| §5.5 Dispositivi | `disposition`, `_flatten_blocks` |
| §6 Default | campi di `LetterDocument` |
| §7 Validazione | `collect_warnings`, `RenderResult` |
| §8 API | `__all__` |
