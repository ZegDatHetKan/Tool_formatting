# Client Approval Points

The HTML review artifact must make these approval points visible. The goal is
to let the client say "approved" or give precise corrections.

## Page/template

- Header and footer come from `assets/Template_Vuoto.docx`.
- Body content must not duplicate the letterhead.
- A4 page proportions and large top margin are intentional.

## Letter opening area

- Delivery method placement and italic behavior.
- Date/place alignment and spacing.
- Recipient block indentation at 8.5 cm.
- Personal appellative mono-line rule, for example `Egr. Sig.ra Nome`.

## Subject

- `OGGETTO:` label is 16 pt bold.
- Current implementation uses split mode by default: label on its own line and
  subject content below.
- Subject content is 12 pt bold and justified.
- This is a known client-sensitive area because of feedback on Trustpilot v3.

## Body hierarchy

- Normal body paragraphs: Times New Roman 12 pt, justified.
- Left subheadings: 14 pt bold, left aligned, kept with following paragraph.
- Ritual headings: 16 pt bold, centered, kept with following paragraph.
- List items: justified, 0.5 cm indent, marker included in text.

## Closing and signature

- Closing is right aligned.
- Signature block is right aligned.
- Names of signatories are bold.
- Spacing around signature lines is intentional and should be approved.

## Attachments and notes

- `Allegati:` label is bold.
- Attachment items are indented.
- Postscript/disclaimer is small and italic.

## Open approval language

The artifact should not imply final client approval. Use language such as:

- "Regola attuale"
- "Da approvare"
- "Punto sensibile"
- "Ereditato dal template"
- "Generato dallo script"

